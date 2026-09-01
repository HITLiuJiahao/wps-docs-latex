import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

type CandidateKind = 'marked' | 'block' | 'command' | 'bare-script';

type FormulaCandidate = {
  start: number;
  end: number;
  raw: string;
  latex: string;
  kind: CandidateKind;
  confidence: 'high' | 'medium';
  display: boolean;
};

type TextSegment = {
  node: Element;
  start: number;
  end: number;
  text: string;
  run: Element | null;
};

type MathNode =
  | { type: 'text'; value: string }
  | { type: 'group'; body: MathNode[] }
  | { type: 'fraction'; numerator: MathNode[]; denominator: MathNode[] }
  | { type: 'radical'; body: MathNode[]; degree?: MathNode[] }
  | { type: 'accent'; accent: string; body: MathNode[] }
  | { type: 'script'; base: MathNode; sub?: MathNode[]; sup?: MathNode[] }
  | { type: 'nary'; symbol: string; sub?: MathNode[]; sup?: MathNode[]; body?: MathNode[] }
  | { type: 'matrix'; rows: MathNode[][][]; cases: boolean };

type ParsedFormula = {
  nodes: MathNode[];
  warnings: string[];
};

function hasNaryLimitMarker(nodes: MathNode[]): boolean {
  return nodes.some((node) => {
    if (node.type === 'text') return /[=∈∉≠≤≥]/.test(node.value);
    if (node.type === 'group') return hasNaryLimitMarker(node.body);
    if (node.type === 'script') return hasNaryLimitMarker([node.base, ...(node.sub ?? []), ...(node.sup ?? [])]);
    if (node.type === 'accent') return hasNaryLimitMarker(node.body);
    return false;
  });
}

export type FormulaRecord = {
  id: string;
  part: string;
  paragraph: number;
  source: string;
  normalized: string;
  confidence: '高' | '中';
  status: '已转换' | '需复核';
  formulaType: '行内公式' | '独立公式';
  warnings: string[];
  reason?: string;
};

export type ConversionReport = {
  version: 1;
  sourceFile: string;
  generatedAt: string;
  processing: '仅在本地浏览器完成';
  scannedParts: string[];
  scannedParagraphs: number;
  candidates: number;
  convertedText: number;
  nativeFormulaObjects: number;
  needsReview: number;
  residualCandidates: number;
  records: FormulaRecord[];
};

export type ConversionResult = {
  blob: Blob;
  outputName: string;
  report: ConversionReport;
};

const greek: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ϵ',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
  sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ',
  psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ',
  Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

const symbols: Record<string, string> = {
  cdot: '·', times: '×', div: '÷', pm: '±', mp: '∓', leq: '≤', geq: '≥',
  neq: '≠', approx: '≈', sim: '∼', equiv: '≡', propto: '∝', infty: '∞',
  partial: '∂', nabla: '∇', to: '→', rightarrow: '→', leftarrow: '←',
  leftrightarrow: '↔', Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔',
  in: '∈', notin: '∉', subset: '⊂', subseteq: '⊆', superset: '⊃',
  supseteq: '⊇', cup: '∪', cap: '∩', land: '∧', lor: '∨', forall: '∀',
  exists: '∃', neg: '¬', degree: '°', ell: 'ℓ', ldots: '…', cdots: '⋯',
  vert: '|', mid: '|', lvert: '|', rvert: '|', langle: '〈', rangle: '〉',
};

const narySymbols: Record<string, string> = {
  sum: '∑', prod: '∏', int: '∫', iint: '∬', iiint: '∭', oint: '∮',
};

const unicodeNarySymbols: Record<string, string> = {
  '∑': '∑', '∏': '∏', '∫': '∫', '∬': '∬', '∭': '∭', '∮': '∮',
};

const structuralCommands = new Set([
  'frac', 'dfrac', 'tfrac', 'cfrac', 'sqrt', 'sum', 'prod', 'int', 'iint', 'iiint', 'oint',
  'hat', 'widehat', 'tilde', 'widetilde', 'bar', 'overline', 'underline', 'vec', 'dot', 'ddot',
  'begin', 'left', 'right', 'bigl', 'bigr', 'Bigl', 'Bigr', 'biggl', 'biggr', 'Biggl', 'Biggr',
]);

const groupingCommands = new Set([
  'text', 'textrm', 'textbf', 'mathrm', 'mathbf', 'mathit', 'mathnormal', 'mathsf', 'mathtt',
  'mathcal', 'mathbb', 'mathfrak', 'operatorname', 'operatornamewithlimits', 'mathop', 'phantom',
]);

const namedOperators = new Set([
  'lim', 'limsup', 'liminf', 'max', 'min', 'sup', 'inf', 'argmax', 'argmin',
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'exp', 'det', 'gcd', 'Pr',
]);

const ignoredCommands = new Set([
  'displaystyle', 'textstyle', 'scriptstyle', 'scriptscriptstyle', 'limits', 'nolimits',
]);

function isCjk(value: string) {
  return /[\u3400-\u9fff]/.test(value);
}

function isIdentifierStart(value: string | undefined) {
  return Boolean(value && /[A-Za-zα-ωΑ-Ω]/.test(value));
}

function isIdentifierCharacter(value: string | undefined) {
  return Boolean(value && /[A-Za-z0-9α-ωΑ-Ω\u0300-\u036f]/.test(value));
}

function stripMarkers(raw: string) {
  if (raw.startsWith('$$') && raw.endsWith('$$')) return raw.slice(2, -2).trim();
  if (raw.startsWith('\\[') && raw.endsWith('\\]')) return raw.slice(2, -2).trim();
  if (raw.startsWith('\\(') && raw.endsWith('\\)')) return raw.slice(2, -2).trim();
  if (raw.startsWith('$') && raw.endsWith('$')) return raw.slice(1, -1).trim();
  return raw.trim();
}

function looksLikeMath(value: string) {
  return /\\[A-Za-z]+|[_^{}]|[=+\-*/<>≤≥≈∑∫√]|[α-ωΑ-Ω]/.test(value);
}

function addCandidate(target: FormulaCandidate[], candidate: FormulaCandidate) {
  if (candidate.end <= candidate.start || !candidate.latex || !looksLikeMath(candidate.latex)) return;
  const overlaps = target.some((item) => candidate.start < item.end && candidate.end > item.start);
  if (!overlaps) target.push(candidate);
}

function commandExtent(text: string, start: number) {
  let cursor = start;
  let depth = 0;
  let spaces = 0;
  while (cursor < text.length && cursor - start < 180) {
    const char = text[cursor];
    if (char === '{' || char === '[' || char === '(') depth += 1;
    if (char === '}' || char === ']' || char === ')') depth = Math.max(0, depth - 1);
    if (depth === 0) {
      if (isCjk(char) || /[。；;！!?]/.test(char) || char === '\n' || char === '\r') break;
      if (char === ' ') {
        spaces += 1;
        if (spaces >= 2) break;
      } else {
        spaces = 0;
      }
    }
    cursor += 1;
  }
  return cursor;
}

function balancedEnd(text: string, start: number, open = '{', close = '}') {
  if (text[start] !== open) return -1;
  let depth = 0;
  for (let cursor = start; cursor < text.length; cursor += 1) {
    if (text[cursor] === open) depth += 1;
    if (text[cursor] === close) {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
  }
  return -1;
}

function trimFormulaRange(text: string, start: number, end: number) {
  while (start < end && /\s/.test(text[start])) start += 1;
  while (end > start && /\s/.test(text[end - 1])) end -= 1;
  while (end > start && /[。；;，,]/.test(text[end - 1])) end -= 1;
  while (end > start && /\s/.test(text[end - 1])) end -= 1;
  if (text[start] === '（' && text[end - 1] === '）') {
    start += 1;
    end -= 1;
  }
  return { start, end };
}

function isLikelyFormulaBlock(value: string) {
  if (!value || value.length > 420 || /^（\d+[）)]/.test(value)) return false;
  const first = value[0];
  const startsAsMath = isIdentifierStart(first) || first === '\\' || first === '(' || first === '[' || first === '{' || Boolean(unicodeNarySymbols[first]);
  if (!startsAsMath) return false;

  const cjkCount = (value.match(/[\u3400-\u9fff]/g) ?? []).length;
  const scriptCount = (value.match(/[_^]/g) ?? []).length;
  const operatorCount = (value.match(/[=~≈≃<>≤≥∈∉]/g) ?? []).length;
  const structuralCount = (value.match(/[\\∑∏∫∬∭∮{}|]/g) ?? []).length;
  const mathHeavy = scriptCount + operatorCount * 2 + structuralCount;
  const cjkAllowance = Math.max(10, Math.floor(value.length * 0.2));

  if (cjkCount > cjkAllowance) return false;
  return operatorCount > 0 || (scriptCount > 1 && structuralCount > 0) || structuralCount >= 3;
}

function addStandaloneFormulaBlock(text: string, target: FormulaCandidate[]) {
  const range = trimFormulaRange(text, 0, text.length);
  const raw = text.slice(range.start, range.end);
  if (!isLikelyFormulaBlock(raw)) return;
  addCandidate(target, {
    start: range.start,
    end: range.end,
    raw,
    latex: raw,
    kind: 'block',
    confidence: 'high',
    display: true,
  });
}

function addBareScriptCandidates(text: string, target: FormulaCandidate[]) {
  let cursor = 0;
  while (cursor < text.length) {
    const char = text[cursor];
    if (!isIdentifierStart(char) || isIdentifierCharacter(text[cursor - 1])) {
      cursor += 1;
      continue;
    }

    const start = cursor;
    cursor += 1;
    while (isIdentifierCharacter(text[cursor])) cursor += 1;
    let end = cursor;
    let hasScript = false;

    while (text[end] === '_' || text[end] === '^') {
      const markerStart = end;
      end += 1;
      if (text[end] === '{') {
        const closing = balancedEnd(text, end);
        if (closing === -1) {
          end = markerStart;
          break;
        }
        end = closing;
      } else if (isIdentifierCharacter(text[end]) || /[+\-−]/.test(text[end] ?? '')) {
        end += 1;
        while (isIdentifierCharacter(text[end])) end += 1;
      } else {
        end = markerStart;
        break;
      }
      hasScript = true;
    }

    if (hasScript) {
      const raw = text.slice(start, end);
      addCandidate(target, {
        start,
        end,
        raw,
        latex: raw,
        kind: 'bare-script',
        confidence: 'medium',
        display: false,
      });
      cursor = end;
    }
  }
}

function detectFormulaCandidates(text: string) {
  const results: FormulaCandidate[] = [];
  const marked = /\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\)|\$\$[\s\S]*?\$\$|\$[^\n$]+?\$/g;
  let match: RegExpExecArray | null;

  while ((match = marked.exec(text))) {
    const raw = match[0];
    addCandidate(results, {
      start: match.index,
      end: match.index + raw.length,
      raw,
      latex: stripMarkers(raw),
      kind: 'marked',
      confidence: 'high',
      display: raw.startsWith('$$') || raw.startsWith('\\['),
    });
  }

  addStandaloneFormulaBlock(text, results);

  const command = /\\([A-Za-z]+)/g;
  while ((match = command.exec(text))) {
    const commandName = match[1];
    const next = text[match.index + match[0].length] ?? '';
    const canStartFormula = structuralCommands.has(commandName)
      || groupingCommands.has(commandName)
      || namedOperators.has(commandName)
      || Boolean(greek[commandName])
      || Boolean(symbols[commandName])
      || Boolean(narySymbols[commandName])
      || next === '{' || next === '_' || next === '^' || next === '[' || next === '(';
    if (!canStartFormula) continue;
    const start = match.index;
    const range = trimFormulaRange(text, start, commandExtent(text, start));
    const raw = text.slice(range.start, range.end);
    addCandidate(results, {
      start: range.start,
      end: range.end,
      raw,
      latex: raw,
      kind: 'command',
      confidence: 'high',
      display: /\\(?:begin\{(?:equation|align|cases|matrix|pmatrix|bmatrix|array)\}|sum|prod|int)/.test(raw),
    });
  }

  addBareScriptCandidates(text, results);
  return results.sort((a, b) => a.start - b.start);
}

class LatexParser {
  private cursor = 0;
  private warnings: string[] = [];

  constructor(private readonly source: string) {}

  parse(): ParsedFormula {
    return { nodes: this.parseSequence(), warnings: this.warnings };
  }

  private parseSequence(end?: string): MathNode[] {
    const nodes: MathNode[] = [];
    while (this.cursor < this.source.length) {
      if (end && this.source[this.cursor] === end) {
        this.cursor += 1;
        break;
      }
      if (this.source[this.cursor] === '}' && !end) {
        this.warnings.push('检测到未配对的右花括号');
        this.cursor += 1;
        continue;
      }
      if (/\s/.test(this.source[this.cursor])) {
        this.cursor += 1;
        continue;
      }
      let node = this.parseAtom();
      if (!node) continue;
      let sub: MathNode[] | undefined;
      let sup: MathNode[] | undefined;
      if (node.type === 'nary' && this.source[this.cursor] === '{') {
        const savedCursor = this.cursor;
        const directLowerLimit = this.parseArgument();
        if (this.source[this.cursor] === '_' || this.source[this.cursor] === '^' || hasNaryLimitMarker(directLowerLimit)) {
          sub = directLowerLimit;
        }
        else this.cursor = savedCursor;
      }
      while (this.source[this.cursor] === '_' || this.source[this.cursor] === '^') {
        const marker = this.source[this.cursor];
        this.cursor += 1;
        const script = this.parseArgument();
        if (marker === '_') sub = script;
        if (marker === '^') sup = script;
      }
      // In OMML a n-ary operator owns its operand in an <m:e> element.  Leaving
      // that element empty makes Word/WPS display a square placeholder.  Treat
      // the remaining expression in this group as the operator's operand, so
      // both the formula structure and its rendered appearance stay intact.
      if (node.type === 'nary') {
        node = { ...node, sub, sup, body: this.parseSequence(end) };
        nodes.push(node);
        return nodes;
      }
      if (sub || sup) {
        if (node.type === 'script') {
          node = { ...node, sub: sub ?? node.sub, sup: sup ?? node.sup };
        } else {
          node = { type: 'script', base: node, sub, sup };
        }
      }
      nodes.push(node);
    }
    if (end && this.source[this.cursor - 1] !== end) this.warnings.push('检测到未闭合的花括号');
    return nodes;
  }

  private parseArgument() {
    this.skipSpaces();
    if (this.source[this.cursor] === '{') {
      this.cursor += 1;
      return this.parseSequence('}');
    }
    const atom = this.parseAtom();
    return atom ? [atom] : [];
  }

  private parseAtom(): MathNode | null {
    const char = this.source[this.cursor];
    if (!char) return null;
    if (char === '{') {
      this.cursor += 1;
      return { type: 'group', body: this.parseSequence('}') };
    }
    if (char === '\\') return this.parseCommand();
    this.cursor += 1;
    if (unicodeNarySymbols[char]) return { type: 'nary', symbol: unicodeNarySymbols[char] };
    if (/[A-Za-z0-9]/.test(char)) {
      let value = char;
      while (this.cursor < this.source.length && /[A-Za-z0-9]/.test(this.source[this.cursor])) {
        value += this.source[this.cursor];
        this.cursor += 1;
      }
      return { type: 'text', value };
    }
    if (char === '*') return { type: 'text', value: '·' };
    if (char === '~') return { type: 'text', value: '∼' };
    return { type: 'text', value: char };
  }

  private parseCommand(): MathNode | null {
    this.cursor += 1;
    const command = this.readCommandName();
    if (!command) return { type: 'text', value: '\\' };
    if (command === '\\') return { type: 'text', value: ' ' };
    if (command === ',' || command === ';' || command === '!' || command === ' ') return null;
    if (command === '{') return { type: 'text', value: '{' };
    if (command === '}') return { type: 'text', value: '}' };
    if (command === '_') return { type: 'text', value: '_' };
    if (command === '^') return { type: 'text', value: '^' };
    if (greek[command]) return { type: 'text', value: greek[command] };
    if (symbols[command]) return { type: 'text', value: symbols[command] };
    if (narySymbols[command]) return { type: 'nary', symbol: narySymbols[command] };

    if (command === 'frac' || command === 'dfrac' || command === 'tfrac' || command === 'cfrac') {
      return { type: 'fraction', numerator: this.parseArgument(), denominator: this.parseArgument() };
    }
    if (command === 'sqrt') {
      this.skipSpaces();
      const degree = this.source[this.cursor] === '[' ? this.parseBracketArgument() : undefined;
      return { type: 'radical', body: this.parseArgument(), degree };
    }
    if (['hat', 'widehat', 'tilde', 'widetilde', 'bar', 'overline', 'underline', 'vec', 'dot', 'ddot'].includes(command)) {
      const accents: Record<string, string> = {
        hat: 'ˆ', widehat: 'ˆ', tilde: '˜', widetilde: '˜', bar: '¯',
        overline: '¯', underline: '_', vec: '⃗', dot: '˙', ddot: '¨',
      };
      return { type: 'accent', accent: accents[command] ?? 'ˆ', body: this.parseArgument() };
    }
    if (groupingCommands.has(command)) {
      if ((command === 'operatorname' || command === 'operatornamewithlimits') && this.source[this.cursor] === '*') this.cursor += 1;
      return { type: 'group', body: this.parseArgument() };
    }
    if (namedOperators.has(command)) return { type: 'text', value: command };
    if (ignoredCommands.has(command)) return null;
    if (command === 'quad' || command === 'qquad' || command === 'enspace') return { type: 'text', value: ' ' };
    if (command === 'left' || command === 'right' || ['bigl', 'bigr', 'Bigl', 'Bigr', 'biggl', 'biggr', 'Biggl', 'Biggr'].includes(command)) {
      this.skipSpaces();
      const delimiter = this.parseDelimiter();
      return delimiter ? { type: 'text', value: delimiter } : null;
    }
    if (command === 'begin') return this.parseEnvironment();
    if (command === 'end') {
      this.readGroupName();
      this.warnings.push('检测到不在环境中的 \\end 命令');
      return null;
    }
    this.warnings.push('暂不识别命令 \\' + command);
    return { type: 'text', value: '\\' + command };
  }

  private parseEnvironment(): MathNode {
    const environment = this.readGroupName();
    if (!environment) {
      this.warnings.push('无法读取公式环境名称');
      return { type: 'text', value: 'begin' };
    }
    const closing = '\\end{' + environment + '}';
    const end = this.source.indexOf(closing, this.cursor);
    if (end === -1) {
      this.warnings.push('环境 ' + environment + ' 没有匹配的 \\end');
      return { type: 'text', value: environment };
    }
    const content = this.source.slice(this.cursor, end);
    this.cursor = end + closing.length;
    const rows = content.split(/\\\\/).map((row) => row.split('&').map((cell) => {
      const nested = new LatexParser(cell).parse();
      this.warnings.push(...nested.warnings);
      return nested.nodes;
    }));
    return { type: 'matrix', rows, cases: environment === 'cases' };
  }

  private parseDelimiter() {
    if (this.source[this.cursor] === '\\') {
      this.cursor += 1;
      const name = this.readCommandName();
      return symbols[name] ?? ({ lbrace: '{', rbrace: '}', lVert: '‖', rVert: '‖', Vert: '‖', '|': '|', '.': '' }[name] ?? name);
    }
    const delimiter = this.source[this.cursor] ?? '';
    this.cursor += 1;
    return delimiter === '.' ? '' : delimiter;
  }

  private readCommandName() {
    if (/[A-Za-z]/.test(this.source[this.cursor] ?? '')) {
      let name = '';
      while (/[A-Za-z]/.test(this.source[this.cursor] ?? '')) {
        name += this.source[this.cursor];
        this.cursor += 1;
      }
      return name;
    }
    const next = this.source[this.cursor] ?? '';
    this.cursor += 1;
    return next;
  }

  private readGroupName() {
    this.skipSpaces();
    if (this.source[this.cursor] !== '{') return '';
    const start = ++this.cursor;
    while (this.cursor < this.source.length && this.source[this.cursor] !== '}') this.cursor += 1;
    const value = this.source.slice(start, this.cursor).trim();
    if (this.source[this.cursor] === '}') this.cursor += 1;
    return value;
  }

  private parseBracketArgument() {
    if (this.source[this.cursor] !== '[') return [];
    const start = ++this.cursor;
    let depth = 1;
    while (this.cursor < this.source.length && depth > 0) {
      const char = this.source[this.cursor++];
      if (char === '[') depth += 1;
      if (char === ']') depth -= 1;
    }
    if (depth !== 0) {
      this.warnings.push('检测到未闭合的方括号');
      return [];
    }
    const parsed = new LatexParser(this.source.slice(start, this.cursor - 1)).parse();
    this.warnings.push(...parsed.warnings);
    return parsed.nodes;
  }

  private skipBalanced(open: string, close: string) {
    if (this.source[this.cursor] !== open) return;
    let depth = 0;
    while (this.cursor < this.source.length) {
      const char = this.source[this.cursor++];
      if (char === open) depth += 1;
      if (char === close) {
        depth -= 1;
        if (depth === 0) break;
      }
    }
  }

  private skipSpaces() {
    while (/\s/.test(this.source[this.cursor] ?? '')) this.cursor += 1;
  }
}

function mathElement(document: XMLDocument, name: string) {
  return document.createElementNS(M_NS, 'm:' + name);
}

function setMathAttribute(element: Element, name: string, value: string) {
  element.setAttributeNS(M_NS, 'm:' + name, value);
}

function appendText(document: XMLDocument, parent: Element, value: string) {
  if (!value) return;
  const run = mathElement(document, 'r');
  const text = mathElement(document, 't');
  text.textContent = value;
  run.appendChild(text);
  parent.appendChild(run);
}

function appendNodes(document: XMLDocument, parent: Element, nodes: MathNode[]) {
  for (const node of nodes) appendNode(document, parent, node);
}

function appendNary(
  document: XMLDocument,
  parent: Element,
  symbol: string,
  sub?: MathNode[],
  sup?: MathNode[],
  body?: MathNode[],
) {
  const nary = mathElement(document, 'nary');
  const properties = mathElement(document, 'naryPr');
  const char = mathElement(document, 'chr');
  setMathAttribute(char, 'val', symbol);
  properties.appendChild(char);
  if (!sub?.length) {
    const hideSubscript = mathElement(document, 'subHide');
    setMathAttribute(hideSubscript, 'val', '1');
    properties.appendChild(hideSubscript);
  }
  if (!sup?.length) {
    const hideSuperscript = mathElement(document, 'supHide');
    setMathAttribute(hideSuperscript, 'val', '1');
    properties.appendChild(hideSuperscript);
  }
  nary.appendChild(properties);
  const subContainer = mathElement(document, 'sub');
  if (sub) appendNodes(document, subContainer, sub);
  nary.appendChild(subContainer);
  const supContainer = mathElement(document, 'sup');
  if (sup) appendNodes(document, supContainer, sup);
  nary.appendChild(supContainer);
  const bodyContainer = mathElement(document, 'e');
  if (body) appendNodes(document, bodyContainer, body);
  nary.appendChild(bodyContainer);
  parent.appendChild(nary);
}

function appendNode(document: XMLDocument, parent: Element, node: MathNode) {
  if (node.type === 'text') {
    appendText(document, parent, node.value);
    return;
  }
  if (node.type === 'group') {
    appendNodes(document, parent, node.body);
    return;
  }
  if (node.type === 'fraction') {
    const fraction = mathElement(document, 'f');
    const numerator = mathElement(document, 'num');
    const denominator = mathElement(document, 'den');
    appendNodes(document, numerator, node.numerator);
    appendNodes(document, denominator, node.denominator);
    fraction.appendChild(numerator);
    fraction.appendChild(denominator);
    parent.appendChild(fraction);
    return;
  }
  if (node.type === 'radical') {
    const radical = mathElement(document, 'rad');
    const properties = mathElement(document, 'radPr');
    const degree = mathElement(document, 'deg');
    if (node.degree?.length) appendNodes(document, degree, node.degree);
    else {
      const hideDegree = mathElement(document, 'degHide');
      setMathAttribute(hideDegree, 'val', '1');
      properties.appendChild(hideDegree);
    }
    const body = mathElement(document, 'e');
    appendNodes(document, body, node.body);
    radical.appendChild(properties);
    radical.appendChild(degree);
    radical.appendChild(body);
    parent.appendChild(radical);
    return;
  }
  if (node.type === 'accent') {
    const accent = mathElement(document, 'acc');
    const properties = mathElement(document, 'accPr');
    const character = mathElement(document, 'chr');
    setMathAttribute(character, 'val', node.accent);
    properties.appendChild(character);
    const body = mathElement(document, 'e');
    appendNodes(document, body, node.body);
    accent.appendChild(properties);
    accent.appendChild(body);
    parent.appendChild(accent);
    return;
  }
  if (node.type === 'nary') {
    appendNary(document, parent, node.symbol, node.sub, node.sup, node.body);
    return;
  }
  if (node.type === 'script') {
    if (node.base.type === 'nary') {
      appendNary(document, parent, node.base.symbol, node.sub, node.sup, node.base.body);
      return;
    }
    const kind = node.sub && node.sup ? 'sSubSup' : node.sub ? 'sSub' : 'sSup';
    const script = mathElement(document, kind);
    const base = mathElement(document, 'e');
    appendNode(document, base, node.base);
    script.appendChild(base);
    if (node.sub) {
      const sub = mathElement(document, 'sub');
      appendNodes(document, sub, node.sub);
      script.appendChild(sub);
    }
    if (node.sup) {
      const sup = mathElement(document, 'sup');
      appendNodes(document, sup, node.sup);
      script.appendChild(sup);
    }
    parent.appendChild(script);
    return;
  }
  if (node.type === 'matrix') {
    const matrix = mathElement(document, 'm');
    for (const row of node.rows) {
      const matrixRow = mathElement(document, 'mr');
      for (const cellNodes of row) {
        const cell = mathElement(document, 'e');
        appendNodes(document, cell, cellNodes);
        matrixRow.appendChild(cell);
      }
      matrix.appendChild(matrixRow);
    }
    if (!node.cases) {
      parent.appendChild(matrix);
      return;
    }
    const delimiter = mathElement(document, 'd');
    const properties = mathElement(document, 'dPr');
    const begin = mathElement(document, 'begChr');
    const end = mathElement(document, 'endChr');
    setMathAttribute(begin, 'val', '{');
    setMathAttribute(end, 'val', '');
    properties.appendChild(begin);
    properties.appendChild(end);
    const body = mathElement(document, 'e');
    body.appendChild(matrix);
    delimiter.appendChild(properties);
    delimiter.appendChild(body);
    parent.appendChild(delimiter);
  }
}

function makeOfficeMath(document: XMLDocument, parsed: ParsedFormula) {
  const math = mathElement(document, 'oMath');
  appendNodes(document, math, parsed.nodes);
  return math;
}

function makeLineBreak(document: XMLDocument) {
  const run = document.createElementNS(W_NS, 'w:r');
  run.appendChild(document.createElementNS(W_NS, 'w:br'));
  return run;
}

function keepParagraphTogether(document: XMLDocument, paragraph: Element) {
  let properties = Array.from(paragraph.children).find(
    (child) => child.namespaceURI === W_NS && child.localName === 'pPr',
  ) as Element | undefined;
  if (!properties) {
    properties = document.createElementNS(W_NS, 'w:pPr');
    paragraph.insertBefore(properties, paragraph.firstChild);
  }
  const alreadyKept = Array.from(properties.children).some(
    (child) => child.namespaceURI === W_NS && child.localName === 'keepLines',
  );
  if (!alreadyKept) {
    const keepLines = document.createElementNS(W_NS, 'w:keepLines');
    keepLines.setAttributeNS(W_NS, 'w:val', '1');
    properties.appendChild(keepLines);
  }
}

function splitLongDisplayFormula(value: string): [string, string] | null {
  // Native Office Math objects cannot wrap inside a Word line. Split only an
  // unusually long display formula at a mathematical operator so the complete
  // expression remains visible as two native formula objects.
  if (value.length < 130) return null;
  const lower = Math.floor(value.length * 0.45);
  const upper = Math.floor(value.length * 0.84);
  const candidates = [...value].flatMap((character, index) => (
    index >= lower && index <= upper && (character === '+' || character === '−') ? [index] : []
  ));
  if (!candidates.length) return null;
  const target = Math.floor(value.length * 0.66);
  const splitAt = candidates.reduce((best, index) => (
    Math.abs(index - target) < Math.abs(best - target) ? index : best
  ));
  return [value.slice(0, splitAt), value.slice(splitAt)];
}

function collectSegments(paragraph: Element) {
  const textNodes = Array.from(paragraph.getElementsByTagNameNS(W_NS, 't')) as Element[];
  const segments: TextSegment[] = [];
  let offset = 0;
  for (const node of textNodes) {
    const text = node.textContent ?? '';
    segments.push({ node, start: offset, end: offset + text.length, text, run: findAncestorRun(node) });
    offset += text.length;
  }
  return segments;
}

function findAncestorRun(node: Element) {
  let current: Element | null = node.parentElement;
  while (current) {
    if (current.namespaceURI === W_NS && current.localName === 'r') return current;
    current = current.parentElement;
  }
  return null;
}

function setWordText(node: Element, value: string) {
  node.textContent = value;
  if (/^\s|\s$/.test(value)) node.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  else node.removeAttributeNS(XML_NS, 'space');
}

function cloneRunWithText(document: XMLDocument, sourceRun: Element, value: string) {
  const run = document.createElementNS(W_NS, 'w:r');
  const properties = Array.from(sourceRun.children).find((child) => child.namespaceURI === W_NS && child.localName === 'rPr');
  if (properties) run.appendChild(properties.cloneNode(true));
  const text = document.createElementNS(W_NS, 'w:t');
  setWordText(text, value);
  run.appendChild(text);
  return run;
}

function locateSegment(segments: TextSegment[], start: number, end: number) {
  const first = segments.find((segment) => start >= segment.start && start < segment.end);
  const last = [...segments].reverse().find((segment) => end > segment.start && end <= segment.end);
  if (!first || !last || !first.run || !last.run) return null;
  return { first, last };
}

function replaceCandidate(
  document: XMLDocument,
  paragraph: Element,
  segments: TextSegment[],
  candidate: FormulaCandidate,
  insertions: Element[],
) {
  const location = locateSegment(segments, candidate.start, candidate.end);
  if (!location) return false;
  const { first, last } = location;
  const firstRun = first.run;
  const lastRun = last.run;
  if (!firstRun || !lastRun) return false;
  const firstStart = candidate.start - first.start;
  const lastEnd = candidate.end - last.start;

  if (first.node === last.node) {
    const before = (first.node.textContent ?? '').slice(0, firstStart);
    const after = (first.node.textContent ?? '').slice(lastEnd);
    setWordText(first.node, before);
    const insertionPoint = before ? firstRun.nextSibling : firstRun;
    for (const insertion of insertions) paragraph.insertBefore(insertion, insertionPoint);
    if (after) paragraph.insertBefore(cloneRunWithText(document, firstRun, after), insertionPoint);
    return true;
  }

  const before = (first.node.textContent ?? '').slice(0, firstStart);
  const after = (last.node.textContent ?? '').slice(lastEnd);
  setWordText(first.node, before);
  setWordText(last.node, after);
  for (const segment of segments) {
    if (segment.start > first.start && segment.end <= last.end && segment.node !== last.node) setWordText(segment.node, '');
  }
  const insertionPoint = before ? firstRun.nextSibling : firstRun;
  for (const insertion of insertions) paragraph.insertBefore(insertion, insertionPoint);
  return true;
}

function friendlyPartName(part: string) {
  if (part === 'word/document.xml') return '正文';
  if (part.includes('header')) return '页眉（' + (part.match(/\d+/)?.[0] ?? '1') + '）';
  if (part.includes('footer')) return '页脚（' + (part.match(/\d+/)?.[0] ?? '1') + '）';
  if (part.includes('footnotes')) return '脚注';
  if (part.includes('endnotes')) return '尾注';
  return part;
}

function wordParts(zip: JSZip) {
  return Object.keys(zip.files).filter((name) => /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(name));
}

function xmlDocument(value: string, part: string) {
  const document = new DOMParser().parseFromString(value, 'application/xml');
  if (document.getElementsByTagName('parsererror').length) throw new Error(friendlyPartName(part) + ' 的内部结构无法读取');
  return document;
}

function withoutExtension(name: string) {
  return name.toLowerCase().endsWith('.docx') ? name.slice(0, -5) : name;
}

async function countValidation(zip: JSZip) {
  let formulaObjects = 0;
  let residualCandidates = 0;
  for (const part of wordParts(zip)) {
    const xml = await zip.file(part)?.async('string');
    if (!xml) continue;
    const document = xmlDocument(xml, part);
    formulaObjects += document.getElementsByTagNameNS(M_NS, 'oMath').length;
    const paragraphs = Array.from(document.getElementsByTagNameNS(W_NS, 'p')) as Element[];
    for (const paragraph of paragraphs) {
      const text = collectSegments(paragraph).map((segment) => segment.text).join('');
      residualCandidates += detectFormulaCandidates(text).length;
    }
  }
  return { formulaObjects, residualCandidates };
}

export async function convertDocx(file: File): Promise<ConversionResult> {
  if (!file.name.toLowerCase().endsWith('.docx')) throw new Error('请选择 .docx 格式的文档');
  const zip = await JSZip.loadAsync(file);
  if (!zip.file('word/document.xml')) throw new Error('这不是可识别的 Word/WPS DOCX 文档');

  const parts = wordParts(zip);
  const records: FormulaRecord[] = [];
  let paragraphsScanned = 0;
  let converted = 0;
  let review = 0;
  let recordNumber = 0;

  for (const part of parts) {
    const entry = zip.file(part);
    if (!entry) continue;
    const xml = await entry.async('string');
    const document = xmlDocument(xml, part);
    const paragraphs = Array.from(document.getElementsByTagNameNS(W_NS, 'p')) as Element[];

    for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
      const paragraph = paragraphs[paragraphIndex];
      const segments = collectSegments(paragraph);
      const text = segments.map((segment) => segment.text).join('');
      paragraphsScanned += 1;
      const candidates = detectFormulaCandidates(text);

      for (const candidate of [...candidates].reverse()) {
        const id = 'F' + String(++recordNumber).padStart(4, '0');
        const parsed = new LatexParser(candidate.latex).parse();
        const warnings = [...new Set(parsed.warnings)];
        const formulaType = candidate.display ? '独立公式' : '行内公式';
        if (!parsed.nodes.length) {
          review += 1;
          records.push({
            id, part: friendlyPartName(part), paragraph: paragraphIndex + 1,
            source: candidate.raw, normalized: candidate.latex,
            confidence: candidate.confidence === 'high' ? '高' : '中',
            status: '需复核', formulaType, warnings, reason: '未能解析出可转换的公式结构',
          });
          continue;
        }
        let insertions = [makeOfficeMath(document, parsed)];
        const splitFormula = candidate.display ? splitLongDisplayFormula(candidate.latex) : null;
        if (splitFormula) {
          const first = new LatexParser(splitFormula[0]).parse();
          const second = new LatexParser(splitFormula[1]).parse();
          if (first.nodes.length && second.nodes.length) {
            for (const warning of [...first.warnings, ...second.warnings]) {
              if (!warnings.includes(warning)) warnings.push(warning);
            }
            insertions = [makeOfficeMath(document, first), makeLineBreak(document), makeOfficeMath(document, second)];
            keepParagraphTogether(document, paragraph);
          }
        }
        const replaced = replaceCandidate(document, paragraph, segments, candidate, insertions);
        if (!replaced) {
          review += 1;
          records.push({
            id, part: friendlyPartName(part), paragraph: paragraphIndex + 1,
            source: candidate.raw, normalized: candidate.latex,
            confidence: candidate.confidence === 'high' ? '高' : '中',
            status: '需复核', formulaType, warnings, reason: '公式跨越了不支持的文档结构，已保留原文',
          });
          continue;
        }
        converted += 1;
        records.push({
          id, part: friendlyPartName(part), paragraph: paragraphIndex + 1,
          source: candidate.raw, normalized: candidate.latex,
          confidence: candidate.confidence === 'high' ? '高' : '中',
          status: '已转换', formulaType, warnings,
        });
      }
    }
    zip.file(part, new XMLSerializer().serializeToString(document));
  }

  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const validation = await countValidation(await JSZip.loadAsync(blob));
  const report: ConversionReport = {
    version: 1,
    sourceFile: file.name,
    generatedAt: new Date().toLocaleString('zh-CN', { hour12: false }),
    processing: '仅在本地浏览器完成',
    scannedParts: parts.map(friendlyPartName),
    scannedParagraphs: paragraphsScanned,
    candidates: records.length,
    convertedText: converted,
    nativeFormulaObjects: validation.formulaObjects,
    needsReview: review,
    residualCandidates: validation.residualCandidates,
    records: records.reverse(),
  };
  return {
    blob,
    outputName: withoutExtension(file.name) + '（LaTeX公式已转换）.docx',
    report,
  };
}

// Exposed only for local regression checks; the app itself calls convertDocx above.
export const converterTesting = { detectFormulaCandidates, LatexParser, splitLongDisplayFormula };
