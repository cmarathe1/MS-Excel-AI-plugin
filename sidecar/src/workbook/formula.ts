import { colToIndex, type CellScalar, type RangeAddress } from '@excelai/shared';

/**
 * A small, deterministic evaluator for canonical en-US Excel formulas.
 *
 * Used by the headless emulator for tests, the capability-probe sandbox and
 * post-action verification. It intentionally implements a core function set
 * exactly rather than all of Excel approximately; unknown functions evaluate
 * to #NAME? so gaps are loud, never silently wrong. (HyperFormula was
 * considered and rejected for now: GPLv3 — incompatible with this repo's MIT
 * distribution.)
 */

export type EvalError = '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#CYCLE!' | '#N/A';

export class FormulaError extends Error {
  constructor(public code: EvalError) {
    super(code);
  }
}

export interface CellResolver {
  /** Resolve a single cell's computed value. */
  getCell(sheet: string, row: number, col: number): CellScalar;
  /** Resolve a range reference into a 2-D array of computed values. */
  getRange(addr: RangeAddress): CellScalar[][];
  sheetExists(sheet: string): boolean;
}

/* ------------------------- tokenizer ------------------------- */

type Token =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'ident'; v: string } // function name or sheet-less ref, resolved later
  | { t: 'ref'; v: string } // ref with sheet prefix (quoted or not)
  | { t: 'op'; v: string }
  | { t: 'lparen' }
  | { t: 'rparen' }
  | { t: 'comma' };

const IDENT_START = /[A-Za-z_]/;
const IDENT_CHAR = /[A-Za-z0-9_.$]/;

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const s = src;
  while (i < s.length) {
    const c = s[i]!;
    if (c === ' ' || c === '\t' || c === '\n') {
      i++;
      continue;
    }
    if (c >= '0' && c <= '9') {
      let j = i;
      while (j < s.length && /[0-9.]/.test(s[j]!)) j++;
      // scientific notation
      if (s[j] === 'e' || s[j] === 'E') {
        let k = j + 1;
        if (s[k] === '+' || s[k] === '-') k++;
        if (s[k] && /[0-9]/.test(s[k]!)) {
          while (k < s.length && /[0-9]/.test(s[k]!)) k++;
          j = k;
        }
      }
      const raw = s.slice(i, j);
      const v = Number(raw);
      if (Number.isNaN(v)) throw new FormulaError('#VALUE!');
      out.push({ t: 'num', v });
      i = j;
      continue;
    }
    if (c === '"') {
      let j = i + 1;
      let val = '';
      while (j < s.length) {
        if (s[j] === '"' && s[j + 1] === '"') {
          val += '"';
          j += 2;
        } else if (s[j] === '"') {
          break;
        } else {
          val += s[j];
          j++;
        }
      }
      if (s[j] !== '"') throw new FormulaError('#VALUE!');
      out.push({ t: 'str', v: val });
      i = j + 1;
      continue;
    }
    if (c === "'") {
      // quoted sheet name: 'My Sheet'!A1
      let j = i + 1;
      let name = '';
      while (j < s.length && s[j] !== "'") {
        name += s[j];
        j++;
      }
      if (s[j] !== "'" || s[j + 1] !== '!') throw new FormulaError('#REF!');
      // consume the reference part after !
      let k = j + 2;
      let ref = '';
      while (k < s.length && /[A-Za-z0-9$:]/.test(s[k]!)) {
        ref += s[k];
        k++;
      }
      out.push({ t: 'ref', v: `'${name}'!${ref}` });
      i = k;
      continue;
    }
    if (IDENT_START.test(c)) {
      let j = i;
      while (j < s.length && IDENT_CHAR.test(s[j]!)) j++;
      let word = s.slice(i, j);
      if (s[j] === '!') {
        // sheet-qualified reference
        let k = j + 1;
        let ref = '';
        while (k < s.length && /[A-Za-z0-9$:]/.test(s[k]!)) {
          ref += s[k];
          k++;
        }
        out.push({ t: 'ref', v: `${word}!${ref}` });
        i = k;
        continue;
      }
      const upper = word.toUpperCase();
      if (upper === 'TRUE') out.push({ t: 'bool', v: true });
      else if (upper === 'FALSE') out.push({ t: 'bool', v: false });
      else out.push({ t: 'ident', v: word });
      i = j;
      continue;
    }
    if (c === '$') {
      // absolute bare reference like $A$1 — strip and re-lex as ident
      let j = i;
      let word = '';
      while (j < s.length && /[A-Za-z0-9$]/.test(s[j]!)) {
        if (s[j] !== '$') word += s[j];
        j++;
      }
      if (s[j] === ':') {
        // range like $A$1:$B$2
        let k = j + 1;
        let word2 = '';
        while (k < s.length && /[A-Za-z0-9$]/.test(s[k]!)) {
          if (s[k] !== '$') word2 += s[k];
          k++;
        }
        out.push({ t: 'ident', v: `${word}:${word2}` });
        i = k;
        continue;
      }
      out.push({ t: 'ident', v: word });
      i = j;
      continue;
    }
    if (c === '(') {
      out.push({ t: 'lparen' });
      i++;
      continue;
    }
    if (c === ')') {
      out.push({ t: 'rparen' });
      i++;
      continue;
    }
    if (c === ',') {
      out.push({ t: 'comma' });
      i++;
      continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&%<>=:'.includes(c)) {
      out.push({ t: 'op', v: c });
      i++;
      continue;
    }
    throw new FormulaError('#VALUE!');
  }
  return out;
}

/* ------------------------- parser (Pratt) ------------------------- */

type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'ref'; v: string }
  | { k: 'range'; v: string }
  | { k: 'call'; name: string; args: Node[] }
  | { k: 'bin'; op: string; l: Node; r: Node }
  | { k: 'un'; op: string; e: Node };

const CELL_REF_RE = /^[A-Z]{1,3}\d+$/;
const RANGE_REF_RE = /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/;

class Parser {
  private pos = 0;
  constructor(private tokens: Token[]) {}

  parse(): Node {
    const n = this.expr(0);
    if (this.pos < this.tokens.length) throw new FormulaError('#VALUE!');
    return n;
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private next(): Token {
    const t = this.tokens[this.pos++];
    if (!t) throw new FormulaError('#VALUE!');
    return t;
  }

  private bindingPower(op: string): number {
    switch (op) {
      case '=':
      case '<>':
      case '<':
      case '>':
      case '<=':
      case '>=':
        return 1;
      case '&':
        return 2;
      case '+':
      case '-':
        return 3;
      case '*':
      case '/':
        return 4;
      case '^':
        return 5;
      default:
        return 0;
    }
  }

  private expr(minBp: number): Node {
    let left = this.atom();
    for (;;) {
      const t = this.peek();
      if (!t || t.t !== 'op') break;
      const bp = this.bindingPower(t.v);
      if (bp === 0 || bp < minBp) break;
      this.next();
      const right = this.expr(bp + 1);
      left = { k: 'bin', op: t.v, l: left, r: right };
    }
    return left;
  }

  private atom(): Node {
    const t = this.next();
    switch (t.t) {
      case 'num':
        return { k: 'num', v: t.v };
      case 'str':
        return { k: 'str', v: t.v };
      case 'bool':
        return { k: 'bool', v: t.v };
      case 'ref': {
        const cleaned = t.v.replace(/\$/g, '');
        return cleaned.includes(':') ? { k: 'range', v: cleaned } : { k: 'ref', v: cleaned };
      }
      case 'op':
        if (t.v === '-') return { k: 'un', op: '-', e: this.atom() };
        if (t.v === '+') return this.atom();
        throw new FormulaError('#VALUE!');
      case 'lparen': {
        const inner = this.expr(0);
        const close = this.next();
        if (close.t !== 'rparen') throw new FormulaError('#VALUE!');
        return inner;
      }
      case 'ident': {
        const upper = t.v.toUpperCase();
        const nxt = this.peek();
        if (nxt && nxt.t === 'lparen') {
          this.next();
          const args: Node[] = [];
          if (this.peek()?.t !== 'rparen') {
            for (;;) {
              args.push(this.expr(0));
              const sep = this.next();
              if (sep.t === 'rparen') break;
              if (sep.t !== 'comma') throw new FormulaError('#VALUE!');
            }
          } else {
            this.next();
          }
          return { k: 'call', name: upper, args };
        }
        // bare cell ref or range (A1, A1:B2)
        if (RANGE_REF_RE.test(upper)) return { k: 'range', v: upper };
        if (CELL_REF_RE.test(upper)) {
          // possibly part of "A1:B2" written with op ':'
          const colon = this.peek();
          if (colon && colon.t === 'op' && colon.v === ':') {
            this.next();
            const end = this.next();
            if (end.t !== 'ident' || !CELL_REF_RE.test(end.v.toUpperCase())) {
              throw new FormulaError('#REF!');
            }
            return { k: 'range', v: `${upper}:${end.v.toUpperCase()}` };
          }
          return { k: 'ref', v: upper };
        }
        throw new FormulaError('#NAME?');
      }
      default:
        throw new FormulaError('#VALUE!');
    }
  }
}

/* ------------------------- evaluation ------------------------- */

function toNumber(v: CellScalar): number {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    if (v.startsWith('#')) throw new FormulaError(v as EvalError);
    if (v.trim() === '') return 0;
    const n = Number(v);
    if (Number.isNaN(n)) throw new FormulaError('#VALUE!');
    return n;
  }
  throw new FormulaError('#VALUE!');
}

function toText(v: CellScalar): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function propagate(v: CellScalar): CellScalar {
  if (typeof v === 'string' && v.startsWith('#')) throw new FormulaError(v as EvalError);
  return v;
}

function flattenNumeric(values: CellScalar[][]): number[] {
  const out: number[] = [];
  for (const row of values) {
    for (const v of row) {
      if (typeof v === 'string' && v.startsWith('#')) throw new FormulaError(v as EvalError);
      if (typeof v === 'number') out.push(v);
    }
  }
  return out;
}

export function evaluateFormula(
  formula: string,
  sheet: string,
  resolver: CellResolver,
): CellScalar {
  const body = formula.startsWith('=') ? formula.slice(1) : formula;
  try {
    const ast = new Parser(tokenize(body)).parse();
    return evalNode(ast, sheet, resolver);
  } catch (e) {
    if (e instanceof FormulaError) return e.code;
    return '#VALUE!';
  }
}

function resolveRangeRef(ref: string, defaultSheet: string): RangeAddress {
  // ref forms: "A1", "A1:B2", "Sheet1!A1", "'My Sheet'!A1:B2"
  let sheet = defaultSheet;
  let body = ref;
  const bang = ref.lastIndexOf('!');
  if (bang >= 0) {
    sheet = ref.slice(0, bang).replace(/^'|'$/g, '');
    body = ref.slice(bang + 1);
  }
  const parts = body.split(':');
  const parseCell = (c: string): { row: number; col: number } => {
    const m = /^([A-Z]{1,3})(\d+)$/.exec(c.toUpperCase().replace(/\$/g, ''));
    if (!m) throw new FormulaError('#REF!');
    return { row: parseInt(m[2]!, 10) - 1, col: colToIndex(m[1]!) };
  };
  const a = parseCell(parts[0]!);
  const b = parts[1] ? parseCell(parts[1]) : a;
  return {
    sheet,
    startRow: Math.min(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endRow: Math.max(a.row, b.row),
    endCol: Math.max(a.col, b.col),
  };
}

function evalNode(n: Node, sheet: string, r: CellResolver): CellScalar {
  switch (n.k) {
    case 'num':
      return n.v;
    case 'str':
      return n.v;
    case 'bool':
      return n.v;
    case 'ref': {
      const addr = resolveRangeRef(n.v, sheet);
      if (!r.sheetExists(addr.sheet)) throw new FormulaError('#REF!');
      return r.getCell(addr.sheet, addr.startRow, addr.startCol);
    }
    case 'range':
      throw new FormulaError('#VALUE!'); // bare range outside a function
    case 'un': {
      const v = toNumber(propagate(evalNode(n.e, sheet, r)));
      return n.op === '-' ? -v : v;
    }
    case 'bin':
      return evalBinary(n, sheet, r);
    case 'call':
      return evalCall(n, sheet, r);
  }
}

function evalBinary(n: Extract<Node, { k: 'bin' }>, sheet: string, r: CellResolver): CellScalar {
  const l = propagate(evalNode(n.l, sheet, r));
  const rt = propagate(evalNode(n.r, sheet, r));
  switch (n.op) {
    case '+':
      return toNumber(l) + toNumber(rt);
    case '-':
      return toNumber(l) - toNumber(rt);
    case '*':
      return toNumber(l) * toNumber(rt);
    case '/': {
      const d = toNumber(rt);
      if (d === 0) throw new FormulaError('#DIV/0!');
      return toNumber(l) / d;
    }
    case '^':
      return Math.pow(toNumber(l), toNumber(rt));
    case '&':
      return toText(l) + toText(rt);
    case '=':
      return cmp(l, rt) === 0;
    case '<>':
      return cmp(l, rt) !== 0;
    case '<':
      return cmp(l, rt) < 0;
    case '>':
      return cmp(l, rt) > 0;
    case '<=':
      return cmp(l, rt) <= 0;
    case '>=':
      return cmp(l, rt) >= 0;
    default:
      throw new FormulaError('#VALUE!');
  }
}

function cmp(a: CellScalar, b: CellScalar): number {
  if (typeof a === 'string' && typeof b === 'string') {
    const al = a.toLowerCase();
    const bl = b.toLowerCase();
    return al < bl ? -1 : al > bl ? 1 : 0;
  }
  const an = toNumber(a);
  const bn = toNumber(b);
  return an < bn ? -1 : an > bn ? 1 : 0;
}

type ArgValues = { scalar?: CellScalar; range?: CellScalar[][] };

function evalArg(n: Node, sheet: string, r: CellResolver): ArgValues {
  if (n.k === 'range') {
    const addr = resolveRangeRef(n.v, sheet);
    if (!r.sheetExists(addr.sheet)) throw new FormulaError('#REF!');
    return { range: r.getRange(addr) };
  }
  return { scalar: evalNode(n, sheet, r) };
}

function numericValues(args: ArgValues[]): number[] {
  const out: number[] = [];
  for (const a of args) {
    if (a.range) out.push(...flattenNumeric(a.range));
    else if (a.scalar !== undefined && a.scalar !== null && a.scalar !== '') {
      out.push(toNumber(propagate(a.scalar)));
    }
  }
  return out;
}

function evalCall(n: Extract<Node, { k: 'call' }>, sheet: string, r: CellResolver): CellScalar {
  const name = n.name;
  const lazy = (i: number): Node | undefined => n.args[i];
  const args = (): ArgValues[] => n.args.map((a) => evalArg(a, sheet, r));

  switch (name) {
    case 'SUM':
      return numericValues(args()).reduce((a, b) => a + b, 0);
    case 'AVERAGE': {
      const vals = numericValues(args());
      if (vals.length === 0) throw new FormulaError('#DIV/0!');
      return vals.reduce((a, b) => a + b, 0) / vals.length;
    }
    case 'MIN': {
      const vals = numericValues(args());
      return vals.length === 0 ? 0 : Math.min(...vals);
    }
    case 'MAX': {
      const vals = numericValues(args());
      return vals.length === 0 ? 0 : Math.max(...vals);
    }
    case 'COUNT': {
      let c = 0;
      for (const a of args()) {
        if (a.range) c += flattenNumeric(a.range).length;
        else if (typeof a.scalar === 'number') c++;
      }
      return c;
    }
    case 'COUNTA': {
      let c = 0;
      for (const a of args()) {
        if (a.range) {
          for (const row of a.range) for (const v of row) if (v !== null && v !== '') c++;
        } else if (a.scalar !== null && a.scalar !== '') c++;
      }
      return c;
    }
    case 'IF': {
      if (n.args.length < 2 || n.args.length > 3) throw new FormulaError('#VALUE!');
      const cond = propagate(evalNode(lazy(0)!, sheet, r));
      const truthy = typeof cond === 'boolean' ? cond : toNumber(cond) !== 0;
      if (truthy) return evalNode(lazy(1)!, sheet, r);
      return n.args.length === 3 ? evalNode(lazy(2)!, sheet, r) : false;
    }
    case 'IFERROR': {
      if (n.args.length !== 2) throw new FormulaError('#VALUE!');
      try {
        return propagate(evalNode(lazy(0)!, sheet, r));
      } catch (e) {
        if (e instanceof FormulaError) return evalNode(lazy(1)!, sheet, r);
        throw e;
      }
    }
    case 'AND': {
      for (const a of args()) {
        const v = a.scalar !== undefined ? a.scalar : null;
        if (a.range) throw new FormulaError('#VALUE!');
        if (!(typeof v === 'boolean' ? v : toNumber(propagate(v)) !== 0)) return false;
      }
      return true;
    }
    case 'OR': {
      for (const a of args()) {
        if (a.range) throw new FormulaError('#VALUE!');
        const v = a.scalar !== undefined ? a.scalar : null;
        if (typeof v === 'boolean' ? v : toNumber(propagate(v)) !== 0) return true;
      }
      return false;
    }
    case 'NOT': {
      const v = propagate(evalNode(lazy(0)!, sheet, r));
      return !(typeof v === 'boolean' ? v : toNumber(v) !== 0);
    }
    case 'ROUND': {
      const a = args();
      const v = toNumber(propagate(a[0]?.scalar ?? null));
      const digits = a.length > 1 ? toNumber(propagate(a[1]?.scalar ?? null)) : 0;
      const f = Math.pow(10, digits);
      return Math.round(v * f) / f;
    }
    case 'ABS':
      return Math.abs(toNumber(propagate(args()[0]?.scalar ?? null)));
    case 'LEN':
      return toText(propagate(args()[0]?.scalar ?? null)).length;
    case 'UPPER':
      return toText(propagate(args()[0]?.scalar ?? null)).toUpperCase();
    case 'LOWER':
      return toText(propagate(args()[0]?.scalar ?? null)).toLowerCase();
    case 'TRIM':
      return toText(propagate(args()[0]?.scalar ?? null)).trim().replace(/ +/g, ' ');
    case 'CONCATENATE':
    case 'CONCAT': {
      let s = '';
      for (const a of args()) {
        if (a.range) {
          for (const row of a.range) for (const v of row) s += toText(propagate(v));
        } else s += toText(propagate(a.scalar ?? null));
      }
      return s;
    }
    default:
      throw new FormulaError('#NAME?');
  }
}

/** Extract all cell/range references from a formula (for dependency tracking). */
export function extractReferences(formula: string, defaultSheet: string): RangeAddress[] {
  const body = formula.startsWith('=') ? formula.slice(1) : formula;
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return [];
  }
  const refs: RangeAddress[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.t === 'ref') {
      try {
        refs.push(resolveRangeRef(t.v.replace(/\$/g, ''), defaultSheet));
      } catch {
        /* unparseable ref — ignore */
      }
    } else if (t.t === 'ident') {
      const upper = t.v.toUpperCase();
      const next = tokens[i + 1];
      const isCall = next?.t === 'lparen';
      if (isCall) continue;
      if (CELL_REF_RE.test(upper) || RANGE_REF_RE.test(upper)) {
        // "A1 : B2" arrives as three tokens; merge them into one range.
        let refText = upper;
        const after = tokens[i + 2];
        if (
          CELL_REF_RE.test(upper) &&
          next?.t === 'op' &&
          next.v === ':' &&
          after?.t === 'ident' &&
          CELL_REF_RE.test(after.v.toUpperCase())
        ) {
          refText = `${upper}:${after.v.toUpperCase()}`;
          i += 2;
        }
        try {
          refs.push(resolveRangeRef(refText, defaultSheet));
        } catch {
          /* ignore */
        }
      }
    }
  }
  return refs;
}
