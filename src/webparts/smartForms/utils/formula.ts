/**
 * Tiny arithmetic evaluator for Calculated fields.
 *
 * Formulas reference other questions by title in braces and support the four
 * operators, unary minus, parentheses and a few aggregate functions:
 *
 *   {Quantity} * {Unit price}
 *   ({Subtotal} + {Shipping}) * 1.2
 *   round({Score} / {Total} * 100)
 *   min({Budget}, {Quoted})
 *
 * Implemented as a shunting-yard parser over an explicit token list. `eval` is
 * never used — a form definition is authored by a site owner but rendered for
 * everyone, so a formula must not be able to execute arbitrary code.
 */

type TokenKind = 'number' | 'operator' | 'lparen' | 'rparen' | 'comma' | 'function';

interface IToken {
  kind: TokenKind;
  value: string;
  number?: number;
}

const PRECEDENCE: { [op: string]: number } = {
  'u-': 4,
  '^': 3,
  '*': 2,
  '/': 2,
  '%': 2,
  '+': 1,
  '-': 1
};

const RIGHT_ASSOCIATIVE: string[] = ['^', 'u-'];

const FUNCTIONS: { [name: string]: { arity: number | 'variadic'; apply: (args: number[]) => number } } = {
  round: { arity: 1, apply: (a) => Math.round(a[0]) },
  floor: { arity: 1, apply: (a) => Math.floor(a[0]) },
  ceil: { arity: 1, apply: (a) => Math.ceil(a[0]) },
  abs: { arity: 1, apply: (a) => Math.abs(a[0]) },
  sqrt: { arity: 1, apply: (a) => Math.sqrt(a[0]) },
  min: { arity: 'variadic', apply: (a) => Math.min(...a) },
  max: { arity: 'variadic', apply: (a) => Math.max(...a) },
  sum: { arity: 'variadic', apply: (a) => a.reduce((acc, n) => acc + n, 0) },
  avg: { arity: 'variadic', apply: (a) => (a.length === 0 ? 0 : a.reduce((acc, n) => acc + n, 0) / a.length) }
};

/** Names referenced by a formula, in the order they appear. */
export const formulaReferences = (formula: string): string[] => {
  const names: string[] = [];
  const pattern = /\{([^{}]*)\}/g;
  let match = pattern.exec(formula || '');
  while (match !== null) {
    const name = match[1].trim();
    if (name.length > 0 && names.indexOf(name) === -1) {
      names.push(name);
    }
    match = pattern.exec(formula || '');
  }
  return names;
};

/**
 * Render a finite number as a plain decimal, never in exponent form
 * (String(1e-7) is "1e-7", which the tokenizer would otherwise misread).
 */
export const plainDecimal = (value: number): string => {
  const text = String(value);
  if (!/e/i.test(text)) {
    return text;
  }
  const negative = value < 0;
  const match = /^(\d+)(?:\.(\d+))?e([+-]\d+)$/i.exec(String(Math.abs(value)));
  if (!match) {
    return value.toFixed(20).replace(/\.?0+$/, '');
  }
  const digits = match[1] + (match[2] || '');
  const pointIndex = match[1].length + parseInt(match[3], 10);
  let result: string;
  if (pointIndex <= 0) {
    result = '0.' + new Array(1 - pointIndex).join('0') + digits;
  } else if (pointIndex >= digits.length) {
    result = digits + new Array(pointIndex - digits.length + 1).join('0');
  } else {
    result = digits.slice(0, pointIndex) + '.' + digits.slice(pointIndex);
  }
  return (negative ? '-' : '') + result;
};

/**
 * Substitute `{Title}` references with numeric literals. Unknown or
 * non-numeric references become 0 so a partially answered form still totals.
 */
const substitute = (formula: string, resolve: (name: string) => number | undefined): string =>
  (formula || '').replace(/\{([^{}]*)\}/g, (_full, name: string) => {
    const value = resolve(String(name).trim());
    if (value === undefined || value === null || isNaN(value)) {
      return '0';
    }
    // wrap negatives so "a - {x}" with x = -5 doesn't become "a - -5" ambiguity
    return value < 0 ? '(0 - ' + plainDecimal(Math.abs(value)) + ')' : plainDecimal(value);
  });

/** Exponent suffix (e.g. "e-7") starting at index i, or '' when none. */
const readExponent = (text: string, i: number): string => {
  const match = /^[eE][+-]?\d+/.exec(text.slice(i));
  return match ? match[0] : '';
};

const tokenize = (input: string): IToken[] | undefined => {
  const tokens: IToken[] = [];
  let i = 0;
  const text = input || '';

  while (i < text.length) {
    const ch = text.charAt(i);

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }

    if (ch >= '0' && ch <= '9') {
      let numberText = '';
      while (i < text.length && ((text.charAt(i) >= '0' && text.charAt(i) <= '9') || text.charAt(i) === '.')) {
        numberText += text.charAt(i);
        i++;
      }
      const intExponent = readExponent(text, i);
      numberText += intExponent;
      i += intExponent.length;
      const parsed = Number(numberText);
      if (isNaN(parsed)) {
        return undefined;
      }
      tokens.push({ kind: 'number', value: numberText, number: parsed });
      continue;
    }

    if (ch === '.') {
      let numberText = '.';
      i++;
      while (i < text.length && text.charAt(i) >= '0' && text.charAt(i) <= '9') {
        numberText += text.charAt(i);
        i++;
      }
      const exponent = readExponent(text, i);
      numberText += exponent;
      i += exponent.length;
      const parsed = Number(numberText);
      if (isNaN(parsed)) {
        return undefined;
      }
      tokens.push({ kind: 'number', value: numberText, number: parsed });
      continue;
    }

    if (/[a-zA-Z]/.test(ch)) {
      let name = '';
      while (i < text.length && /[a-zA-Z0-9_]/.test(text.charAt(i))) {
        name += text.charAt(i);
        i++;
      }
      const lower = name.toLowerCase();
      if (!FUNCTIONS[lower]) {
        return undefined;
      }
      tokens.push({ kind: 'function', value: lower });
      continue;
    }

    if (ch === '(') {
      tokens.push({ kind: 'lparen', value: ch });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ kind: 'rparen', value: ch });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ kind: 'comma', value: ch });
      i++;
      continue;
    }

    if ('+-*/%^'.indexOf(ch) !== -1) {
      const previous = tokens[tokens.length - 1];
      const isUnary =
        ch === '-' &&
        (!previous || previous.kind === 'operator' || previous.kind === 'lparen' || previous.kind === 'comma');
      tokens.push({ kind: 'operator', value: isUnary ? 'u-' : ch });
      i++;
      continue;
    }

    // anything else is not a formula we understand
    return undefined;
  }

  return tokens;
};

const applyOperator = (op: string, stack: number[]): boolean => {
  if (op === 'u-') {
    if (stack.length < 1) {
      return false;
    }
    stack.push(-(stack.pop() as number));
    return true;
  }
  if (stack.length < 2) {
    return false;
  }
  const b = stack.pop() as number;
  const a = stack.pop() as number;
  switch (op) {
    case '+':
      stack.push(a + b);
      return true;
    case '-':
      stack.push(a - b);
      return true;
    case '*':
      stack.push(a * b);
      return true;
    case '/':
      stack.push(b === 0 ? 0 : a / b);
      return true;
    case '%':
      stack.push(b === 0 ? 0 : a % b);
      return true;
    case '^':
      stack.push(Math.pow(a, b));
      return true;
    default:
      return false;
  }
};

/**
 * Evaluate a formula. `resolve` maps a referenced question title to its current
 * numeric answer. Returns undefined when the formula is malformed, so the UI can
 * show a hint rather than a misleading zero.
 */
export const evaluateFormula = (
  formula: string,
  resolve: (name: string) => number | undefined
): number | undefined => {
  if (!formula || formula.trim().length === 0) {
    return undefined;
  }
  const tokens = tokenize(substitute(formula, resolve));
  if (!tokens || tokens.length === 0) {
    return undefined;
  }

  const values: number[] = [];
  const operators: IToken[] = [];
  // argument counts for each open function call
  const argCounts: number[] = [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    if (token.kind === 'number') {
      values.push(token.number as number);
      continue;
    }

    if (token.kind === 'function') {
      operators.push(token);
      const next = tokens[i + 1];
      if (!next || next.kind !== 'lparen') {
        return undefined;
      }
      continue;
    }

    if (token.kind === 'lparen') {
      operators.push(token);
      const previous = tokens[i - 1];
      if (previous && previous.kind === 'function') {
        // an empty argument list is invalid; assume at least one
        argCounts.push(1);
      }
      continue;
    }

    if (token.kind === 'comma') {
      while (operators.length > 0 && operators[operators.length - 1].kind !== 'lparen') {
        if (!applyOperator((operators.pop() as IToken).value, values)) {
          return undefined;
        }
      }
      if (operators.length === 0 || argCounts.length === 0) {
        return undefined;
      }
      // the comma must sit directly inside a function call's own parentheses,
      // so "max((1,2))" (a bare group holding a comma) is rejected
      const openParen = operators.length - 1;
      const owner = operators[openParen - 1];
      if (!owner || owner.kind !== 'function') {
        return undefined;
      }
      argCounts[argCounts.length - 1]++;
      continue;
    }

    if (token.kind === 'rparen') {
      while (operators.length > 0 && operators[operators.length - 1].kind !== 'lparen') {
        if (!applyOperator((operators.pop() as IToken).value, values)) {
          return undefined;
        }
      }
      if (operators.length === 0) {
        return undefined;
      }
      operators.pop(); // discard the '('
      if (operators.length > 0 && operators[operators.length - 1].kind === 'function') {
        const fn = operators.pop() as IToken;
        const spec = FUNCTIONS[fn.value];
        const count = argCounts.pop() || 1;
        if (spec.arity !== 'variadic' && spec.arity !== count) {
          return undefined;
        }
        if (values.length < count) {
          return undefined;
        }
        const args = values.splice(values.length - count, count);
        values.push(spec.apply(args));
      }
      continue;
    }

    // operator
    const precedence = PRECEDENCE[token.value];
    while (operators.length > 0) {
      const top = operators[operators.length - 1];
      if (top.kind !== 'operator') {
        break;
      }
      const topPrecedence = PRECEDENCE[top.value];
      const shouldPop =
        RIGHT_ASSOCIATIVE.indexOf(token.value) === -1
          ? topPrecedence >= precedence
          : topPrecedence > precedence;
      if (!shouldPop) {
        break;
      }
      if (!applyOperator((operators.pop() as IToken).value, values)) {
        return undefined;
      }
    }
    operators.push(token);
  }

  while (operators.length > 0) {
    const op = operators.pop() as IToken;
    if (op.kind !== 'operator') {
      return undefined;
    }
    if (!applyOperator(op.value, values)) {
      return undefined;
    }
  }

  if (values.length !== 1) {
    return undefined;
  }
  const result = values[0];
  return isNaN(result) || !isFinite(result) ? undefined : result;
};

/** True when a formula parses. Used by the designer to validate as you type. */
export const isValidFormula = (formula: string): boolean =>
  evaluateFormula(formula, () => 1) !== undefined;
