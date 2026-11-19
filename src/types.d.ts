declare module 'ssf' {
  const SSF: {
    format(fmt: string | number, value: unknown, opts?: { date1904?: boolean }): string;
    is_date(fmt: string): boolean;
  };
  export default SSF;
}

declare module 'fast-formula-parser/formulas/operators.js' {
  const operators: unknown;
  export default operators;
}

declare module 'fast-formula-parser' {
  export class FormulaError {
    constructor(error: string, details?: string);
    readonly error: string;
    readonly message: string;
  }
  export interface CellRef {
    sheet: string;
    row: number;
    col: number;
  }
  export interface RangeRef {
    sheet: string;
    from: { row: number; col: number };
    to: { row: number; col: number };
  }
  export interface ParserConfig {
    onCell?: (ref: CellRef) => unknown;
    onRange?: (ref: RangeRef) => unknown[][];
    onVariable?: (name: string, sheetName: string) => CellRef | RangeRef | undefined;
    functions?: Record<string, (...args: unknown[]) => unknown>;
  }
  export default class FormulaParser {
    constructor(config?: ParserConfig);
    parse(formula: string, position: CellRef, allowReturnArray?: boolean): unknown;
    static FormulaError: typeof FormulaError;
  }
}
