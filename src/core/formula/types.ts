export enum TokenType {
  FUNCTION = 'FUNCTION',
  OPERATOR = 'OPERATOR',
  CELL_REFERENCE = 'CELL_REFERENCE',
  RANGE_REFERENCE = 'RANGE_REFERENCE',
  NUMBER = 'NUMBER',
  TEXT = 'TEXT',
  ERROR = 'ERROR',
  UNKNOWN = 'UNKNOWN'
}

export interface Token {
  type: TokenType;
  value: string;
  position: {
    start: number;
    end: number;
  };
}

export interface FormulaNode {
  type: string;
  value: string;
  children: FormulaNode[];
  token: Token;
}

export interface ParsedFormula {
  formula: string;
  references: string[];
  functions: string[];
  dependencies: Set<string>;
  complexity: number;
}

// Custom type for XLSX formula parsing result
export type XLSXParsedFormula = (string | number | {
  t: string;
  v: string;
})[];

export interface FormulaAnalysis {
  isValid: boolean;
  errors: string[];
  complexity: number;
  volatileFunctions: string[];
  dependencies: string[];
  usedRanges: string[];
}