import { describe, test, expect } from 'vitest';
import { ExcelFormulaParser } from '../../../src/core/formula/FormulaParser';
import { TokenType } from '../../../src/core/formula/types';

describe('ExcelFormulaParser', () => {
  const parser = new ExcelFormulaParser();

  describe('Basic Parsing', () => {
    test('parses simple arithmetic', () => {
      const result = parser.parse('A1 + B1');
      expect(result.type).toBe('OPERATOR');
      expect(result.value).toBe('+');
      expect(result.children).toHaveLength(2);
      expect(result.children[0].type).toBe(TokenType.CELL_REFERENCE);
      expect(result.children[1].type).toBe(TokenType.CELL_REFERENCE);
    });

    test('parses numbers', () => {
      const result = parser.parse('1.5 + 2');
      expect(result.type).toBe('OPERATOR');
      expect(result.children[0].type).toBe(TokenType.NUMBER);
      expect(result.children[0].value).toBe('1.5');
    });

    test('parses scientific notation', () => {
      const result = parser.parse('1.5E+5');
      expect(result.type).toBe(TokenType.NUMBER);
      expect(result.value).toBe('1.5E+5');
    });

    test('handles whitespace', () => {
      const result = parser.parse('  SUM(  A1:A10  )  ');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('SUM');
    });
  });

  describe('Function Parsing', () => {
    test('parses simple function', () => {
      const result = parser.parse('SUM(A1:A10)');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('SUM');
      expect(result.children).toHaveLength(1);
      expect(result.children[0].type).toBe(TokenType.RANGE_REFERENCE);
    });

    test('parses nested functions', () => {
      const result = parser.parse('IF(SUM(A1:A10)>0,MAX(B1:B10),0)');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('IF');
      expect(result.children).toHaveLength(3);
    });

    test('handles function with multiple arguments', () => {
      const result = parser.parse('SUMIFS(A1:A10,B1:B10,">0",C1:C10,"<100")');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('SUMIFS');
      expect(result.children.length).toBeGreaterThan(2);
    });

    test('handles empty function arguments', () => {
      expect(() => parser.parse('SUM()')).toThrow();
    });
  });

  describe('Cell References', () => {
    test('parses absolute references', () => {
      const result = parser.parse('$A$1 + B$2');
      expect(result.children[0].type).toBe(TokenType.CELL_REFERENCE);
      expect(result.children[0].value).toBe('$A$1');
    });

    test('parses range references', () => {
      const result = parser.parse('A1:B10');
      expect(result.type).toBe(TokenType.RANGE_REFERENCE);
      expect(result.value).toBe('A1:B10');
    });

    test('parses mixed references', () => {
      const result = parser.parse('$A1 + A$1');
      expect(result.type).toBe('OPERATOR');
      expect(result.children[0].value).toBe('$A1');
      expect(result.children[1].value).toBe('A$1');
    });
  });

  describe('Formula Analysis', () => {
    test('analyzes simple formula', () => {
      const analysis = parser.analyze('A1 + B1');
      expect(analysis.isValid).toBe(true);
      expect(analysis.complexity).toBeLessThan(1);
      expect(analysis.dependencies).toHaveLength(2);
      expect(analysis.errors).toHaveLength(0);
    });

    // TODO need to revist it 
/*     test('analyzes complex formula', () => {
      const analysis = parser.analyze('IF(SUM(A1:A10)>0,MAX(B1:B10),0)');
      expect(analysis.isValid).toBe(true);
      expect(analysis.complexity).toBeGreaterThan(2);
      expect(analysis.dependencies.length).toBeGreaterThan(0);
    });

    test('detects volatile functions', () => {
      const analysis = parser.analyze('NOW() + A1');
      expect(analysis.volatileFunctions).toContain('NOW');
      expect(analysis.errors.length).toBeGreaterThan(0);
    });

    test('identifies circular references', () => {
      const analysis = parser.analyze('A1 + B1 + A1');
      expect(analysis.errors).toContain(expect.stringContaining('circular reference'));
    }); */
  });

  describe('Error Handling', () => {
    test('handles unmatched parentheses', () => {
      const analysis = parser.analyze('SUM(A1:A10');
      expect(analysis.isValid).toBe(false);
      expect(analysis.errors).toContain('Unmatched parentheses');
    });

    test('handles invalid function names', () => {
      const result = parser.parse('INVALID(A1)');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('INVALID');
    });

    test('handles malformed formulas', () => {
      const analysis = parser.analyze('A1 + + B1');
      expect(analysis.isValid).toBe(false);
    });

    test('handles empty formulas', () => {
      expect(() => parser.parse('')).toThrow();
    });
  });

  describe('Complex Operations', () => {
    test('handles complex mathematical expressions', () => {
      const result = parser.parse('(A1 + B1) * (C1 - D1) / E1');
      expect(result.type).toBe('OPERATOR');
      expect(result.children.length).toBeGreaterThan(0);
    });

    test('handles text concatenation', () => {
      const result = parser.parse('"Hello" & " " & A1');
      expect(result.type).toBe('OPERATOR');
      expect(result.value).toBe('&');
    });

    test('handles comparison operators', () => {
      const result = parser.parse('A1 >= B1');
      expect(result.type).toBe('OPERATOR');
      expect(result.value).toBe('>=');
    });
  });

  describe('Performance', () => {
    test('handles large formulas efficiently', () => {
      const start = performance.now();
      const largeFormula = 'IF(AND(A1>0,B1>0,C1>0),SUM(D1:D100),AVERAGE(E1:E100))';
      const result = parser.analyze(largeFormula);
      const duration = performance.now() - start;
      
      expect(duration).toBeLessThan(100); // Should complete within 100ms
      expect(result.isValid).toBe(true);
    });
  });

  describe('Feature Edge Cases', () => {
    test('handles nested parentheses correctly', () => {
      const result = parser.parse('((A1 + B1) * (C1 + D1))');
      expect(result.type).toBe('OPERATOR');
    });

    test('handles multiple operators in sequence', () => {
      const result = parser.parse('A1 + B1 * C1 / D1');
      expect(result.type).toBe('OPERATOR');
    });

    test('handles string literals with spaces', () => {
      const result = parser.parse('"Hello World" & A1');
      expect(result.children[0].type).toBe(TokenType.TEXT);
      expect(result.children[0].value).toBe('Hello World');
    });

    test('handles array formulas', () => {
      const result = parser.parse('SUM(IF(A1:A10>0,B1:B10,0))');
      expect(result.type).toBe('FUNCTION');
      expect(result.value).toBe('SUM');
    });
  });
});