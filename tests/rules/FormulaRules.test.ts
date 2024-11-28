import { describe, test, expect } from 'vitest';
import { validationFormula } from '../../src/rules/FormulaRules';

describe('FormulaRules', () => {
  test('detects division by zero', () => {
    const issues = validationFormula('A1/0', 'B2', 'Sheet1');
    expect(issues[0]).toMatchObject({
      type: 'formula',
      severity: 'error',
      message: expect.stringContaining('division by zero')
    });
  });

  test('detects complex nested formulas', () => {
    const formula = 'IF(A1>0,IF(A1<10,IF(A1<5,"Low","Medium"),"High"))';
    const issues = validationFormula(formula, 'B2', 'Sheet1');
    expect(issues[0]).toMatchObject({
      type: 'formula',
      severity: 'warning',
      message: expect.stringContaining('Complex formula')
    });
  });

  test('accepts simple formulas', () => {
    const issues = validationFormula('A1+B1*2', 'C1', 'Sheet1');
    expect(issues.length).toBe(0);
  });
});