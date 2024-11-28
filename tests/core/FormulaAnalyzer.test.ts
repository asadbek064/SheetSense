import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelParser } from '../../src/core/Parser';
import { FormulaAnalyzer } from '../../src/core/FormulaAnalyzer';

test('detects complex formulas', () => {
  const wb = utils.book_new();
  // Create worksheet directly with proper formula structure
  const ws = {
    '!ref': 'A1:B2',
    'A1': { v: 'Simple', t: 's' },
    'B1': { v: 'Complex', t: 's' },
    'A2': { v: 0, t: 'n', f: 'A1+B1' },
    'B2': { v: 0, t: 'n', f: 'IF(A1>0,IF(A1<10,IF(A1<5,"Low","Medium"),"High"))' }
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  const parser = new ExcelParser(wb);
  const formulas = parser.getFormulaCells(wb.Sheets['Sheet1']);
  
  const complexFormula = formulas.get('B2');
  if (!complexFormula) throw new Error('Formula not found');
  
  const analyzer = new FormulaAnalyzer(parser);
  const issues = analyzer.analyzeFormula(complexFormula, 'B2', 'Sheet1');

  expect(issues.length).toBeGreaterThan(0);
  expect(issues[0]).toMatchObject({
    type: 'formula',
    severity: 'warning',
    message: expect.stringContaining('Complex formula')
  });
});
