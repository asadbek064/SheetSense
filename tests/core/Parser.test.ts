import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelParser } from '../../src/core/Parser';

test('detects formulas in worksheet', () => {
  // Create worksheet with explicit formula
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:B2',
    'A1': { v: 'Header', t: 's' },
    'B1': { v: 'Formula', t: 's' },
    'B2': { v: 2, t: 'n', f: 'A1*2' }  // Explicit formula
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  const parser = new ExcelParser(wb);
  const formulas = parser.getFormulaCells(wb.Sheets['Sheet1']);
  
  expect(formulas.size).toBeGreaterThan(0);
  expect(formulas.get('B2')).toBe('A1*2');
});