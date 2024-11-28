import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelAnalyzer } from '../../src/core/Analyzer';

test('detects inconsistent number formatting', () => {
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:A4',
    'A1': { v: 'Amount', t: 's' },
    'A2': { v: '1000.00', t: 's' },
    'A3': { v: '1,000', t: 's' },
    'A4': { v: '1000.0', t: 's' }
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();

  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'style',
      severity: 'warning',
      message: expect.stringContaining('Inconsistent number format')
    })
  );
});