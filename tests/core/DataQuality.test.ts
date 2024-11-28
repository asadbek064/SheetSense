import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelAnalyzer } from '../../src/core/Analyzer';

test('detects mixed data types in columns', () => {
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:A4',
    A1: { v: 'Header', t: 's' },
    A2: { v: 42, t: 'n' },
    A3: { v: 'text', t: 's' },
    A4: { v: true, t: 'b' }
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();

  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'data',
      severity: 'warning',
      message: expect.stringContaining('Mixed data types')
    })
  );
});

test('detects invalid dates', () => {
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:A3',
    A1: { v: '2024-01-01', t: 's' },
    A2: { v: 'invalid date', t: 's' },
    A3: { v: '2024-13-45', t: 's' }
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();

  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'data',
      severity: 'error',
      message: expect.stringContaining('Invalid date')
    })
  );
});