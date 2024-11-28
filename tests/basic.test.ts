import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelAnalyzer } from '../src/core/Analyzer';

test('detects basic formula errors', () => {
  // Create workbook with explicit formula
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:B2',
    'A1': { v: 'Value', t: 's' },
    'B1': { v: 'Formula', t: 's' },
    'A2': { v: 100, t: 'n' },
    'B2': { v: 0, t: 'n', f: 'A1/0' }  // Explicit formula
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];

  //console.log('Created workbook:', JSON.stringify(wb, null, 2));
  
  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();
  
  //console.log('Analysis result:', JSON.stringify(analysis, null, 2));
  
  expect(analysis.issues.length).toBeGreaterThan(0);
  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'formula',
      severity: 'error',
      message: expect.stringContaining('division by zero')
    })
  );
});
