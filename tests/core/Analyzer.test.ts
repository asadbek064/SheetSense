import { test, expect } from 'vitest';
import { WorkBook } from 'xlsx';
import { ExcelAnalyzer } from '../../src/core/Analyzer';

test('analyzes workbook', () => {
  // Create a workbook with a complex formula
  const wb = {
    SheetNames: ['Sheet1'],
    Sheets: {
      'Sheet1': {
        '!ref': 'A1:B2',
        'A1': { v: 'Simple', t: 's' },
        'B1': { v: 'Complex', t: 's' },
        'A2': { v: 0, t: 'n', f: 'A1+B1' },
        'B2': { 
          v: 0, 
          t: 'n', 
          f: 'IF(A1>0,IF(A1<10,IF(A1<5,"Low","Medium"),"High"))' 
        }
      }
    }
  } as WorkBook;

  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();

  // Should detect the complex nested IF formula
  expect(analysis.issues.length).toBeGreaterThan(0);
  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'formula',
      severity: 'warning',
      message: expect.stringContaining('Complex formula')
    })
  );
  expect(analysis.metadata.formulaCount).toBe(2);
});