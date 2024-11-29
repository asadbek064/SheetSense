import { describe, test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelAnalyzer } from '../../src/core/Analyzer';

describe('CircularReferences', () => {
  test('detects direct self-reference', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:A1',
      'A1': { v: 0, t: 'n', f: 'A1+1' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();

    expect(analysis.issues).toContainEqual(
      expect.objectContaining({
        type: 'formula',
        severity: 'error',
        message: expect.stringContaining('Circular reference'),
        cell: 'A1'
      })
    );
  });

  test('detects indirect circular reference', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:C1',
      'A1': { v: 0, t: 'n', f: 'B1+1' },
      'B1': { v: 0, t: 'n', f: 'C1+1' },
      'C1': { v: 0, t: 'n', f: 'A1+1' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();

    const circularIssues = analysis.issues.filter(i => 
      i.type === 'formula' && 
      i.severity === 'error' && 
      i.message.includes('Circular reference')
    );

    expect(circularIssues.length).toBeGreaterThan(0);
    expect(circularIssues.map(i => i.cell)).toContain('A1');
  });

  test('accepts valid cross-references', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:B1',
      'A1': { v: 1, t: 'n' },
      'B1': { v: 2, t: 'n', f: 'A1*2' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();

    const circularIssues = analysis.issues.filter(i => 
      i.type === 'formula' && 
      i.message.includes('Circular reference')
    );

    expect(circularIssues.length).toBe(0);
  });
});