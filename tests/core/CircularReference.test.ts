import { describe, test, expect } from 'vitest';
import { utils } from 'xlsx';
import { ExcelAnalyzer } from '../../src/core/Analyzer';
import { readFileSync } from 'fs';

describe('CircularReferences', () => {
  test('detects direct self-reference', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:A1',
      'A1': { t: 'str', v: '=A1+1' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();
    
    //console.log('Analysis:', JSON.stringify(analysis, null, 2));
    //console.log('Cell A1:', JSON.stringify(ws['A1'], null, 2));

    expect(analysis.metadata.formulaCount).toBe(1);
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
      'A1': { t: 'str', v: '=B1+1' },
      'B1': { t: 'str', v: '=C1+1' },
      'C1': { t: 'str', v: '=A1+1' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();

    //console.log('Analysis:', JSON.stringify(analysis, null, 2));

    expect(analysis.metadata.formulaCount).toBe(3);
    const circularIssues = analysis.issues.filter((i: any) => 
      i.type === 'formula' && 
      i.severity === 'error' && 
      i.message.includes('Circular reference')
    );

    expect(circularIssues.length).toBe(3);
    expect(circularIssues.map((i: any) => i.cell)).toContain('A1');
    expect(circularIssues.map((i: any) => i.cell)).toContain('B1');
    expect(circularIssues.map((i: any) => i.cell)).toContain('C1');
  });

  test('accepts valid cross-references', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:B1',
      'A1': { t: 'n', v: 1 },
      'B1': { t: 'str', v: '=A1*2' }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new ExcelAnalyzer(wb);
    const analysis = analyzer.analyze();

    expect(analysis.metadata.formulaCount).toBe(1);
    expect(analysis.issues.filter((i: any) => 
      i.type === 'formula' && 
      i.message.includes('Circular reference')
    ).length).toBe(0);
  });


  /* with files */
  const TEST_FILES_DIR = 'res/test_files/circular_reference';

  test('detects direct circular reference from file', () => {
    const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/direct_circular.xlsx`);
    const analyzer = ExcelAnalyzer.fromBuffer(xlsxBuffer);
    const analysis = analyzer.analyze();
    
    //console.log('Direct Circular Analysis:', JSON.stringify(analysis, null, 2));

    expect(analysis.metadata.formulaCount).toBe(1);
    expect(analysis.issues).toContainEqual(
      expect.objectContaining({
        type: 'formula',
        severity: 'error',
        message: expect.stringContaining('Circular reference'),
        cell: 'A1'
      })
    );
  });

  test('detects indirect circular reference from file', () => {
    const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/indirect_circular.xlsx`);
    const analyzer = ExcelAnalyzer.fromBuffer(xlsxBuffer);
    const analysis = analyzer.analyze();

    //console.log('Indirect Circular Analysis:', JSON.stringify(analysis, null, 2));

    expect(analysis.metadata.formulaCount).toBe(3);
    const circularIssues = analysis.issues.filter((i: any) => 
      i.type === 'formula' && 
      i.severity === 'error' && 
      i.message.includes('Circular reference')
    );

    expect(circularIssues.length).toBe(3);
    expect(circularIssues.map((i: any) => i.cell).sort()).toEqual(['A1', 'B1', 'C1']);
  });

  test('accepts valid cross-references from file', () => {
    const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/valid_cross_reference.xlsx`);
    const analyzer = ExcelAnalyzer.fromBuffer(xlsxBuffer);
    const analysis = analyzer.analyze();

    //console.log('Valid Cross Reference Analysis:', JSON.stringify(analysis, null, 2));

    expect(analysis.metadata.formulaCount).toBe(1);
    expect(analysis.issues.filter((i: any) => 
      i.type === 'formula' && 
      i.message.includes('Circular reference')
    ).length).toBe(0);
  });
});``