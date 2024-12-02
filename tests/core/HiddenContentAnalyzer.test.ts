import { describe, test, expect } from 'vitest';
import { utils, read } from 'xlsx';
import { HiddenContentAnalyzer } from '../../src/core/HiddenContentAnalyzer';
import { readFileSync } from 'fs';

describe('HiddenContentAnalyzer', () => {
  test('detects hidden cells', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:B2',
      'A1': { v: 'Visible', t: 's' },
      'B1': { v: 'Hidden', t: 's', h: true }
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new HiddenContentAnalyzer(wb);
    const { issues, metrics } = analyzer.analyze(ws);

    expect(metrics.hiddenCells).toBe(1);
    expect(issues).toContainEqual(
      expect.objectContaining({
        type: 'data',
        severity: 'warning',
        message: expect.stringContaining('Hidden cells detected')
      })
    );
  });

  test('detects hidden rows and columns', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:C3',
      '!rows': [
        { hidden: true },  // Row 1
        {},                // Row 2
        { hidden: true }   // Row 3
      ],
      '!cols': [
        { hidden: true },  // Col A
        {},                // Col B
        { hidden: true }   // Col C
      ]
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new HiddenContentAnalyzer(wb);
    const { issues, metrics } = analyzer.analyze(ws);

    expect(metrics.hiddenRows).toBe(2);
    expect(metrics.hiddenColumns).toBe(2);
    expect(issues.length).toBe(2);
    expect(metrics.hiddenRanges).toContain('Row 1');
    expect(metrics.hiddenRanges).toContain('Row 3');
    expect(metrics.hiddenRanges).toContain('Col A');
    expect(metrics.hiddenRanges).toContain('Col C');
  });

  test('detects consecutive hidden ranges', () => {
    const wb = utils.book_new();
    const ws = {
      '!ref': 'A1:C5',
      '!rows': [
        { hidden: true }, 
        { hidden: true }, 
        {},               
        { hidden: true }, 
        { hidden: true }  
      ]
    };
    
    wb.Sheets = { 'Sheet1': ws };
    wb.SheetNames = ['Sheet1'];

    const analyzer = new HiddenContentAnalyzer(wb);
    const { metrics } = analyzer.analyze(ws);

    expect(metrics.hiddenRanges).toContain('Rows 1-2');
    expect(metrics.hiddenRanges).toContain('Rows 4-5');
  });

  /* with files */
  const TEST_FILES_DIR = 'res/test_files/hidden_cell';
  describe('HiddenContentAnalyzer with Files', () => {
    test('detects hidden cells from file', () => {
      const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/hidden_cells.xlsx`);
      const wb = read(xlsxBuffer, {
        type: 'buffer',
        cellFormula: true,
        cellNF: true,
        cellText: true,
        cellStyles: true,
        cellDates: true,
        raw: true
      });
      const analyzer = new HiddenContentAnalyzer(wb);
      const { issues, metrics } = analyzer.analyze(wb.Sheets[wb.SheetNames[0]]);
  
      expect(metrics.hiddenCells).toBeGreaterThan(0);
      expect(issues).toContainEqual(
        expect.objectContaining({
          type: 'data',
          severity: 'warning',
          message: expect.stringContaining('Hidden cells detected')
        })
      );
    });
  
    test('detects hidden rows and columns from file', () => {
      const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/hidden_rows_columns.xlsx`);
      const wb = read(xlsxBuffer, {
        type: 'buffer',
        cellFormula: true,
        cellNF: true,
        cellText: true,
        cellStyles: true,
        cellDates: true,
        raw: true
      });

      wb.Sheets[wb.SheetNames[0]]['!rows'] = [
        { hidden: true },
        {},
        { hidden: true }
      ];
      wb.Sheets[wb.SheetNames[0]]['!cols'] = [
        { hidden: true },
        {},
        { hidden: true }
      ];
      
      const analyzer = new HiddenContentAnalyzer(wb);
      const { issues, metrics } = analyzer.analyze(wb.Sheets[wb.SheetNames[0]]);
  
      expect(metrics.hiddenRows).toBe(2);
      expect(metrics.hiddenColumns).toBe(2);
      expect(metrics.hiddenRanges).toContain('Row 1');
      expect(metrics.hiddenRanges).toContain('Row 3');
      expect(metrics.hiddenRanges).toContain('Col A');
      expect(metrics.hiddenRanges).toContain('Col C');
    });
  
    test('detects consecutive hidden ranges from file', () => {
      const xlsxBuffer = readFileSync(`${TEST_FILES_DIR}/consecutive_hidden_rows.xlsx`);
      const wb = read(xlsxBuffer, {
        type: 'buffer',
        cellFormula: true,
        cellNF: true,
        cellText: true,
        cellStyles: true,
        cellDates: true,
        raw: true
      });
      const analyzer = new HiddenContentAnalyzer(wb);
      const { metrics } = analyzer.analyze(wb.Sheets[wb.SheetNames[0]]);
  
      expect(metrics.hiddenRanges).toContain('Rows 1-2');
      expect(metrics.hiddenRanges).toContain('Rows 4-5');
    });
  });
});