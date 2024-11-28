import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { MetadataCollector } from '../../src/core/MetadataCollector';

test('collects complete workbook metadata with all features', () => {
  // Create test workbook with multiple sheets and varied content
  const wb = utils.book_new();
  
  // Sheet 1 with formulas and mixed content
  const ws1 = {
    '!ref': 'A1:C3',
    'A1': { v: 'Header', t: 's' },
    'B1': { v: 'Value', t: 's' },
    'C1': { v: 'Formula', t: 's' },
    'A2': { v: 1, t: 'n' },
    'B2': { v: 2, t: 'n' },
    'C2': { v: 3, t: 'n', f: 'A2+B2' },
    'A3': { v: '', t: 's' },  // Empty cell
    'B3': { v: 4, t: 'n' },
    'C3': { v: 8, t: 'n', f: 'B3*2' }  // Another formula
  };

  // Sheet 2 with different content
  const ws2 = {
    '!ref': 'A1:B2',
    'A1': { v: 'Test', t: 's' },
    'B1': { v: 10, t: 'n', f: 'A1*2' }  // Formula in second sheet
  };

  utils.book_append_sheet(wb, ws1, 'Sheet1');
  utils.book_append_sheet(wb, ws2, 'Sheet2');

  // Add named ranges
  if (!wb.Workbook) wb.Workbook = { Names: [] };
  wb.Workbook.Names = [
    { Name: 'TestRange1', Ref: 'Sheet1!A1:C3' },
    { Name: 'TestRange2', Ref: 'Sheet2!A1:B2' }
  ];

  const collector = new MetadataCollector(wb);
  const metadata = collector.collect();

  // Test formula count
  expect(metadata.formulaCount).toBe(3);
  expect(metadata.formulaCount).toBeGreaterThan(0);
  
  // Test sheet count
  expect(metadata.sheetCount).toBe(2);
  
  // Test named ranges
  expect(metadata.namedRanges).toHaveLength(2);
  expect(metadata.namedRanges).toContain('TestRange1');
  expect(metadata.namedRanges).toContain('TestRange2');
  
  // Test basic metrics
  expect(metadata.metrics).toBeDefined();
  expect(metadata.metrics.totalCells).toBe(11); // Total cells across both sheets
  expect(metadata.metrics.dataCells).toBe(10); // Cells with actual data
  expect(metadata.metrics.emptyCells).toBe(1); // Empty cells
  expect(metadata.metrics.percentagePopulated).toBeCloseTo(90.909, 3); // (10/11) * 100
});

test('handles edge cases in metadata collection', () => {
  const wb = utils.book_new();
  
  // Empty sheet
  const wsEmpty = { '!ref': 'A1:A1' };
  
  // Sheet with only formulas
  const wsFormulas = {
    '!ref': 'A1:A3',
    'A1': { v: 1, t: 'n', f: '1+1' },
    'A2': { v: 2, t: 'n', f: '2+2' },
    'A3': { v: 3, t: 'n', f: '3+3' }
  };
  
  utils.book_append_sheet(wb, wsEmpty, 'EmptySheet');
  utils.book_append_sheet(wb, wsFormulas, 'FormulasSheet');
  
  const collector = new MetadataCollector(wb);
  const metadata = collector.collect();
  
  // Test edge cases
  expect(metadata.sheetCount).toBe(2);
  expect(metadata.formulaCount).toBe(3); // Only formula cells
  expect(metadata.namedRanges).toHaveLength(0); // No named ranges
  expect(metadata.metrics.totalCells).toBe(3);
  expect(metadata.metrics.dataCells).toBe(3);
  expect(metadata.metrics.emptyCells).toBe(0);
  expect(metadata.metrics.percentagePopulated).toBe(100);
});