import { describe, test, expect } from 'vitest';
import { validationFormula } from '../../src/rules/FormulaRules';

describe('FormulaRules', () => {
  test('detects division by zero', () => {
    const issues = validationFormula('A1/0', 'B2', 'Sheet1');
    expect(issues[0]).toMatchObject({
      type: 'formula',
      severity: 'error',
      message: expect.stringContaining('division by zero')
    });
  });

  test('detects complex nested formulas', () => {
    const formula = 'IF(A1>0,IF(A1<10,IF(A1<5,"Low","Medium"),"High"))';
    const issues = validationFormula(formula, 'B2', 'Sheet1');
    expect(issues[0]).toMatchObject({
      type: 'formula',
      severity: 'warning',
      message: expect.stringContaining('Complex formula')
    });
  });

  test('accepts simple formulas', () => {
    const issues = validationFormula('A1+B1*2', 'C1', 'Sheet1');
    expect(issues.length).toBe(0);
  });
});import { test, expect } from 'vitest';
import { validateDataTypes, validateDates } from '../../src/rules/DataRules';

test('validates data types in column', () => {
  const values = [1, '2', true, '4'];
  const issues = validateDataTypes('A', values);
  
  expect(issues.length).toBe(1);
  expect(issues[0]).toMatchObject({
    type: 'data',
    severity: 'warning',
    message: 'Mixed data types in column A'
  });
});

test('validates date formats', () => {
  const dates = ['2024-01-01', '2024-13-32', 'invalid'];
  const issues = validateDates('A', dates);

  expect(issues.length).toBe(2);
  expect(issues[0]).toMatchObject({
    type: 'data',
    severity: 'error',
    message: expect.stringContaining('Invalid date format')
  });
});import { test, expect } from 'vitest';
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

  console.log('Created workbook:', JSON.stringify(wb, null, 2));
  
  const analyzer = new ExcelAnalyzer(wb);
  const analysis = analyzer.analyze();
  
  console.log('Analysis result:', JSON.stringify(analysis, null, 2));
  
  expect(analysis.issues.length).toBeGreaterThan(0);
  expect(analysis.issues).toContainEqual(
    expect.objectContaining({
      type: 'formula',
      severity: 'error',
      message: expect.stringContaining('division by zero')
    })
  );
});
import { utils, WorkBook } from 'xlsx';

export function createTestWorkbook(cells: Record<string, { v: any, t: string, f?: string }>): WorkBook {
  const wb = utils.book_new();
  const ws = {
    ...cells,
    '!ref': `A1:${getMaxRef(Object.keys(cells))}`
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];
  
  return wb;
}

function getMaxRef(addresses: string[]): string {
  let maxCol = 'A';
  let maxRow = 1;
  
  for (const addr of addresses) {
    const match = addr.match(/([A-Z]+)(\d+)/);
    if (match) {
      const [, col, row] = match;
      if (col > maxCol) maxCol = col;
      if (+row > maxRow) maxRow = +row;
    }
  }
  
  return `${maxCol}${maxRow}`;
}
import { createTestWorkbook } from '../helpers/createTestSheet';

export const dataQualityTestData = {
  'A1': { v: 'ID', t: 's' },
  'A2': { v: 1, t: 'n' },
  'A3': { v: '2', t: 's' },
  'B1': { v: 'Date', t: 's' },
  'B2': { v: '2024-01-01', t: 's' },
  'B3': { v: 'invalid', t: 's' }
};

export const formattingTestData = {
  'A1': { v: 'Amount', t: 's' },
  'A2': { v: '1000.00', t: 's' },
  'A3': { v: '1,000', t: 's' },
  'A4': { v: ' 1000.0 ', t: 's' }
};import { test, expect } from 'vitest';
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
});import { test, expect } from 'vitest';
import { utils } from 'xlsx';
import { MetadataCollector } from '../../src/core/MetadataCollector';

test('collects basic workbook metadata', () => {
  // Create test workbook
  const wb = utils.book_new();
  const ws1 = {
    '!ref': 'A1:C3',
    'A1': { v: 'Header', t: 's' },
    'B1': { v: 'Value', t: 's' },
    'C1': { v: 'Formula', t: 's' },
    'A2': { v: 1, t: 'n' },
    'B2': { v: 2, t: 'n' },
    'C2': { v: 3, t: 'n', f: 'A2+B2' }
  };

  const ws2 = {
    '!ref': 'A1:B2',
    'A1': { v: 'Test', t: 's' },
    'B1': { v: 10, t: 'n', f: 'A1*2' }
  };

  utils.book_append_sheet(wb, ws1, 'Sheet1');
  utils.book_append_sheet(wb, ws2, 'Sheet2');

  // Add named range
  if (!wb.Workbook) wb.Workbook = { Names: [] };
  wb.Workbook.Names = [{ Name: 'TestRange', Ref: 'Sheet1!A1:C3' }];

  const collector = new MetadataCollector(wb);
  const metadata = collector.collect();

  // Test basic counts
  expect(metadata.sheetCount).toBe(2);
  expect(metadata.formulaCount).toBe(2); // Two formulas total
  expect(metadata.namedRanges).toContain('TestRange');

  // Test metrics
  expect(metadata.metrics.totalCells).toBeGreaterThan(0);
  expect(metadata.metrics.dataCells).toBeGreaterThan(0);
  expect(metadata.metrics.percentagePopulated).toBeGreaterThan(0);
});

test('handles empty workbook', () => {
  const wb = utils.book_new();
  utils.book_append_sheet(wb, { '!ref': 'A1:A1' }, 'EmptySheet');

  const collector = new MetadataCollector(wb);
  const metadata = collector.collect();

  expect(metadata.sheetCount).toBe(1);
  expect(metadata.formulaCount).toBe(0);
  expect(metadata.namedRanges).toHaveLength(0);
  expect(metadata.metrics.totalCells).toBe(0);
});

test('calculates correct metrics for mixed content', () => {
  const wb = utils.book_new();
  const ws = {
    '!ref': 'A1:B3',
    'A1': { v: 'Header', t: 's' },
    'B1': { v: '', t: 's' },  // Empty cell
    'A2': { v: 1, t: 'n' },
    'B2': { v: 2, t: 'n', f: 'A2*2' }, // Formula cell
  };

  utils.book_append_sheet(wb, ws, 'Sheet1');

  const collector = new MetadataCollector(wb);
  const metadata = collector.collect();

  expect(metadata.formulaCount).toBe(1);
  expect(metadata.metrics.dataCells).toBe(3); // Header, A2, B2
  expect(metadata.metrics.emptyCells).toBe(1); // B1
  expect(metadata.metrics.totalCells).toBe(4);
  expect(metadata.metrics.percentagePopulated).toBe(75); // 3/4 cells populated
});import { test, expect } from 'vitest';
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
});import { test, expect } from 'vitest';
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
});import { test, expect } from 'vitest';
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

  console.log('Found formulas:', Array.from(formulas.entries()));
  
  expect(formulas.size).toBeGreaterThan(0);
  expect(formulas.get('B2')).toBe('A1*2');
});import { defineConfig } from 'vite';
import path from 'path';
import { configDefaults } from 'vitest/config';

export default defineConfig({
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  },
  build: {
    lib: {
      entry: path.resolve(__dirname, 'src/index.ts'),
      name: 'SheetSense',
      fileName: (format) => `sheetsense.${format}.js`,
      formats: ['es', 'umd']
    },
    sourcemap: true,
    outDir: 'dist',
    emptyOutDir: true
  },
  test: {
    globals: true,
    environment: 'node',
    exclude: [...configDefaults.exclude, 'e2e/*'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html']
    }
  }
});export { ExcelAnalyzer } from './core/Analyzer';
export { ExcelParser } from './core/Parser';
export { FormulaAnalyzer } from './core/FormulaAnalyzer';
export { StyleAnalyzer } from './core/StyleAnalyzer';
export * from './types/Analysis';export * from './Analysis';
export * from './Issue';import { z } from 'zod';
import { IssueSchema } from './Issue';

export const AnalysisSchema = z.object({
  issues: z.array(IssueSchema),
  metadata: z.object({
    formulaCount: z.number(),
    sheetCount: z.number(),
    namedRanges: z.array(z.string()),
    volatileFunctions: z.number(),
    externalReferences: z.number(),
  }),
});

export type Analysis = z.infer<typeof AnalysisSchema>;import { z } from 'zod';

export const IssueType = z.enum(['formula', 'style', 'data']);
export type IssueType = z.infer<typeof IssueType>;

export const IssueSeverity = z.enum(['error', 'warning', 'info']);
export type IssueSeverity = z.infer<typeof IssueSeverity>;

export const IssueSchema = z.object({
  type: IssueType,
  severity: IssueSeverity,
  cell: z.string(),
  sheet: z.string(),
  message: z.string(),
  suggestion: z.string().optional(),
});

export type Issue = z.infer<typeof IssueSchema>;import { Issue, IssueType, IssueSeverity } from '../types';

export function validateDataTypes(column: string, values: any[], hasHeader: boolean = true): Issue[] {
  const dataValues = values.slice(hasHeader ? 1 : 0);  // Skip header if present
  if (dataValues.length === 0) return [];

  const types = new Set(dataValues.map(v => typeof v));
  
  if (types.size > 1) {
    return [{
      type: 'data',
      severity: 'warning',
      message: `Mixed data types in column ${column}`,
      cell: `${column}1:${column}${values.length}`,
      sheet: 'Sheet1',
      suggestion: 'Consider converting to consistent data type'
    }];
  }
  
  return [];
}


export function validateDates(column: string, dates: string[]): Issue[] {
  const issues: Issue[] = [];
  
  dates.forEach((date, index) => {
    const parsed = new Date(date);
    if (isNaN(parsed.getTime())) {
      issues.push({
        type: 'data' as IssueType,
        severity: 'error' as IssueSeverity,
        message: `Invalid date format at ${column}${index + 1}`,
        cell: `${column}${index + 1}`,
        sheet: 'Sheet1',
        suggestion: 'Use YYYY-MM-DD format'
      });
    }
  });
  
  return issues;
}import { Issue } from "@/types/Issue";

export function validationFormula(formula: string, cell: string, sheet: string): Issue[] {
    const issues: Issue[] = [];

    if (!formula) return issues;

    if (isComplexFormula(formula)) {
        issues.push({
          type: 'formula',
          severity: 'warning',
          message: 'Complex formula detected',
          cell,
          sheet,
          suggestion: 'Consider breaking this formula into smaller parts'
        });
      }
    
      if (hasZeroDivision(formula)) {
        issues.push({
          type: 'formula',
          severity: 'error',
          message: 'Possible division by zero',
          cell,
          sheet,
          suggestion: 'Add error handling with IFERROR()'
        });
      }
    
      return issues;
} 

function isComplexFormula(formula: string): boolean {
    const nestedFunctionCount = (formula.match(/[A-Z]+\(/g) || []).length;
    const ifCount = (formula.match(/IF\(/g) || []).length;
    return nestedFunctionCount > 2 || ifCount > 1;
}

function hasZeroDivision(formula: string): boolean {
    return /\/\s*(0|[A-Z]+[0-9]+|\([^)]*\))/.test(formula);
}import { WorkBook, Sheet as WorkSheet } from 'xlsx';
import { Analysis, Issue } from '../types';
import { FormulaAnalyzer } from './FormulaAnalyzer';
import { StyleAnalyzer } from './StyleAnalyzer';
import { ExcelParser } from './Parser';
export class ExcelAnalyzer {
  private parser: ExcelParser;
  private formulaAnalyzer: FormulaAnalyzer;
  private styleAnalyzer: StyleAnalyzer;

  constructor(private workbook: WorkBook) {
    this.parser = new ExcelParser(workbook);
    this.formulaAnalyzer = new FormulaAnalyzer(this.parser);
    this.styleAnalyzer = new StyleAnalyzer(this.parser);
  }

  analyze(): Analysis {
    const issues: Issue[] = [];
    const metadata = {
      formulaCount: 0,
      sheetCount: this.workbook?.SheetNames?.length || 0,
      namedRanges: [],
      volatileFunctions: 0,
      externalReferences: 0
    };

    for (const sheetName of this.workbook?.SheetNames || []) {
      const sheet = this.workbook.Sheets[sheetName];
      const columns = this.extractColumns(sheet);
      
      // Data quality checks
      for (const [col, cells] of columns.entries()) {
        // Data type analysis
        const types = new Set(cells.map(c => c.t));
        if (types.size > 1) {
          issues.push(this.createDataTypeIssue(col, sheetName));
        }

        // Date validation
        const dateCells = cells.filter(c => this.mightBeDate(c.v));
        if (dateCells.length > 0) {
          const invalidDates = dateCells.filter(c => !this.isValidDate(c.v));
          if (invalidDates.length > 0) {
            issues.push(this.createDateIssue(col, sheetName));
          }
        }

        // Style checks
        const styleIssues = this.styleAnalyzer.analyzeNumberFormatting(
          cells.map(c => String(c.v)),
          col,
          sheetName
        );
        issues.push(...styleIssues);
      }

      // Formula checks
      const formulaCells = this.parser.getFormulaCells(sheet);
      metadata.formulaCount += formulaCells.size;

      for (const [cell, formula] of formulaCells) {
        if (this.isComplexFormula(formula)) {
          issues.push({
            type: 'formula',
            severity: 'warning',
            cell,
            sheet: sheetName,
            message: 'Complex formula detected',
            suggestion: 'Consider breaking this formula into smaller parts'
          });
        }

        if (this.hasZeroDivision(formula)) {
          issues.push({
            type: 'formula',
            severity: 'error',
            cell,
            sheet: sheetName,
            message: 'Possible division by zero',
            suggestion: 'Add error handling with IFERROR()'
          });
        }
      }
    }

    return { issues, metadata };
  }

  private getColumnValues(sheet: WorkSheet): Map<string, any[]> {
    const columns = new Map<string, any[]>();
    const ref = sheet['!ref'];
    if (!ref) return columns;

    for (const cell in sheet) {
      if (cell[0] === '!') continue;
      const col = cell.match(/[A-Z]+/)?.[0];
      if (col) {
        if (!columns.has(col)) columns.set(col, []);
        columns.get(col)?.push(sheet[cell].v);
      }
    }

    return columns;
  }

  private isDateColumn(values: any[]): boolean {
    return values.some(v => typeof v === 'string' && 
      (v.includes('-') || v.includes('/') || v.includes('.')));
  }

  private isComplexFormula(formula: string): boolean {
    const nestedFunctionCount = (formula.match(/[A-Z]+\(/g) || []).length;
    const operatorCount = (formula.match(/[\+\-\*\/\^]/g) || []).length;
    const ifCount = (formula.match(/IF\(/g) || []).length;
    const complexityScore = nestedFunctionCount + (operatorCount * 0.5) + (ifCount * 2);
    return complexityScore > 3;
  }

  private hasZeroDivision(formula: string): boolean {
    return /\/\s*(0|[A-Z]+[0-9]+|\([^)]*\))/.test(formula) && !formula.includes('IFERROR(');
  }


  private extractColumns(sheet: WorkSheet): Map<string, any[]> {
    const columns = new Map<string, any[]>();
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      const col = cellAddress.match(/[A-Z]+/)?.[0];
      if (col) {
        if (!columns.has(col)) columns.set(col, []);
        columns.get(col)?.push(sheet[cellAddress]);
      }
    }
    return columns;
  }

  private mightBeDate(value: any): boolean {
    return typeof value === 'string' && 
           /(\d{4}-\d{2}-\d{2}|invalid date|\d{2}\/\d{2}\/\d{4})/.test(value);
  }

  private isValidDate(value: any): boolean {
    if (typeof value !== 'string') return false;
    const date = new Date(value);
    return date.toString() !== 'Invalid Date';
  }

  private createDataTypeIssue(col: string, sheet: string): Issue {
    return {
      type: 'data',
      severity: 'warning',
      message: `Mixed data types in column ${col}`,
      cell: col,
      sheet,
      suggestion: 'Convert to consistent data type'
    };
  }

  private createDateIssue(col: string, sheet: string): Issue {
    return {
      type: 'data',
      severity: 'error',
      message: `Invalid date format in column ${col}`,
      cell: col,
      sheet,
      suggestion: 'Use YYYY-MM-DD format'
    };
  }
}import { Issue } from '../types';
import { ExcelParser } from './Parser';

export class FormulaAnalyzer {
  constructor(private parser: ExcelParser) {}

  analyzeFormula(formula: string | undefined, cell: string, sheet: string): Issue[] {
    if (!formula) return [];
    
    const issues: Issue[] = [];

    if (this.isComplexFormula(formula)) {
      issues.push({
        type: 'formula',
        severity: 'warning',
        cell,
        sheet,
        message: 'Complex formula detected',
        suggestion: 'Consider breaking this formula into smaller parts'
      });
    }

    return issues;
  }

  private isComplexFormula(formula: string): boolean {
    return this.calculateComplexity(formula) > 5;
  }

  private calculateComplexity(formula: string): number {
    let complexity = 0;
    
    // Count nested functions
    complexity += (formula.match(/[A-Z]+\(/g) || []).length;
    
    // Count operators
    complexity += (formula.match(/[\+\-\*\/\^]/g) || []).length * 0.5;
    
    // Count IF statements
    complexity += (formula.match(/IF\(/g) || []).length * 2;
    
    return complexity;
  }
}

import { Issue } from '../types';
import { ExcelParser } from './Parser';

export class StyleAnalyzer {
  analyzeNumberFormatting(values: string[], column: string, sheet: string, hasHeader: boolean = true): Issue[] {
    // Filter out header and non-numeric strings
    const numberValues = values
      .slice(hasHeader ? 1 : 0)  // Skip header if present
      .filter(v => this.isNumeric(v));

    if (numberValues.length === 0) return [];

    const numberFormats = new Set(
      numberValues.map(v => this.detectNumberFormat(v))
    );

    if (numberFormats.size > 1) {
      return [{
        type: 'style',
        severity: 'warning',
        message: `Inconsistent number format in column ${column}`,
        cell: column,
        sheet,
        suggestion: 'Use consistent number formatting'
      }];
    }

    return [];
  }

  private isNumeric(value: string): boolean {
    return /^[\d,.]+(\.?\d+)?$/.test(value.trim());
  }
  
  private detectNumberFormat(value: string): string {
    if (/^\d+\.\d{2}$/.test(value)) return 'X.XX';
    if (/^\d{1,3}(,\d{3})*$/.test(value)) return 'X,XXX';
    if (/^\d+\.\d+$/.test(value)) return 'X.X';
    return 'other';
  }
}import { WorkBook, Sheet as WorkSheet } from 'xlsx';

export class ExcelParser {
  constructor(private workbook: WorkBook) {}

  getFormulaCells(sheet: WorkSheet): Map<string, string> {
    const formulas = new Map<string, string>();
    
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      
      const cell = sheet[cellAddress];
      // Check for formula property
      if (cell && typeof cell === 'object' && 'f' in cell && cell.f) {
        console.log('Found formula:', cellAddress, cell.f);
        formulas.set(cellAddress, cell.f);
      } else if (cell && typeof cell === 'object' && 'v' in cell && 
                 typeof cell.v === 'string' && cell.v.startsWith('=')) {
        // Also check for string values that look like formulas
        const formula = cell.v.substring(1); // Remove the '=' prefix
        console.log('Found string formula:', cellAddress, formula);
        formulas.set(cellAddress, formula);
      }
    }
    
    return formulas;
  }
}import { WorkBook, Sheet as WorkSheet } from 'xlsx';

export interface WorkbookMetadata {
  formulaCount: number;
  sheetCount: number;
  namedRanges: string[];
  metrics: {
    totalCells: number;
    dataCells: number;
    emptyCells: number;
    percentagePopulated: number;
  };
}

export class MetadataCollector {
  constructor(private workbook: WorkBook) {}

  collect(): WorkbookMetadata {
    const metadata: WorkbookMetadata = {
      formulaCount: 0,
      sheetCount: this.workbook.SheetNames.length,
      namedRanges: this.getNamedRanges(),
      metrics: {
        totalCells: 0,
        dataCells: 0,
        emptyCells: 0,
        percentagePopulated: 0
      }
    };

    // Process each sheet
    for (const sheetName of this.workbook.SheetNames) {
      const sheet = this.workbook.Sheets[sheetName];
      metadata.formulaCount += this.countFormulas(sheet);
      this.collectSheetMetrics(sheet, metadata.metrics);
    }

    // Calculate percentage populated
    metadata.metrics.percentagePopulated = 
      (metadata.metrics.dataCells / metadata.metrics.totalCells) * 100;

    return metadata;
  }

  private countFormulas(sheet: WorkSheet): number {
    let count = 0;
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      const cell = sheet[cellAddress];
      if (cell && typeof cell === 'object' && 'f' in cell) {
        count++;
      }
    }
    return count;
  }

  private getNamedRanges(): string[] {
    const namedRanges: string[] = [];
    if (this.workbook.Workbook?.Names) {
      for (const name of this.workbook.Workbook.Names) {
        namedRanges.push(name.Name);
      }
    }
    return namedRanges;
  }

  private collectSheetMetrics(sheet: WorkSheet, metrics: WorkbookMetadata['metrics']) {
    const range = sheet['!ref'];
    if (!range) return;

    // Count cells
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      metrics.totalCells++;
      
      const cell = sheet[cellAddress];
      if (cell && 'v' in cell && cell.v !== undefined && cell.v !== '') {
        metrics.dataCells++;
      } else {
        metrics.emptyCells++;
      }
    }
  }
}