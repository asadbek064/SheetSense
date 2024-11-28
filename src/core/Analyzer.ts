import { WorkBook, Sheet as WorkSheet } from 'xlsx';
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
    this.styleAnalyzer = new StyleAnalyzer();
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
}