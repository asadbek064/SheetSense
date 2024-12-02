import { WorkBook, Sheet as WorkSheet } from 'xlsx';
import { Issue } from '../types';

// convert number to column letter
function getColumnLetter(index: number): string {
    let column = '';
    while (index >= 0) {
      column = String.fromCharCode((index % 26) + 65) + column;
      index = Math.floor(index / 26) - 1;
    }
    return column;
  }

export interface HiddenContentMetrics {
  hiddenCells: number;
  hiddenRows: number;
  hiddenColumns: number;
  hiddenRanges: string[];
}

export class HiddenContentAnalyzer {
  constructor(private workbook: WorkBook) {}

  analyze(sheet: WorkSheet): { issues: Issue[]; metrics: HiddenContentMetrics } {
    const issues: Issue[] = [];
    const metrics: HiddenContentMetrics = {
      hiddenCells: 0,
      hiddenRows: 0,
      hiddenColumns: 0,
      hiddenRanges: []
    };
    
    // Check for hidden rows and columns
    if (sheet['!rows']) {
      const hiddenRows = this.findHiddenRanges(sheet['!rows'], 'row');
      metrics.hiddenRows = hiddenRows.length;
      metrics.hiddenRanges.push(...hiddenRows);

      if (hiddenRows.length > 0) {
        issues.push({
          type: 'data',
          severity: 'warning',
          message: `Hidden rows detected: ${hiddenRows.join(', ')}`,
          cell: 'Sheet',
          sheet: this.getSheetName(sheet),
          suggestion: 'Review hidden rows for sensitive data'
        });
      }
    }

    if (sheet['!cols']) {
      const hiddenCols = this.findHiddenRanges(sheet['!cols'], 'column');
      metrics.hiddenColumns = hiddenCols.length;
      metrics.hiddenRanges.push(...hiddenCols);

      if (hiddenCols.length > 0) {
        issues.push({
          type: 'data',
          severity: 'warning',
          message: `Hidden columns detected: ${hiddenCols.join(', ')}`,
          cell: 'Sheet',
          sheet: this.getSheetName(sheet),
          suggestion: 'Review hidden columns for sensitive data'
        });
      }
    }

    // Check for individually hidden cells
    const hiddenCells = this.findHiddenCells(sheet);
    metrics.hiddenCells = hiddenCells.length;

    if (hiddenCells.length > 0) {
      issues.push({
        type: 'data',
        severity: 'warning',
        message: `Hidden cells detected: ${hiddenCells.join(', ')}`,
        cell: hiddenCells.join(', '),
        sheet: this.getSheetName(sheet),
        suggestion: 'Review hidden cells for sensitive data'
      });
    }

    return { issues, metrics };
  }

  private findHiddenRanges(items: any[], type: 'row' | 'column'): string[] {
    const ranges: string[] = [];
    let start: number | null = null;
    let prev: number | null = null;
  
    items.forEach((item, index) => {
      const current = index + 1;
      
      if (item?.hidden) {
        if (start === null) {
          start = current;
        } else if (prev !== null && current > prev + 1) {
          // Gap detected, close current range
          ranges.push(this.formatRange(start, prev, type));
          start = current;
        }
        prev = current;
      } else if (start !== null) {
        ranges.push(this.formatRange(start, prev!, type));
        start = null;
        prev = null;
      }
    });

    if (start !== null) {
      ranges.push(this.formatRange(start, prev!, type));
    }

    return ranges;
  }


  private formatRange(start: number, end: number, type: 'row' | 'column'): string {
    if (type === 'column') {
      const startCol = getColumnLetter(start - 1);
      if (start === end) {
        return `Col ${startCol}`;
      }
      const endCol = getColumnLetter(end - 1);
      return `Cols ${startCol}-${endCol}`;
    } else {
      if (start === end) {
        return `Row ${start}`;
      }
      return `Rows ${start}-${end}`;
    }
  }

  private findHiddenCells(sheet: WorkSheet): string[] {
    const hiddenCells: string[] = [];

    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;

      const cell = sheet[cellAddress];
      if (cell?.h) {
        hiddenCells.push(cellAddress);
      }
    }

    return hiddenCells;
  }

  private getSheetName(sheet: WorkSheet): string {
    return this.workbook.SheetNames.find(name => 
      this.workbook.Sheets[name] === sheet) || 'Unknown';
  }
}