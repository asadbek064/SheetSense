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
}