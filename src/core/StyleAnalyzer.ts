import { Issue } from '../types';
import { ExcelParser } from './Parser';

export class StyleAnalyzer {
  constructor(private parser: ExcelParser) {}

  analyzeNumberFormatting(values: string[], column: string, sheet: string): Issue[] {
    const numberFormats = new Set(
      values
        .filter(v => typeof v === 'string')
        .map(v => this.detectNumberFormat(v))
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

  private detectNumberFormat(value: string): string {
    if (/^\d+\.\d{2}$/.test(value)) return 'X.XX';
    if (/^\d{1,3}(,\d{3})*$/.test(value)) return 'X,XXX';
    if (/^\d+\.\d+$/.test(value)) return 'X.X';
    return 'other';
  }
}