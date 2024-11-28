import { Issue } from '../types/Issue';
import { ExcelParser } from './Parser';
import { ExcelFormulaParser } from './formula/FormulaParser';

export class FormulaAnalyzer {
  private formulaParser: ExcelFormulaParser;

  constructor(private parser: ExcelParser) {
    this.formulaParser = new ExcelFormulaParser();
  }

  analyzeFormula(formula: string | undefined, cell: string, sheet: string): Issue[] {
    if (!formula) return [];
    
    try {
      const analysis = this.formulaParser.analyze(formula);
      const issues: Issue[] = [];

      // Convert formula analysis to standard issues
      
      // Handle validation errors
      if (!analysis.isValid) {
        analysis.errors.forEach(error => {
          issues.push({
            type: 'formula',
            severity: 'error',
            message: `Formula error: ${error}`,
            cell,
            sheet,
            suggestion: 'Review and correct formula syntax'
          });
        });
      }

      // Handle complexity warnings
      if (analysis.complexity > 5) {
        issues.push({
          type: 'formula',
          severity: 'warning',
          message: 'Complex formula detected',
          cell,
          sheet,
          suggestion: 'Consider breaking down into smaller parts or using helper cells'
        });
      }

      // Handle volatile functions
      if (analysis.volatileFunctions.length > 0) {
        issues.push({
          type: 'formula',
          severity: 'warning',
          message: `Volatile functions detected: ${analysis.volatileFunctions.join(', ')}`,
          cell,
          sheet,
          suggestion: 'Consider using non-volatile alternatives or caching results'
        });
      }

      // Handle large range references
      const largeRanges = analysis.usedRanges.filter(range => {
        const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const [, startCol, startRow, endCol, endRow] = match;
          const rowDiff = parseInt(endRow) - parseInt(startRow);
          return rowDiff > 1000; // Threshold for large ranges
        }
        return false;
      });

      if (largeRanges.length > 0) {
        issues.push({
          type: 'formula',
          severity: 'warning',
          message: 'Large range references detected',
          cell,
          sheet,
          suggestion: 'Consider using more efficient range references or database functions'
        });
      }

      // Handle circular references
      if (analysis.dependencies.includes(cell)) {
        issues.push({
          type: 'formula',
          severity: 'error',
          message: 'Circular reference detected',
          cell,
          sheet,
          suggestion: 'Remove circular reference to prevent calculation errors'
        });
      }

      // Handle nested IF statements
      const nestedIfs = formula.match(/IF\([^)]*IF\([^)]*\)[^)]*\)/g);
      if (nestedIfs && nestedIfs.length > 0) {
        issues.push({
          type: 'formula',
          severity: 'warning',
          message: 'Nested IF statements detected',
          cell,
          sheet,
          suggestion: 'Consider using SWITCH or IFS functions for better readability'
        });
      }

      // Handle error prone patterns
      if (formula.includes('/0') || formula.includes('/ 0')) {
        issues.push({
          type: 'formula',
          severity: 'error',
          message: 'Potential division by zero',
          cell,
          sheet,
          suggestion: 'Add error handling using IFERROR or IF functions'
        });
      }

      return issues;

    } catch (error) {
      return [{
        type: 'formula',
        severity: 'error',
        message: `Formula parsing error: ${error}`,
        cell,
        sheet,
        suggestion: 'Check formula syntax and structure'
      }];
    }
  }

  /**
   * Checks if a formula is overly complex
   */
  private isComplexFormula(formula: string): boolean {
    const analysis = this.formulaParser.analyze(formula);
    return analysis.complexity > 5;
  }

  /**
   * Checks if a formula contains potential performance issues
   */
  private hasPerformanceIssues(formula: string): boolean {
    const analysis = this.formulaParser.analyze(formula);
    return analysis.volatileFunctions.length > 0 || 
           analysis.usedRanges.some(range => this.isLargeRange(range));
  }

  /**
   * Checks if a range reference is considered large
   */
  private isLargeRange(range: string): boolean {
    const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (match) {
      const [, startCol, startRow, endCol, endRow] = match;
      const rowDiff = parseInt(endRow) - parseInt(startRow);
      return rowDiff > 1000;
    }
    return false;
  }

  /**
   * Checks for error-prone patterns in formulas
   */
  private hasErrorPronePatterns(formula: string): boolean {
    const errorPatterns = [
      /\/\s*0/,                     // Division by zero
      /IF\([^,]*,[^,]*\)$/,        // IF missing else clause
      /VLOOKUP\([^,]*,[^,]*,[^,]*\)$/, // VLOOKUP missing exact match parameter
      /\b(AND|OR)\([^,]*\)/        // Single argument logical functions
    ];

    return errorPatterns.some(pattern => pattern.test(formula));
  }
}