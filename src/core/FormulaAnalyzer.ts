import { Issue } from '../types';
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

