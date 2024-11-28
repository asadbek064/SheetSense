import { Issue } from "@/types/Issue";

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
}