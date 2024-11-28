import { Issue, IssueType, IssueSeverity } from '../types';

export function validateDataTypes(column: string, values: any[]): Issue[] {
  const types = new Set(values.map(v => typeof v));
  
  if (types.size > 1) {
    return [{
      type: 'data' as IssueType,
      severity: 'warning' as IssueSeverity,
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
}