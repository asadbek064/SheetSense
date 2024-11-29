import { Issue, IssueType, IssueSeverity } from '../types';
export function validateDataTypes(column: string, values: any[], hasHeader: boolean = true): Issue[] {
  if (values.length === 0) return [];

  // Get raw values and types, preserving original type information
  const rawValues = values.map(v => {
    const type = typeof v;
    return {
      value: type === 'object' && v ? v.v : v, // Handle XLSX cell objects
      type: type === 'object' && v ? v.t : type // Get XLSX type or JS type
    };
  });

  // Skip header
  const dataValues = rawValues.slice(1);
  if (dataValues.length === 0) return [];

  // Check for binary/indicator columns (0/1)
  const uniqueValues = new Set(dataValues.map(v => String(v.value).trim()));
  if (uniqueValues.size <= 2 && [...uniqueValues].every(v => ['0', '1'].includes(v))) {
    return [];
  }

  // Check for numeric column
  const isNumericColumn = dataValues.every(v => {
    if (v.type === 'n') return true; // Excel number type
    if (typeof v.value === 'number') return true;
    const strVal = String(v.value).trim().replace(/,/g, '');
    return !isNaN(Number(strVal)) && strVal !== '';
  });

  if (isNumericColumn) return [];

  // If we get here, check for mixed types
  const types = new Set(dataValues.map(v => v.type));
  if (types.size > 1) {
    return [{
      type: 'data',
      severity: 'warning',
      message: `Mixed data types in column ${column}`,
      cell: `${column}2:${column}${values.length}`,
      sheet: 'Sheet1',
      suggestion: 'Convert to consistent data type'
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