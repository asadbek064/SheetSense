import { test, expect } from 'vitest';
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
});