import { createTestWorkbook } from '../helpers/createTestSheet';

export const dataQualityTestData = {
  'A1': { v: 'ID', t: 's' },
  'A2': { v: 1, t: 'n' },
  'A3': { v: '2', t: 's' },
  'B1': { v: 'Date', t: 's' },
  'B2': { v: '2024-01-01', t: 's' },
  'B3': { v: 'invalid', t: 's' }
};

export const formattingTestData = {
  'A1': { v: 'Amount', t: 's' },
  'A2': { v: '1000.00', t: 's' },
  'A3': { v: '1,000', t: 's' },
  'A4': { v: ' 1000.0 ', t: 's' }
};