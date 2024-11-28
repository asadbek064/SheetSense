import { utils, WorkBook } from 'xlsx';

export function createTestWorkbook(cells: Record<string, { v: any, t: string, f?: string }>): WorkBook {
  const wb = utils.book_new();
  const ws = {
    ...cells,
    '!ref': `A1:${getMaxRef(Object.keys(cells))}`
  };
  
  wb.Sheets = { 'Sheet1': ws };
  wb.SheetNames = ['Sheet1'];
  
  return wb;
}

function getMaxRef(addresses: string[]): string {
  let maxCol = 'A';
  let maxRow = 1;
  
  for (const addr of addresses) {
    const match = addr.match(/([A-Z]+)(\d+)/);
    if (match) {
      const [, col, row] = match;
      if (col > maxCol) maxCol = col;
      if (+row > maxRow) maxRow = +row;
    }
  }
  
  return `${maxCol}${maxRow}`;
}
