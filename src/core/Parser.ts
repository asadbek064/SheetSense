import { WorkBook, Sheet as WorkSheet } from 'xlsx';

export class ExcelParser {
  constructor(private workbook: WorkBook) {}

  getFormulaCells(sheet: WorkSheet): Map<string, string> {
    const formulas = new Map<string, string>();
    
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      
      const cell = sheet[cellAddress];
      if (!cell) continue;

      let formula = '';
      
      // Handle explicit formula field
      if ('f' in cell && cell.f) {
        formula = cell.f;
        formulas.set(cellAddress, formula);
      } 
      // Handle string-type cells that might be formulas
      else if (cell.t === 'str' && cell.v && typeof cell.v === 'string') {
        const value = cell.v.trim();
        if (value.startsWith('=')) {
          formula = value.substring(1);
          formulas.set(cellAddress, formula);
        }
      }
    }

    return formulas;
  }
}