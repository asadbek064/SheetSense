import { WorkBook, Sheet as WorkSheet } from 'xlsx';

export class ExcelParser {
  constructor(private workbook: WorkBook) {}

  getFormulaCells(sheet: WorkSheet): Map<string, string> {
    const formulas = new Map<string, string>();
    
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      
      const cell = sheet[cellAddress];
      // Check for formula property
      if (cell && typeof cell === 'object' && 'f' in cell && cell.f) {
        console.log('Found formula:', cellAddress, cell.f);
        formulas.set(cellAddress, cell.f);
      } else if (cell && typeof cell === 'object' && 'v' in cell && 
                 typeof cell.v === 'string' && cell.v.startsWith('=')) {
        // Also check for string values that look like formulas
        const formula = cell.v.substring(1); // Remove the '=' prefix
        console.log('Found string formula:', cellAddress, formula);
        formulas.set(cellAddress, formula);
      }
    }
    
    return formulas;
  }
}