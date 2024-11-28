import { WorkBook, Sheet as WorkSheet } from 'xlsx';

export interface WorkbookMetadata {
  formulaCount: number;
  sheetCount: number;
  namedRanges: string[];
  metrics: {
    totalCells: number;
    dataCells: number;
    emptyCells: number;
    percentagePopulated: number;
  };
}

export class MetadataCollector {
  constructor(private workbook: WorkBook) {}

  collect(): WorkbookMetadata {
    const metadata: WorkbookMetadata = {
      formulaCount: 0,
      sheetCount: this.workbook.SheetNames.length,
      namedRanges: this.getNamedRanges(),
      metrics: {
        totalCells: 0,
        dataCells: 0,
        emptyCells: 0,
        percentagePopulated: 0
      }
    };

    // Process each sheet
    for (const sheetName of this.workbook.SheetNames) {
      const sheet = this.workbook.Sheets[sheetName];
      metadata.formulaCount += this.countFormulas(sheet);
      this.collectSheetMetrics(sheet, metadata.metrics);
    }

    // Calculate percentage populated
    metadata.metrics.percentagePopulated = 
      (metadata.metrics.dataCells / metadata.metrics.totalCells) * 100;

    return metadata;
  }

  private countFormulas(sheet: WorkSheet): number {
    let count = 0;
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      const cell = sheet[cellAddress];
      if (cell && typeof cell === 'object' && 'f' in cell) {
        count++;
      }
    }
    return count;
  }

  private getNamedRanges(): string[] {
    const namedRanges: string[] = [];
    if (this.workbook.Workbook?.Names) {
      for (const name of this.workbook.Workbook.Names) {
        namedRanges.push(name.Name);
      }
    }
    return namedRanges;
  }

  private collectSheetMetrics(sheet: WorkSheet, metrics: WorkbookMetadata['metrics']) {
    const range = sheet['!ref'];
    if (!range) return;

    // Count cells
    for (const cellAddress in sheet) {
      if (cellAddress[0] === '!') continue;
      metrics.totalCells++;
      
      const cell = sheet[cellAddress];
      if (cell && 'v' in cell && cell.v !== undefined && cell.v !== '') {
        metrics.dataCells++;
      } else {
        metrics.emptyCells++;
      }
    }
  }
}