# [SheetSense](https://ezy.ovh/hyyit)
[![private registry](https://img.shields.io/badge/verdaccio-blue?style=flat-square)](https://vdo.asadk.dev/-/web/detail/sheetsense)
[![npm version](https://img.shields.io/npm/v/sheetsense.svg?style=flat-square)](https://www.npmjs.com/package/sheetsense)
[![npm downloads](https://img.shields.io/npm/dm/sheetsense.svg?style=flat-square)](https://www.npmjs.com/package/sheetsense)
[![license](https://img.shields.io/npm/l/sheetsense.svg?style=flat-square)](https://github.com/asadbek064/sheetsense/blob/main/LICENSE)


A robust TypeScript library for Excel workbook analysis that helps detect quality issues, validate formulas, and ensure data consistency across spreadsheets. SheetSense provides comprehensive validation and analysis capabilities to help maintain high-quality Excel workbooks.

## Features

- Formula Analysis
  - Complexity scoring and validation
  - Division by zero detection
  - Circular reference detection
  - Pattern recognition system
  - Risk assessment system
- Data Validation
  - Type consistency checking
  - Date format validation
  - Column type analysis
  - Mixed type detection
- Hidden Content Detection
  - Hidden cell identification
  - Hidden row/column analysis
  - Hidden range reporting
- Style Checks
  - Number format consistency
  - Format pattern detection
- Metadata Analysis
  - Statistics tracking
  - Formula counting
  - Named ranges listing
  - Hidden content metrics

### Installation
```bash
pnpm install sheetsense
```

### Usage
```ts
import { ExcelAnalyzer } from 'sheetsense';
import { readFile } from 'xlsx';

// Load your workbook
const workbook = XLSX.read(data,{
  type: 'buffer',
  cellFormula: true,
  cellNF: true,
  cellText: true,
  cellStyles: true,
  cellDates: true,
  raw: true
});

// Create analyzer instance
const analyzer = new ExcelAnalyzer(workbook);

// Perform analysis
const analysis = analyzer.analyze();

// Access results
console.log(analysis.issues);         // Array of detected issues
console.log(analysis.metadata);       // Workbook metadata and statistics
```

### Analysis Results
The analyzer returns and `Analysis` object containing:
```ts
{
  issues: Array<{
    type: 'formula' | 'style' | 'data';
    severity: 'error' | 'warning' | 'info';
    cell: string;
    sheet: string;
    message: string;
    suggestion?: string;
  }>;
  metadata: {
    formulaCount: number;
    sheetCount: number;
    namedRanges: string[];
    volatileFunctions: number;
    externalReferences: number;
    hiddenContent: {
      totalHiddenCells: number;
      totalHiddenRows: number;
      totalHiddenColumns: number;
      hiddenRanges: string[];
    };
  };
}
```

### Development
```bash
pnpm install
pnpm test
pnpm run coverage
pnpm run build
```

### Contributing
Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Support
For support, bug reports and feature requests please use GitHub Issues.

## Roadmap
## 🚧 In Development
### Formula Analysis

- ✅ Integrate xlsx for parsing
- ✅ Pattern recognition system
- ✅ Risk assessment system
-    Formula explanation generator
-    Complex formula detection 
-    Basic complexity scoring

### Core Features

- Rule engine interface
- Dependency mapper
- Documentation generator
- Performance analyzer

### Future Features

- Team collaboration
- Audit logging (winston)
- User management
- Rule import/export
- Analysis dashboard
- Claude API integration

