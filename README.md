# SheetSense
A robust TypeScript library for Excel workbook analysis that helps detect quality issues, validate formulas, and ensure data consistency across spreadsheets. SheetSense provides comprehensive validation and analysis capabilities to help maintain high-quality Excel workbooks.

### Installation
```bash
pnpm install sheetsense
```

### Usage
```ts
import { ExcelAnalyzer } from 'sheetsense';
import { readFile } from 'xlsx';

// Load your workbook
const workbook = readFile('example.xlsx');

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

## SheetSense TODO
### Implemented ✅
### Formula Analysis

-    Basic complexity scoring
- ✅ Division by zero detection 
-    Complex formula detection 
- ✅ Basic formula parser

### Data Validation

- ✅ Data type consistency 
- ✅ Date format validation
- ✅ Column type analysis 
- ✅ Mixed type detection 

### Style Checks

- ✅ Number format consistency
- ✅ Format pattern detection 

### Metadata

- ✅ Basic statistics tracking
- ✅ Formula count
- ✅ Sheet count
- ✅ Named ranges list
- ✅ Basic metrics



## 🚧 In Development
### Formula Analysis

- Integrate formula.js/formula-parser
- Pattern recognition system
- Enhanced complexity scoring
- Risk assessment system
- Formula explanation generator

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

