import { FormulaNode, TokenType, Token, FormulaAnalysis } from "./types";
export class ExcelFormulaParser {
  private readonly OPERATOR_PRECEDENCE: Record<string, number> = {
    ":": 5,
    "^": 4,
    "*": 3,
    "/": 3,
    "+": 2,
    "-": 2,
    "&": 1,
    "=": 0,
    ">": 0,
    "<": 0,
    ">=": 0,
    "<=": 0,
    "<>": 0,
  };

  private readonly VOLATILE_FUNCTIONS = new Set([
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "INDIRECT",
    "OFFSET",
    "CELL",
  ]);

  private readonly FUNCTION_ARITY: Record<string, number | [number, number]> = {
    SUM: [1, Infinity],
    AVERAGE: [1, Infinity],
    COUNT: [1, Infinity],
    MIN: [1, Infinity],
    MAX: [1, Infinity],
    IF: 3,
    AND: [2, Infinity],
    OR: [2, Infinity],
    NOT: 1,
    VLOOKUP: [3, 4],
    HLOOKUP: [3, 4],
    INDEX: [2, 3],
    MATCH: [2, 3],
    SUMIF: [2, 3],
    COUNTIF: 2,
    LEFT: [1, 2],
    RIGHT: [1, 2],
    MID: 3,
    NOW: 0,
    TODAY: 0,
    RAND: 0,
    SUMIFS: [3, Infinity],
  };

  parse(formula: string): FormulaNode {
    formula = formula.trim().replace(/^=/, "");
    if (!formula) throw new Error("Empty formula");

    const tokens = this.tokenize(formula);
    return this.parseTokens(tokens);
  }

  private tokenize(formula: string): Token[] {
    const tokens: Token[] = [];
    let position = 0;
    let argSeparatorExpected = false;

    while (position < formula.length) {
      const char = formula[position];

      // Skip whitespace
      if (/\s/.test(char)) {
        position++;
        continue;
      }

      // Handle argument separator
      if (char === ",") {
        tokens.push({
          type: TokenType.OPERATOR,
          value: ",",
          position: { start: position, end: position + 1 },
        });
        position++;
        argSeparatorExpected = false;
        continue;
      }

      // Handle parentheses
      if (char === "(" || char === ")") {
        tokens.push({
          type: TokenType.OPERATOR,
          value: char,
          position: { start: position, end: position + 1 },
        });
        position++;
        continue;
      }

      // Handle operators
      if (/[+\-*\/^&=<>:]/.test(char)) {
        let value = char;
        if (position + 1 < formula.length) {
          const nextChar = formula[position + 1];
          if (
            (char === "<" && (nextChar === ">" || nextChar === "=")) ||
            (char === ">" && nextChar === "=")
          ) {
            value += nextChar;
            position++;
          }
        }
        tokens.push({
          type: TokenType.OPERATOR,
          value,
          position: { start: position, end: position + value.length },
        });
        position += value.length;
        continue;
      }

      // Handle strings
      if (char === '"' || char === "'") {
        let value = "";
        const start = position;
        position++;
        while (position < formula.length && formula[position] !== char) {
          value += formula[position];
          position++;
        }
        position++; // Skip closing quote
        tokens.push({
          type: TokenType.TEXT,
          value,
          position: { start, end: position },
        });
        continue;
      }

      // Handle numbers
      if (/[0-9]/.test(char)) {
        let value = "";
        const start = position;
        let hasDecimal = false;
        let hasExponent = false;

        while (position < formula.length) {
          const currentChar = formula[position];
          if (/[0-9]/.test(currentChar)) {
            value += currentChar;
          } else if (currentChar === "." && !hasDecimal) {
            value += currentChar;
            hasDecimal = true;
          } else if (
            (currentChar === "E" || currentChar === "e") &&
            !hasExponent
          ) {
            value += currentChar;
            hasExponent = true;
            if (
              position + 1 < formula.length &&
              /[+-]/.test(formula[position + 1])
            ) {
              position++;
              value += formula[position];
            }
          } else {
            break;
          }
          position++;
        }

        tokens.push({
          type: TokenType.NUMBER,
          value,
          position: { start, end: position },
        });
        continue;
      }

      // Handle references and functions
      if (/[A-Za-z$]/.test(char)) {
        let value = "";
        const start = position;
        while (
          position < formula.length &&
          /[A-Za-z0-9$_:]/.test(formula[position])
        ) {
          value += formula[position];
          position++;
        }

        // Check if it's a range reference
        if (value.includes(":")) {
          tokens.push({
            type: TokenType.RANGE_REFERENCE,
            value,
            position: { start, end: position },
          });
        } else if (position < formula.length && formula[position] === "(") {
          // It's a function
          tokens.push({
            type: TokenType.FUNCTION,
            value,
            position: { start, end: position },
          });
        } else {
          // It's a cell reference
          tokens.push({
            type: TokenType.CELL_REFERENCE,
            value,
            position: { start, end: position },
          });
        }
      } else {
        position++;
      }
    }

    return tokens;
  }

 
  private parseTokens(tokens: Token[]): FormulaNode {
    const output: FormulaNode[] = [];
    const operators: Token[] = [];
    let expectingOperand = true;
    let currentFunctionArgs: number[] = [];

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];

      switch (token.type) {
        case TokenType.NUMBER:
        case TokenType.TEXT:
        case TokenType.CELL_REFERENCE:
        case TokenType.RANGE_REFERENCE:
          output.push(this.createNode(token));
          expectingOperand = false;
          break;

        case TokenType.FUNCTION:
          operators.push(token);
          currentFunctionArgs.push(0);
          expectingOperand = true;
          break;

        case TokenType.OPERATOR:
          if (token.value === '(') {
            operators.push(token);
            expectingOperand = true;
          } else if (token.value === ')') {
            while (operators.length > 0 && operators[operators.length - 1].value !== '(') {
              this.processOperator(operators.pop()!, output);
            }
            if (operators.length > 0) {
              operators.pop(); // Remove '('
              if (operators.length > 0 && operators[operators.length - 1].type === TokenType.FUNCTION) {
                const func = operators.pop()!;
                const argCount = currentFunctionArgs.pop()! + 1;
                const args = output.splice(-argCount);
                output.push(this.createFunctionNode(func, args));
              }
            }
            expectingOperand = false;
          } else if (token.value === ',') {
            while (operators.length > 0 && operators[operators.length - 1].value !== '(') {
              this.processOperator(operators.pop()!, output);
            }
            if (currentFunctionArgs.length > 0) {
              currentFunctionArgs[currentFunctionArgs.length - 1]++;
            }
            expectingOperand = true;
          } else {
            while (operators.length > 0 &&
                   operators[operators.length - 1].type === TokenType.OPERATOR &&
                   this.OPERATOR_PRECEDENCE[operators[operators.length - 1].value] >= 
                   this.OPERATOR_PRECEDENCE[token.value]) {
              this.processOperator(operators.pop()!, output);
            }
            operators.push(token);
            expectingOperand = true;
          }
          break;
      }
    }

    while (operators.length > 0) {
      const operator = operators.pop()!;
      if (operator.value === '(' || operator.value === ')') {
        throw new Error('Unmatched parentheses');
      }
      this.processOperator(operator, output);
    }

    if (output.length !== 1) {
      throw new Error('Invalid formula structure');
    }

    return output[0];
  }

  private processOperator(operator: Token, output: FormulaNode[]) {
    if (operator.value === ',') return;
    
    if (output.length < 2) {
      throw new Error('Invalid formula structure: not enough operands');
    }
    const right = output.pop()!;
    const left = output.pop()!;
    output.push({
      type: 'OPERATOR',
      value: operator.value,
      children: [left, right],
      token: operator
    });
  }

  private createNode(token: Token): FormulaNode {
    return {
      type: token.type,
      value: token.value,
      children: [],
      token,
    };
  }

  private createFunctionNode(func: Token, args: FormulaNode[]): FormulaNode {
    const minArity = this.getFunctionArity(func.value);
    if (args.length < minArity) {
      throw new Error(`Function ${func.value} requires at least ${minArity} arguments, got ${args.length}`);
    }
    return {
      type: 'FUNCTION',
      value: func.value,
      children: args,
      token: func
    };
  }
  private getFunctionArity(functionName: string): number {
    const arity = this.FUNCTION_ARITY[functionName.toUpperCase()];
    if (!arity) return 1;
    return Array.isArray(arity) ? arity[0] : arity;
  }

  private getMaxArity(functionName: string): number {
    const arity = this.FUNCTION_ARITY[functionName.toUpperCase()];
    if (!arity) return 1;
    return Array.isArray(arity) ? arity[1] : arity;
  }

  analyze(formula: string): FormulaAnalysis {
    try {
      const node = this.parse(formula);
      const volatileFuncs = this.getVolatileFunctions(node);
      const deps = Array.from(this.getDependencies(node));
      const ranges = this.getUsedRanges(node);
      
      const analysis: FormulaAnalysis = {
        isValid: true,
        errors: [],
        complexity: this.calculateComplexity(node),
        volatileFunctions: volatileFuncs,
        dependencies: deps,
        usedRanges: ranges
      };

      // Check for circular references
      const refCount = new Map<string, number>();
      deps.forEach(ref => {
        refCount.set(ref, (refCount.get(ref) || 0) + 1);
      });
      
      refCount.forEach((count, ref) => {
        if (count > 1) {
          analysis.errors.push(`Circular reference detected: ${ref}`);
        }
      });

      // Add warnings for volatile functions
      if (volatileFuncs.length > 0) {
        analysis.errors.push(`Volatile functions used: ${volatileFuncs.join(', ')}`);
      }

      // Check for unmatched parentheses
      const openCount = (formula.match(/\(/g) || []).length;
      const closeCount = (formula.match(/\)/g) || []).length;
      if (openCount !== closeCount) {
        analysis.errors.push('Unmatched parentheses');
        analysis.isValid = false;
      }

      return analysis;
    } catch (error) {
      return {
        isValid: false,
        errors: [(error as Error).message],
        complexity: 0,
        volatileFunctions: [],
        dependencies: [],
        usedRanges: []
      };
    }
  }

  private calculateComplexity(node: FormulaNode): number {
    let complexity = 0;
    const traverse = (n: FormulaNode, depth: number = 0) => {
      if (n.type === "FUNCTION") {
        complexity += 1 + depth * 0.5;
      } else if (n.type === "OPERATOR") {
        complexity += 0.5;
      }
      n.children.forEach((child) => traverse(child, depth + 1));
    };
    traverse(node);
    return complexity;
  }

  private getVolatileFunctions(node: FormulaNode): string[] {
    const volatileFuncs: string[] = [];
    const traverse = (n: FormulaNode) => {
      if (n.type === 'FUNCTION' && this.VOLATILE_FUNCTIONS.has(n.value.toUpperCase())) {
        volatileFuncs.push(n.value.toUpperCase());
      }
      n.children.forEach(traverse);
    };
    traverse(node);
    return volatileFuncs;
  }

  private getDependencies(node: FormulaNode): Set<string> {
    const deps = new Set<string>();
    const traverse = (n: FormulaNode) => {
      if (n.type === TokenType.CELL_REFERENCE) {
        deps.add(n.value);
      }
      n.children.forEach(traverse);
    };
    traverse(node);
    return deps;
  }

  private getUsedRanges(node: FormulaNode): string[] {
    const ranges: string[] = [];
    const traverse = (n: FormulaNode) => {
      if (n.type === TokenType.RANGE_REFERENCE) {
        ranges.push(n.value);
      }
      n.children.forEach(traverse);
    };
    traverse(node);
    return ranges;
  }
}
