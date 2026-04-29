/**
 * Mock AI service for demonstrating functionality without backend integration
 * Simulates realistic AI responses and operation workflows
 */

import type { AIOperation } from '../types/index';
import { generateId, generateSampleData } from '../utils/spreadsheetUtils';

export class MockAIService {
  /**
   * Process a user prompt and return simulated AI operations
   */
  static async processPrompt(prompt: string, workbookId: string, worksheetId?: string): Promise<{
    promptId: string;
    operations: AIOperation[];
  }> {
    const promptId = generateId();
    const operations = this.generateOperationsFromPrompt(prompt, worksheetId);
    
    // Simulate processing delay
    await new Promise(resolve => setTimeout(resolve, 1000 + Math.random() * 2000));
    
    return { promptId, operations };
  }

  /**
   * Generate realistic operations based on the prompt content
   */
  private static generateOperationsFromPrompt(prompt: string, worksheetId?: string): AIOperation[] {
    const lowerPrompt = prompt.toLowerCase();
    const operations: AIOperation[] = [];

    // Create worksheets
    if (lowerPrompt.includes('worksheet') || lowerPrompt.includes('sheet')) {
      const match = lowerPrompt.match(/(\d+)\s+(?:worksheet|sheet)/);
      const count = match ? parseInt(match[1]) : 1;
      
      const departmentNames = ['Sales', 'Marketing', 'Finance', 'HR', 'Operations'];
      for (let i = 0; i < Math.min(count, 5); i++) {
        operations.push({
          id: generateId(),
          type: 'create_worksheet',
          description: `Creating ${departmentNames[i] || `Sheet ${i + 1}`} worksheet`,
          status: 'pending',
          timestamp: new Date(),
          worksheetId,
        });
      }
    }

    // Populate data
    if (lowerPrompt.includes('data') || lowerPrompt.includes('table') || lowerPrompt.includes('fill')) {
      operations.push({
        id: generateId(),
        type: 'populate_data',
        description: 'Generating sample data for the spreadsheet',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
        cellRange: { startRow: 0, startCol: 0, endRow: 10, endCol: 4 },
      });
    }

    // Import JSON
    if (lowerPrompt.includes('json')) {
      operations.push({
        id: generateId(),
        type: 'import_json',
        description: 'Converting JSON data to spreadsheet format',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
      });
    }

    // Format cells
    if (lowerPrompt.includes('format') || lowerPrompt.includes('bold') || lowerPrompt.includes('header')) {
      operations.push({
        id: generateId(),
        type: 'format_cells',
        description: 'Applying formatting to headers and data',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
        cellRange: { startRow: 0, startCol: 0, endRow: 0, endCol: 10 },
      });
    }

    // Create formulas
    if (lowerPrompt.includes('formula') || lowerPrompt.includes('calculate') || lowerPrompt.includes('sum')) {
      operations.push({
        id: generateId(),
        type: 'create_formula',
        description: 'Creating formulas for calculations',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
      });
    }

    // Sort data
    if (lowerPrompt.includes('sort')) {
      operations.push({
        id: generateId(),
        type: 'sort_data',
        description: 'Sorting data by specified criteria',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
      });
    }

    // Generate summary
    if (lowerPrompt.includes('summary') || lowerPrompt.includes('total')) {
      operations.push({
        id: generateId(),
        type: 'calculate_summary',
        description: 'Creating summary statistics and totals',
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
      });
    }

    // Default operation if nothing specific is found
    if (operations.length === 0) {
      operations.push({
        id: generateId(),
        type: 'populate_data',
        description: `Processing request: ${prompt}`,
        status: 'pending',
        timestamp: new Date(),
        worksheetId,
      });
    }

    return operations;
  }

  /**
   * Simulate operation execution with realistic timing and results
   */
  static async executeOperation(operation: AIOperation): Promise<any> {
    // Simulate processing time based on operation type
    const processingTimes = {
      create_worksheet: 2000,
      populate_data: 3000,
      format_cells: 1500,
      create_formula: 2500,
      import_json: 4000,
      sort_data: 2000,
      filter_data: 1800,
      calculate_summary: 2200,
      generate_chart: 3500,
    };

    const delay = processingTimes[operation.type] || 2000;
    await new Promise(resolve => setTimeout(resolve, delay));

    // Generate realistic results based on operation type
    switch (operation.type) {
      case 'create_worksheet':
        return {
          worksheetId: generateId(),
          name: operation.description.includes('Sales') ? 'Sales' : 'New Sheet',
          message: 'Worksheet created successfully',
        };

      case 'populate_data':
        return {
          cellsModified: 50,
          data: generateSampleData(10, 4),
          message: 'Sample data populated successfully',
        };

      case 'format_cells':
        return {
          formattedCells: ['A1', 'B1', 'C1', 'D1'],
          formatting: { bold: true, backgroundColor: '#f3f4f6' },
          message: 'Headers formatted with bold styling',
        };

      case 'create_formula':
        return {
          formulas: [
            { cell: 'E1', formula: '=SUM(B:B)', description: 'Total sales' },
            { cell: 'E2', formula: '=AVERAGE(B:B)', description: 'Average sales' },
          ],
          message: 'Formulas created for calculations',
        };

      case 'import_json':
        return {
          rowsImported: 25,
          columnsCreated: 4,
          message: 'JSON data imported successfully',
        };

      case 'sort_data':
        return {
          sortedRows: 45,
          criteria: 'Sales (descending)',
          message: 'Data sorted by sales amount',
        };

      case 'calculate_summary':
        return {
          summaryData: {
            totalSales: 125000,
            averageSales: 2500,
            topProduct: 'Laptop',
            topRegion: 'North',
          },
          message: 'Summary statistics calculated',
        };

      default:
        return {
          message: 'Operation completed successfully',
        };
    }
  }

  /**
   * Generate example prompts for user guidance
   */
  static getExamplePrompts(): string[] {
    return [
      "Create 5 worksheets for different departments (Sales, Marketing, Finance, HR, Operations)",
      "Fill the current sheet with sample sales data including products, amounts, and regions",
      "Import this JSON data as a table: [{\"Product\": \"Laptop\", \"Sales\": 1500, \"Region\": \"North\"}]",
      "Format the first row as headers with bold text and gray background",
      "Create formulas to calculate total sales and average sales in column E",
      "Sort the data by sales amount in descending order",
      "Generate a summary table with totals for each region",
      "Create a pivot table showing sales by product category",
      "Add conditional formatting to highlight sales above $1000",
      "Generate a chart showing monthly sales trends"
    ];
  }

  /**
   * Generate realistic error scenarios for testing
   */
  static simulateError(operation: AIOperation): Error {
    const errorScenarios = [
      new Error('Invalid data format in JSON import'),
      new Error('Cell range exceeds worksheet boundaries'),
      new Error('Formula syntax error detected'),
      new Error('Insufficient data for chart generation'),
      new Error('Network timeout during data processing'),
    ];

    const randomError = errorScenarios[Math.floor(Math.random() * errorScenarios.length)];
    return randomError;
  }

  /**
   * Get operation status with realistic progress updates
   */
  static getOperationProgress(operation: AIOperation, elapsedTime: number): {
    status: AIOperation['status'];
    progress: number;
    message?: string;
  } {
    const totalTime = 3000; // 3 seconds average
    const progress = Math.min(Math.floor((elapsedTime / totalTime) * 100), 95);

    if (elapsedTime < totalTime * 0.3) {
      return {
        status: 'pending',
        progress: 0,
        message: 'Queued for processing...',
      };
    } else if (elapsedTime < totalTime * 0.9) {
      return {
        status: 'in-progress',
        progress,
        message: this.getProgressMessage(operation.type, progress),
      };
    } else if (Math.random() > 0.1) { // 90% success rate
      return {
        status: 'completed',
        progress: 100,
        message: 'Operation completed successfully',
      };
    } else {
      return {
        status: 'failed',
        progress: 0,
        message: 'Operation failed due to validation error',
      };
    }
  }

  /**
   * Get context-appropriate progress messages
   */
  private static getProgressMessage(operationType: AIOperation['type'], progress: number): string {
    const messages = {
      create_worksheet: [
        'Analyzing worksheet requirements...',
        'Creating worksheet structure...',
        'Setting up columns and formatting...',
        'Finalizing worksheet setup...',
      ],
      populate_data: [
        'Generating sample data...',
        'Formatting cell values...',
        'Applying data validation...',
        'Completing data population...',
      ],
      format_cells: [
        'Analyzing formatting requirements...',
        'Applying text styles...',
        'Setting cell borders and colors...',
        'Finalizing formatting...',
      ],
      create_formula: [
        'Analyzing formula requirements...',
        'Creating formula expressions...',
        'Validating formula syntax...',
        'Applying formulas to cells...',
      ],
      import_json: [
        'Parsing JSON structure...',
        'Mapping data to cells...',
        'Validating data types...',
        'Completing import process...',
      ],
      sort_data: [
        'Analyzing sort criteria...',
        'Sorting data rows...',
        'Reordering columns if needed...',
        'Completing sort operation...',
      ],
      filter_data: [
        'Applying filter criteria...',
        'Filtering data rows...',
        'Updating visible data...',
        'Completing filter operation...',
      ],
      calculate_summary: [
        'Calculating summary statistics...',
        'Aggregating data values...',
        'Generating summary report...',
        'Completing summary calculation...',
      ],
      generate_chart: [
        'Analyzing chart data...',
        'Creating chart structure...',
        'Applying chart styling...',
        'Generating final chart...',
      ],
    };

    const operationMessages = messages[operationType] || [
      'Processing request...',
      'Analyzing data...',
      'Applying changes...',
      'Completing operation...',
    ];

    const messageIndex = Math.floor((progress / 100) * operationMessages.length);
    return operationMessages[Math.min(messageIndex, operationMessages.length - 1)];
  }
}