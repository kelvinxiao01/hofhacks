/**
 * Service for creating and managing pivot tables in Excel
 */
export class PivotTableService {
  private static instance: PivotTableService;
  
  private constructor() {}
  
  public static getInstance(): PivotTableService {
    if (!PivotTableService.instance) {
      PivotTableService.instance = new PivotTableService();
    }
    return PivotTableService.instance;
  }
  
  /**
   * Create a pivot table in the current worksheet
   * @param options The options for creating the pivot table
   */
  public async createPivotTable(options: {
    name: string;
    sourceRange: string;
    destinationRange: string;
    rows?: string[];
    columns?: string[];
    values?: Array<{ field: string; function?: string }>;
  }): Promise<void> {
    try {
      console.log('Creating pivot table with options:', options);
      
      await Excel.run(async (context) => {
        // Get the active worksheet
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Create the pivot table
        const pivotTable = worksheet.pivotTables.add(
          options.name,
          options.sourceRange,
          options.destinationRange
        );

        // Load the hierarchies to ensure we can access them
        pivotTable.hierarchies.load("items");
        await context.sync();

        // Add row hierarchies
        if (options.rows && options.rows.length > 0) {
          for (const row of options.rows) {
            const hierarchy = pivotTable.hierarchies.getItem(row);
            pivotTable.rowHierarchies.add(hierarchy);
          }
        }

        // Add column hierarchies
        if (options.columns && options.columns.length > 0) {
          for (const column of options.columns) {
            const hierarchy = pivotTable.hierarchies.getItem(column);
            pivotTable.columnHierarchies.add(hierarchy);
          }
        }

        // Add value hierarchies
        if (options.values && options.values.length > 0) {
          for (const value of options.values) {
            const hierarchy = pivotTable.hierarchies.getItem(value.field);
            const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);
            
            // Set the summarization function if specified
            if (value.function) {
              switch (value.function.toLowerCase()) {
                case 'sum':
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
                  break;
                case 'count':
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
                  break;
                case 'average':
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.average;
                  break;
                case 'max':
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.max;
                  break;
                case 'min':
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.min;
                  break;
                default:
                  dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
              }
            }
          }
        }

        await context.sync();
        console.log('Pivot table created successfully');
      });
    } catch (error) {
      console.error('Error creating pivot table:', error);
      throw error;
    }
  }
} 