import { AuditFinding } from '../schemas/auditSchemas'; // wait, it's in ../schemas/auditSchemas

export interface OfficeOps {
  convertRangeToTable: (address: string) => Promise<{ success: boolean; tableName?: string }>;
  applyTrimToRange: (address: string) => Promise<{ success: boolean; address?: string; original?: any }>;
  writeCellValue: (address: string, value: string, options?: any) => Promise<{ success: boolean }>;
  deleteRows: (address: string) => Promise<{ success: boolean; address?: string; originalValues?: any }>;
  formatCells: (address: string, options: any) => Promise<{ success: boolean }>;
  getFormulaAtCell?: (address: string) => Promise<{ address: string, value: any, formula: any }>;
}

export interface RevertPayload {
  type: 'table' | 'values' | 'formula' | 'row_delete' | 'format';
  tableName?: string;
  address?: string;
  original?: any;
}

export type EnrichedFinding = AuditFinding & { loc?: string; category?: string; };

export class ActionExecutor {
  private ops: OfficeOps;

  constructor(officeOps: OfficeOps) {
    this.ops = officeOps;
  }

  /**
   * Executes a specific Action based on the finding.
   * Returns a RevertPayload that can be used to undo the action.
   */
  async execute(issue: EnrichedFinding): Promise<RevertPayload | null> {
    try {
      const action = issue.action_type;
      let loc = issue.loc;
      
      if (!loc || loc === '—') {
        console.log('No location provided for issue', issue.id);
        return null;
      }

      // Convert specific custom rules from legacy AnalysisPane
      if (issue.title?.includes('Excel Table') || issue.title?.startsWith('Convert Range')) {
        const res = await this.ops.convertRangeToTable(loc);
        if (res?.success) return { type: 'table', tableName: res.tableName };
        return null;
      } 
      
      if (issue.category === 'trailing-space' || issue.title?.includes('Whitespace')) {
        const res = await this.ops.applyTrimToRange(loc);
        if (res?.success) return { type: 'values', address: res.address, original: res.original };
        return null;
      }

      // Use the strict action_type from the JSON Schema
      switch (action) {
        case 'FIX_FORMULA':
          if (issue.suggested_formula) {
            const singleLoc = loc.includes(':') ? loc.split(':')[0] : loc;
            
            // Try to capture original state for revert
            let original = '';
            if (this.ops.getFormulaAtCell) {
              const cellState = await this.ops.getFormulaAtCell(singleLoc);
              original = cellState?.formula || cellState?.value || '';
            }

            await this.ops.writeCellValue(singleLoc, issue.suggested_formula);
            return { type: 'formula', address: singleLoc, original };
          }
          break;

        case 'DELETE_ROWS':
          const res = await this.ops.deleteRows(loc);
          if (res.success) {
            return { type: 'row_delete', address: res.address, original: res.originalValues };
          }
          break;

        case 'FORMAT_CELLS':
          // Default highlight style for fixes
          const formatOptions = { fill: '#ECFDF5', fontColor: '#065F46' }; 
          await this.ops.formatCells(loc, formatOptions);
          return { type: 'format', address: loc };

        case 'INVESTIGATE':
        case 'CREATE_DASHBOARD':
        case 'EXTERNAL_ENRICH':
        default:
          // These are interactive actions or UI-level routing handled externally
          console.log(`Action ${action} is interactive or not implemented for auto-fix.`);
          break;
      }
    } catch (err) {
      console.error('Execution failed for finding:', issue, err);
      throw err;
    }
    return null;
  }
}
