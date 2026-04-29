import { useState, useCallback, useRef } from 'react';

/* global Excel, Office */

export const useOffice = () => {
  const toColLetter = (index) => {
    let letter = '';
    let i = index;
    while (i >= 0) {
      letter = String.fromCharCode((i % 26) + 65) + letter;
      i = Math.floor(i / 26) - 1;
    }
    return letter;
  };

  const getSelectionContext = useCallback(async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      const headerRow = sheet.getRange("1:1");
      headerRow.load('values');
      await context.sync();

      const fullAddress = range.address || '';
      const localAddress = fullAddress.includes('!') ? fullAddress.split('!')[1] : fullAddress;
      const startRef = localAddress.split(':')[0] || 'A1';
      const colMatch = startRef.match(/[A-Z]+/i);
      const rowMatch = startRef.match(/\d+/);
      const startColLetters = colMatch ? colMatch[0].toUpperCase() : 'A';
      const startRow = rowMatch ? parseInt(rowMatch[0], 10) : 1;

      const colToIndex = (letters) =>
        letters.split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
      const startCol = colToIndex(startColLetters);
      const headersByColumn = {};

      const cells = [];
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const colLetter = toColLetter(startCol + c);
          const cellAddress = `${colLetter}${startRow + r}`;
          if (!headersByColumn[colLetter]) {
            headersByColumn[colLetter] = headerRow.values?.[0]?.[startCol + c] || colLetter;
          }
          cells.push({
            address: cellAddress,
            value: range.values?.[r]?.[c],
            formula: range.formulas?.[r]?.[c],
          });
        }
      }

      return {
        sheetName: sheet.name,
        address: localAddress,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        headersByColumn,
        selectedHeaders: Object.values(headersByColumn),
        cells,
      };
    });
  }, []);

  const getFormulaAtCell = useCallback(async (cellAddress) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      cell.load(['address', 'values', 'formulas']);
      await context.sync();
      return {
        address: cell.address.includes('!') ? cell.address.split('!')[1] : cell.address,
        value: cell.values?.[0]?.[0],
        formula: cell.formulas?.[0]?.[0],
      };
    });
  }, []);

  const insertFormulaBelow = useCallback(async (formula, options = {}) => {
    return await Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load(['rowIndex', 'columnIndex', 'rowCount']);
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      await context.sync();

      const targetRow = selected.rowIndex + selected.rowCount;
      const targetCol = selected.columnIndex;
      const target = sheet.getCell(targetRow, targetCol);
      target.load(['address', 'values']);
      await context.sync();

      const existingValue = target.values?.[0]?.[0];
      const hasValue = existingValue !== null && existingValue !== undefined && String(existingValue).trim() !== '';
      const targetAddress = target.address.includes('!') ? target.address.split('!')[1] : target.address;

      if (hasValue && !options.forceOverwrite) {
        return {
          blocked: true,
          targetAddress,
          existingValue,
        };
      }

      target.formulas = [[formula]];
      await context.sync();

      return {
        success: true,
        address: targetAddress,
      };
    });
  }, []);

  const writeToSessionLog = async (action, details) => {
    try {
      await Excel.run(async (context) => {
        const customXmlParts = context.workbook.customXmlParts;
        const part = customXmlParts.add(`<Log><Action>${action}</Action><Details>${JSON.stringify(details)}</Details><Time>${new Date().toISOString()}</Time></Log>`);
        await context.sync();
      });
    } catch (err) {
      console.warn('Logging failed:', err);
    }
  };

  const saveToWorkbookMemory = async (key, data) => {
    return await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      settings.add(key, JSON.stringify(data));
      await context.sync();
      return true;
    });
  };

  const loadFromWorkbookMemory = async (key) => {
    return await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      const setting = settings.getItemOrNullObject(key);
      setting.load('value');
      await context.sync();
      if (setting.isNullObject) return null;
      try {
        return JSON.parse(setting.value);
      } catch {
        return setting.value;
      }
    });
  };

  const getFullSheetContext = useCallback(async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const range = sheet.getUsedRange();
      range.load(['values', 'formulas', 'address', 'rowCount', 'columnCount', 'numberFormat']);
      
      const namedItems = context.workbook.names;
      namedItems.load('items/name,items/value');
      
      const sheets = context.workbook.worksheets;
      sheets.load('items/name,items/visibility');

      // Check if this range is already part of a Table
      const tables = sheet.tables;
      tables.load("items/address");
      
      await context.sync();

      const headers = range.values[0] || [];
      const columnTypes = [];
      if (range.values.length > 1) {
        headers.forEach((h, i) => {
          let type = 'string';
          // Scan up to 5 rows to figure out the true data type, in case row 1 is empty
          for (let r = 1; r < Math.min(6, range.values.length); r++) {
            const val = range.values[r][i];
            if (val !== "" && val !== null) {
               if (typeof val === 'number') type = 'number';
               else if (val instanceof Date) type = 'date';
               else if (typeof val === 'boolean') type = 'boolean';
               break; // Found the type, stop scanning
            }
          }
          columnTypes.push({ header: h, type });
        });
      }

      const buildColumnStats = (vals) => {
        const stats = {};
        if (vals.length < 2) return stats;
        const head = vals[0];
        
        head.forEach((h, i) => {
          const colVals = vals.slice(1).map(r => r[i]).filter(v => typeof v === 'number');
          if (colVals.length > 0) {
            const sum = colVals.reduce((a, b) => a + b, 0);
            const mean = sum / colVals.length;
            const min = Math.min(...colVals);
            const max = Math.max(...colVals);
            
            // Add Standard Deviation calculation
            const variance = colVals.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / colVals.length;
            const stdDev = Math.sqrt(variance);

            stats[h] = { mean, min, max, stdDev, count: colVals.length };
          }
        });
        return stats;
      };

      const stats = buildColumnStats(range.values);

      // Check if this range is already part of a Table
      let isTable = false;
      const rangeAddress = range.address ? (range.address.split('!')[1] || range.address) : '';
      for (const table of tables.items) {
        const tableAddress = table.address ? (table.address.split('!')[1] || table.address) : '';
        if (rangeAddress && tableAddress && rangeAddress.includes(tableAddress)) {
          isTable = true;
          break;
        }
      }

      return {
        sheetName: sheet.name,
        address: range.address,
        dimensions: `${range.rowCount} rows × ${range.columnCount} columns`,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        headers,
        columnTypes,
        stats,
        isTable,
        namedRanges: namedItems.items.map(n => ({ name: n.name, value: n.value })),
        sheetNames: sheets.items.map(s => ({ name: s.name, visible: s.visibility === 'Visible' })),
      };
    });
  }, []);

  const createPivotAndChart = useCallback(async (config) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const source = sheet.getUsedRange();
      
      const newSheet = context.workbook.worksheets.add("AI Report " + Math.floor(Math.random() * 1000));
      const pivotTable = context.workbook.pivotTables.add(config.name || "AIPivot", source, newSheet.getRange("B2"));
      
      config.rows?.forEach(r => pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(r)));
      config.columns?.forEach(c => pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(c)));
      config.values?.forEach(v => {
        const field = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(v.header));
        field.summarizeBy = v.func || "Sum";
      });

      const chart = newSheet.charts.add(config.chartType || "ColumnClustered", newSheet.getRange("B2").getSurroundingRegion(), "Auto");
      chart.title.text = config.title || "AI Generated Insights";
      chart.setPosition("H2", "P20");

      newSheet.activate();
      await context.sync();
      return { success: true, sheetName: newSheet.name };
    });
  }, []);

  const createAutonomousDashboard = useCallback(async (configStr) => {
    return await Excel.run(async (context) => {
      let config;
      try {
        config = typeof configStr === 'string' ? JSON.parse(configStr) : configStr;
      } catch (e) {
        config = configStr;
      }

      const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sourceSheet.getUsedRange();
      
      // 1. Create Dashboard Sheet
      let dbName = config.title || "Executive Summary";
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      
      let finalName = dbName;
      let count = 1;
      while (sheets.items.find(s => s.name === finalName)) {
        finalName = `${dbName} ${++count}`;
      }
      
      const dbSheet = context.workbook.worksheets.add(finalName);
      dbSheet.showGridlines = false;

      // 2. Add Title Header
      const headerRange = dbSheet.getRange("B2:L3");
      headerRange.merge();
      headerRange.values = [[config.title || "Intelligent Executive Summary"]];
      headerRange.format.font.size = 20;
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#1a1a1a";
      headerRange.format.verticalAlignment = "Center";
      headerRange.format.fill.color = "#f3f2f1";

      let currentRow = 5;

      // 3. Process KPI Metrics (Top Row)
      if (config.metrics && config.metrics.length > 0) {
        const kpiStartRow = currentRow;
        config.metrics.slice(0, 4).forEach((m, i) => {
          const colLetter = String.fromCharCode(66 + (i * 3)); // B, E, H, K
          const labelCell = dbSheet.getRange(`${colLetter}${kpiStartRow}`);
          const valCell = dbSheet.getRange(`${colLetter}${kpiStartRow + 1}`);
          
          labelCell.values = [[m.label.toUpperCase()]];
          labelCell.format.font.size = 9;
          labelCell.format.font.color = "#605e5c";
          labelCell.format.font.bold = true;

          // Placeholder logic: We'll use a hidden pivot to get the actual value or a formula
          // For now, let's just create a very small 1x1 pivot to extract the metric
          try {
            const tempPivot = context.workbook.pivotTables.add("KPI_" + i, sourceRange, dbSheet.getRange("Z100")); // hidden
            const field = tempPivot.dataHierarchies.add(tempPivot.hierarchies.getItem(m.header));
            field.summarizeBy = m.func || "Sum";
            
            // Link the cell to the pivot result
            valCell.formulas = [[`='${finalName}'!Z102`]]; 
            valCell.format.font.size = 16;
            valCell.format.font.bold = true;
            valCell.format.font.color = "#217346";
            valCell.format.numberFormat = "#,##0";
          } catch(e) {
            valCell.values = [["--"]];
          }
        });
        currentRow += 4;
      }

      // 4. Process Pivots & Charts
      for (const p of config.pivots || []) {
        try {
          const ptRange = dbSheet.getRange(`B${currentRow}`);
          const pivotTable = context.workbook.pivotTables.add(p.name.replace(/\s+/g, ''), sourceRange, ptRange);
          
          p.rows?.forEach(r => pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(r)));
          p.values?.forEach(v => {
            const field = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(v.header));
            field.summarizeBy = v.func || "Sum";
          });

          const chart = dbSheet.charts.add(p.chartType || "ColumnClustered", ptRange.getSurroundingRegion(), "Auto");
          chart.title.text = p.name;
          chart.setPosition(`H${currentRow}`, `L${currentRow + 12}`);
          
          currentRow += 15;
        } catch (err) {
          console.warn("Failed to create pivot:", p.name, err);
        }
      }

      dbSheet.activate();
      await context.sync();
      await writeToSessionLog('Dashboard Created', { name: finalName });
      return { success: true, sheetName: finalName };
    });
  }, []);

  const convertRangeToTable = useCallback(async (address) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = address ? sheet.getRange(address) : sheet.getUsedRange();
      
      const tables = sheet.tables;
      tables.load("items/address");
      range.load("address");
      await context.sync();

      const rangeAddr = range.address.split('!')[1] || range.address;
      for (const t of tables.items) {
        const tAddr = t.address.split('!')[1] || t.address;
        if (rangeAddr === tAddr || rangeAddr.includes(tAddr) || tAddr.includes(rangeAddr)) {
          // Already a table or overlaps
          return { success: true, tableName: t.name, alreadyTable: true };
        }
      }

      try {
        const table = sheet.tables.add(range, true);
        table.name = "Table_" + Math.floor(Math.random() * 10000);
        await context.sync();
        return { success: true, tableName: table.name };
      } catch (e) {
        return { success: false, error: e.message };
      }
    });
  }, []);

  const revertTableToRange = useCallback(async (tableName) => {
    return await Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);
      table.convertToRange();
      await context.sync();
      return { success: true };
    });
  }, []);

  const applyTrimToRange = useCallback(async (rangeAddress) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.load(['values', 'address']);
      await context.sync();

      const original = JSON.parse(JSON.stringify(range.values));
      const trimmed = range.values.map(row => row.map(cell => 
        (typeof cell === 'string') ? cell.trim() : cell
      ));
      
      range.values = trimmed;
      await context.sync();
      return { success: true, address: range.address, original };
    });
  }, []);

  const revertCellValues = useCallback(async (address, originalValues) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.values = originalValues;
      await context.sync();
      return { success: true };
    });
  }, []);

  const navigateToCell = useCallback(async (cellAddress) => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange(cellAddress).select();
      await context.sync();
    });
  }, []);

  const highlightCells = useCallback(async (cellAddress, color = '#FFE699') => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange(cellAddress).format.fill.color = color;
      await context.sync();
    });
  }, []);

  const writeCellValue = useCallback(async (cellAddress, value, options = {}) => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      
      if (value.toString().startsWith('=')) {
        cell.formulas = [[value]];
      } else {
        cell.values = [[value]];
      }
      
      await context.sync();
      return { success: true };
    });
  }, []);

  return {
    getSelectionContext,
    getFormulaAtCell,
    insertFormulaBelow,
    getFullSheetContext,
    createPivotAndChart,
    createAutonomousDashboard,
    convertRangeToTable,
    revertTableToRange,
    applyTrimToRange,
    revertCellValues,
    navigateToCell,
    highlightCells,
    writeCellValue,
    saveToWorkbookMemory,
    loadFromWorkbookMemory,
  };
};
