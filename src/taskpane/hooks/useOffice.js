import { useState, useCallback, useRef } from 'react';

/* global Excel, Office */

export const useOffice = () => {
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
        config = JSON.parse(configStr);
      } catch (e) {
        // AI might send raw object
        config = configStr;
      }

      const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sourceSheet.getUsedRange();
      
      // 1. Create Dashboard Sheet
      let dbName = "Executive Summary";
      let count = 1;
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      
      while (sheets.items.find(s => s.name === dbName)) {
        dbName = `Executive Summary ${++count}`;
      }
      
      const dbSheet = context.workbook.worksheets.add(dbName);
      dbSheet.showGridlines = false;

      // 2. Add Title Header
      const headerRange = dbSheet.getRange("B2:L3");
      headerRange.merge();
      headerRange.values = [[config.title || "Intelligent Executive Summary"]];
      headerRange.format.font.size = 18;
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#1a1a1a";
      headerRange.format.verticalAlignment = "Center";

      // 3. Process Pivots
      let startRow = 6;
      for (const p of config.pivots || []) {
        try {
          const ptRange = dbSheet.getRange(`B${startRow}`);
          const pivotTable = context.workbook.pivotTables.add(p.name, sourceRange, ptRange);
          
          p.rows?.forEach(r => pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(r)));
          p.values?.forEach(v => {
            const field = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(v.header));
            field.summarizeBy = v.func || "Sum";
          });

          const chart = dbSheet.charts.add(p.chartType || "ColumnClustered", ptRange.getSurroundingRegion(), "Auto");
          chart.title.text = p.name;
          chart.setPosition(`H${startRow}`, `L${startRow + 15}`);
          
          startRow += 18;
        } catch (err) {
          console.warn("Failed to create pivot:", p.name, err);
        }
      }

      dbSheet.activate();
      await context.sync();
      await writeToSessionLog('Dashboard Created', { name: dbName });
      return { success: true, sheetName: dbName };
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