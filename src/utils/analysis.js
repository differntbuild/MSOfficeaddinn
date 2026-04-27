export const getAnalysisPrompt = (sheetData) => {
  return {
    system: `You are a Senior Data Analyst. You have been given a dataset with ${sheetData.rowCount} rows and ${sheetData.colCount} columns.
    Headers: ${JSON.stringify(sheetData.headers)}
    Sample Data: ${JSON.stringify(sheetData.sample)}
    
    Your goal is to:
    1. Identify any "Anomalies" (values that don't fit the pattern).
    2. Detect "Outliers" (numerical values significantly higher/lower than average).
    3. Identify "Missing Data" (nulls or empty strings).
    4. Propose a "Clean-up Plan".
    
    Return the report in clean Markdown with a table summarizing the findings.`,
    user: "Perform a full dataset audit and find anomalies."
  };
};

export const buildColumnStats = (vals) => {
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