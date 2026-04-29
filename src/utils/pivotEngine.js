export const getVisualizationPrompt = (headers, request) => {
  return `You are an Excel Data Architect. The user wants to build a pivot table, chart, or metrics dashboard.
  
Available Headers: ${JSON.stringify(headers)}
User Request: "${request}"

CRITICAL INSTRUCTIONS:
1. Determine the best structure for the Pivot Table to answer the request.
2. Choose the best chart type (ColumnClustered, BarClustered, Line, Pie, Doughnut, Area, StackedColumn).
3. Identify Key Performance Indicators (KPIs) if relevant (e.g. Total Revenue, Average Price).
4. Return ONLY valid JSON. No markdown backticks, no explanations.

Expected JSON Format:
{
  "title": "Business Reporting Dashboard",
  "pivots": [
    {
      "name": "Primary Metric Breakdown",
      "rows": ["Exact Header Name"],
      "values": [{"header": "Exact Header Name", "func": "Sum"}], 
      "chartType": "ColumnClustered" 
    }
  ],
  "metrics": [
    { "label": "Total Revenue", "header": "Revenue", "func": "Sum" }
  ]
}

Note for 'func': Allowed values are "Sum", "Count", "Average", "Min", "Max".`;
};