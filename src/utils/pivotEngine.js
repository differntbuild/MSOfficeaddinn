export const getVisualizationPrompt = (headers, request) => {
  return `You are an Excel Data Architect. The user wants to build a pivot table or chart based on their data.
  
Available Headers: ${JSON.stringify(headers)}
User Request: "${request}"

CRITICAL INSTRUCTIONS:
1. Determine the best structure for the Pivot Table to answer the request.
2. Choose the best chart type (ColumnClustered, BarClustered, Line, Pie, Doughnut).
3. Return ONLY valid JSON. No markdown backticks, no explanations.

Expected JSON Format:
{
  "name": "AIDashboard",
  "rows": ["Exact Header Name"],
  "values": [{"header": "Exact Header Name", "func": "Sum"}], 
  "chartType": "ColumnClustered" 
}

Note for 'func': Allowed values are "Sum", "Count", or "Average".`;
};