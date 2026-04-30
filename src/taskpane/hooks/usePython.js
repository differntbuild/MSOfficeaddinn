import { useState, useCallback, useRef } from 'react';

/**
 * usePython - A hook to manage the local Pyodide runtime for advanced data analysis.
 * This runs a full Python environment (with pandas/numpy) directly in the user's browser.
 */
export const usePython = () => {
  const [isReady, setIsReady] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const pyodideRef = useRef(null);

  const loadRuntime = useCallback(async () => {
    if (pyodideRef.current || isLoading) return;
    
    setIsLoading(true);
    try {
      // 1. Initialize the Pyodide runtime
      // @ts-ignore - pyodide is loaded via CDN in index.html
      const py = await window.loadPyodide();
      
      // 2. Load the analytical stack (Pandas + Dependencies)
      await py.loadPackage(['pandas', 'numpy']);
      
      pyodideRef.current = py;
      setIsReady(true);
    } catch (err) {
      console.error('Failed to load Pyodide:', err);
      setError('Failed to initialize Python runtime.');
    } finally {
      setIsLoading(false);
    }
  }, [isLoading]);

  const runAnalysis = useCallback(async (pythonCode, data) => {
    if (!pyodideRef.current) throw new Error('Python runtime not initialized');

    const py = pyodideRef.current;

    try {
      // 1. Transfer data to Python as a CSV string (robust escaping)
      const csvData = [
        data.headers.map(h => `"${h.toString().replace(/"/g, '""')}"`).join(','),
        ...data.values.slice(1).map(row => 
          row.map(val => {
            if (val === null || val === undefined) return '';
            const str = val.toString().replace(/"/g, '""');
            return `"${str}"`;
          }).join(',')
        )
      ].join('\n');

      py.globals.set('raw_csv', csvData);

      // 2. Setup the analysis environment
      const wrapperCode = `
import pandas as pd
import io
import json

# Load data from the JS transfer variable
df = pd.read_csv(io.StringIO(raw_csv))

# The user-generated code from the LLM goes here
def execute_analysis(df):
    ${pythonCode.split('\n').map(line => '    ' + line).join('\n')}

# Run it and capture the result
try:
    results = execute_analysis(df)
    # Ensure results is a valid JSON string
    output_json = json.dumps(results, default=str)
except Exception as e:
    output_json = json.dumps({"error": str(e)})

output_json
`;

      const resultJson = await py.runPythonAsync(wrapperCode);
      return JSON.parse(resultJson);
    } catch (err) {
      console.error('Python execution failed:', err);
      return { error: err.message };
    }
  }, []);

  return { isReady, isLoading, error, loadRuntime, runAnalysis };
};
