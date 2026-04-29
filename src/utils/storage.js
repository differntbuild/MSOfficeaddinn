export const saveApiKey = (key) => localStorage.setItem('groq_api_key', key);
export const getApiKey = () => localStorage.getItem('groq_api_key');

const DEFAULT_TOGGLES = {
  context: true,
  sheetNames: true,
  namedRanges: false,
  explanations: true,
  confirm: true,
  circular: true,
  hardcoded: true,
  dataTypes: false,
};

export const saveSettings = (settings) => {
  localStorage.setItem('groq_settings', JSON.stringify(settings));
};

export const getSettings = () => {
  const defaults = {
    model: 'auto',
    temperature: 0.7,
    contextMode: 'selection',
    contextLimits: {
      selectionCells: 200,
      sampleRows: 30,
      fullCells: 400,
    },
    toggles: DEFAULT_TOGGLES,
  };
  const saved = localStorage.getItem('groq_settings');
  if (!saved) return defaults;
  try {
    const parsed = JSON.parse(saved);
    return {
      ...defaults,
      ...parsed,
      toggles: { ...DEFAULT_TOGGLES, ...(parsed?.toggles || {}) },
      contextLimits: { ...defaults.contextLimits, ...(parsed?.contextLimits || {}) },
    };
  } catch {
    return defaults;
  }
};
