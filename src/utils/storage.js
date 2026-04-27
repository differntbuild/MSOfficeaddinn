export const saveApiKey = (key) => localStorage.setItem('groq_api_key', key);
export const getApiKey = () => localStorage.getItem('groq_api_key');

export const saveSettings = (settings) => {
  localStorage.setItem('groq_settings', JSON.stringify(settings));
};

export const getSettings = () => {
  const defaults = { model: 'groq/compound', temperature: 0.7 };
  const saved = localStorage.getItem('groq_settings');
  return saved ? JSON.parse(saved) : defaults;
};