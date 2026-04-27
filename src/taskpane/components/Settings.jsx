import React, { useState, useEffect } from 'react';
import { getSettings, saveSettings } from '../../utils/storage';

const Settings = () => {
  const [apiKey, setApiKey] = useState('');
  const [model, setModel] = useState('groq/compound');
  const [temperature, setTemperature] = useState(0.7);
  const [isSaved, setIsSaved] = useState(false);
  const [toggles, setToggles] = useState({
    context: true,
    sheetNames: true,
    namedRanges: false,
    explanations: true,
    confirm: true,
    circular: true,
    hardcoded: true,
    dataTypes: false,
  });

  useEffect(() => {
    const savedKey = localStorage.getItem('groq_api_key');
    if (savedKey) setApiKey(savedKey);
    
    const settings = getSettings();
    setModel(settings.model);
    setTemperature(settings.temperature);
  }, []);

  const handleSave = () => {
    localStorage.setItem('groq_api_key', apiKey);
    saveSettings({ model, temperature });
    setIsSaved(true);
    setTimeout(() => setIsSaved(false), 2000);
  };

  const toggle = (key) => {
    setToggles(prev => ({ ...prev, [key]: !prev[key] }));
  };

  return (
    <div className="pane active" style={{ flexDirection: 'column', background: 'white' }}>
      <div className="settings-body scroll-area" style={{ flex: 1, overflowY: 'auto', paddingTop: 20 }}>
        
        <div className="settings-section">
          <div className="settings-section-head">Model & API</div>
          
          <div className="settings-row">
            <div>
              <div className="settings-row-label">Intelligence Model</div>
              <div className="settings-row-sub">Choose your LLM engine</div>
            </div>
          </div>
          <div className="settings-row" style={{ borderBottom: 'none', paddingTop: 4 }}>
            <select 
              className="settings-input"
              value={model}
              onChange={(e) => setModel(e.target.value)}
              style={{ padding: '8px', cursor: 'pointer' }}
            >
              <optgroup label="Groq Optimization">
                <option value="groq/compound">Groq Compound (New)</option>
              </optgroup>
              <optgroup label="Meta Llama">
                <option value="llama-3.3-70b-versatile">Llama 3.3 70B (Versatile)</option>
                <option value="llama-3.1-8b-instant">Llama 3.1 8B (Fast)</option>
              </optgroup>
              <optgroup label="Mixtral (Mistral AI)">
                <option value="mixtral-8x7b-32768">Mixtral 8x7B (32k)</option>
              </optgroup>
            </select>
          </div>

          <div className="settings-row" style={{ marginTop: 8 }}>
            <div>
              <div className="settings-row-label">API Key</div>
              <div className="settings-row-sub">Groq API Key (gsk_...)</div>
            </div>
            <span className={`settings-tag ${apiKey ? 'tag-connected' : 'tag-pro'}`}>
              {apiKey ? 'Configured' : 'Missing'}
            </span>
          </div>
          <div className="settings-row" style={{ borderBottom: 'none', paddingTop: 4 }}>
             <input 
              className="settings-input" 
              type="password" 
              placeholder="gsk_..." 
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
            />
          </div>
        </div>

        <div className="settings-section">
          <div className="settings-section-head">Context Integration</div>
          <div className="settings-row">
            <div>
              <div className="settings-row-label">Selection Context</div>
              <div className="settings-row-sub">Automatically include active selection</div>
            </div>
            <div className={`toggle ${toggles.context ? 'on' : ''}`} onClick={() => toggle('context')}></div>
          </div>
          <div className="settings-row">
            <div>
              <div className="settings-row-label">Knowledge of Sheets</div>
              <div className="settings-row-sub">Allow AI to see names of other worksheets</div>
            </div>
            <div className={`toggle ${toggles.sheetNames ? 'on' : ''}`} onClick={() => toggle('sheetNames')}></div>
          </div>
        </div>

        <div className="settings-section">
          <div className="settings-section-head">Intelligence Behaviour</div>
          <div className="settings-row">
            <div>
              <div className="settings-row-label">Proactive Explanations</div>
              <div className="settings-row-sub">Explain formulas after insertion</div>
            </div>
            <div className={`toggle ${toggles.explanations ? 'on' : ''}`} onClick={() => toggle('explanations')}></div>
          </div>
          <div className="settings-row">
            <div>
              <div className="settings-row-label">Write Confirmation</div>
              <div className="settings-row-sub">Always ask before overwriting cells</div>
            </div>
            <div className={`toggle ${toggles.confirm ? 'on' : ''}`} onClick={() => toggle('confirm')}></div>
          </div>
        </div>

        <div className="settings-section">
          <div className="settings-section-head">Audit Engine Defaults</div>
          <div className="settings-row">
            <div className="settings-row-label">Circular Reference Checks</div>
            <div className={`toggle ${toggles.circular ? 'on' : ''}`} onClick={() => toggle('circular')}></div>
          </div>
          <div className="settings-row">
            <div className="settings-row-label">Hardcoded Value Flags</div>
            <div className={`toggle ${toggles.hardcoded ? 'on' : ''}`} onClick={() => toggle('hardcoded')}></div>
          </div>
        </div>

        <div style={{ padding: '0 14px 24px' }}>
          <button 
            className="api-save-btn" 
            onClick={handleSave}
            style={{ 
              width: '100%', 
              background: '#217346', 
              color: 'white', 
              border: 'none', 
              borderRadius: 6, 
              padding: '12px', 
              fontSize: 13, 
              fontWeight: 600, 
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
            }}
          >
            {isSaved ? 'Preferences Saved!' : 'Save All Preferences'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default Settings;