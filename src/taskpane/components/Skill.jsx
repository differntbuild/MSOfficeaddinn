import React, { useState, useEffect } from 'react';

const STORAGE_KEY = 'addin_skills';

const defaultSkills = [
  { id: 1, name: 'Grade Analysis', prompt: 'Add a grade column based on marks using A B C D F scale', icon: '🎓' },
  { id: 2, name: 'Sum Selected', prompt: 'Sum the selected cells and insert total below', icon: '∑' },
  { id: 3, name: 'Full Audit', prompt: 'Audit this sheet and find all errors and warnings', icon: '🔍' },
  { id: 4, name: 'Data Summary', prompt: 'Analyze this data and give me key insights and statistics', icon: '📊' },
];

export default function Skill({ onRunSkill }) {
  const [skills, setSkills] = useState(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : defaultSkills;
    } catch { return defaultSkills; }
  });
  const [newName, setNewName] = useState('');
  const [newPrompt, setNewPrompt] = useState('');
  const [isAdding, setIsAdding] = useState(false);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(skills));
  }, [skills]);

  const addSkill = () => {
    if (!newName.trim() || !newPrompt.trim()) return;
    setSkills(prev => [...prev, {
      id: Date.now(),
      name: newName.trim(),
      prompt: newPrompt.trim(),
      icon: '⚡',
    }]);
    setNewName('');
    setNewPrompt('');
    setIsAdding(false);
  };

  const deleteSkill = (id) => setSkills(prev => prev.filter(s => s.id !== id));

  return (
    <div style={{ padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: '#605e5c', letterSpacing: '0.05em', textTransform: 'uppercase', marginBottom: 2 }}>
        Saved Skills
      </div>

      {skills.map(skill => (
        <div key={skill.id} style={{ background: 'white', border: '0.5px solid #e0e0e0', borderRadius: 6, padding: '9px 11px', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 16, flexShrink: 0 }}>{skill.icon}</span>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: '#201f1e' }}>{skill.name}</div>
            <div style={{ fontSize: 11, color: '#a09e9c', marginTop: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 160 }}>{skill.prompt}</div>
          </div>
          <div style={{ display: 'flex', gap: 4, flexShrink: 0 }}>
            <button onClick={() => onRunSkill(skill.prompt)} style={{ background: '#217346', color: 'white', border: 'none', borderRadius: 4, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
              Run
            </button>
            <button onClick={() => deleteSkill(skill.id)} style={{ background: 'none', color: '#a09e9c', border: '0.5px solid #e0e0e0', borderRadius: 4, padding: '4px 7px', fontSize: 11, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
              ✕
            </button>
          </div>
        </div>
      ))}

      {isAdding ? (
        <div style={{ background: 'white', border: '0.5px solid #217346', borderRadius: 6, padding: '10px' }}>
          <input value={newName} onChange={e => setNewName(e.target.value)} placeholder="Skill name" style={inputStyle} />
          <textarea value={newPrompt} onChange={e => setNewPrompt(e.target.value)} placeholder="What should this skill do?" rows={2} style={{ ...inputStyle, resize: 'none', marginTop: 6 }} />
          <div style={{ display: 'flex', gap: 6, marginTop: 8 }}>
            <button onClick={addSkill} style={{ background: '#217346', color: 'white', border: 'none', borderRadius: 4, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Save</button>
            <button onClick={() => setIsAdding(false)} style={{ background: 'none', color: '#605e5c', border: '0.5px solid #d0d0d0', borderRadius: 4, padding: '5px 10px', fontSize: 12, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button onClick={() => setIsAdding(true)} style={{ background: 'none', border: '0.5px dashed #d0d0d0', borderRadius: 6, padding: '8px', fontSize: 12, color: '#605e5c', cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
          + Save new skill
        </button>
      )}
    </div>
  );
}

const inputStyle = {
  width: '100%', border: '0.5px solid #d0d0d0', borderRadius: 4,
  padding: '6px 9px', fontSize: 12, fontFamily: 'Segoe UI, system-ui, sans-serif',
  outline: 'none', boxSizing: 'border-box', background: '#f8f8f8',
};