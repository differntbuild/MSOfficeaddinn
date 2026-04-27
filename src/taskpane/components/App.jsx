import React, { useState, useEffect, useRef } from 'react';
import ChatPane from './ChatPane';
import Settings from './Settings';
import AnalysisPane from './AnalysisPane';
import Skill from './Skill';
import '../components/taskpane.css';

const App = () => {
  const [selectedTab, setSelectedTab] = useState('chat');
  const [host, setHost] = useState('Excel');
  const chatRef = useRef(null);

  useEffect(() => {
    if (window.Office) {
      Office.onReady((info) => {
        if (info.host) {
          setHost(info.host === Office.HostType.Excel ? 'Excel' : 'Word');
        }
      });
    }
  }, []);

  const brandColor = host === 'Excel' ? '#217346' : '#2B579A';

  const handleRunSkill = (prompt) => {
    setSelectedTab('chat');
    setTimeout(() => chatRef.current?.sendMessage(prompt), 150);
  };

  const tabs = [
    { id: 'chat', label: 'Chat', icon: <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><path d="M2 3h12a1 1 0 011 1v7a1 1 0 01-1 1H9l-2 2-2-2H2a1 1 0 01-1-1V4a1 1 0 011-1z" stroke="currentColor" strokeWidth="1.3" fill="none"/></svg> },
    { id: 'analyze', label: 'Audit', icon: <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="currentColor" strokeWidth="1.3"/><path d="M5.5 8l2 2 3-3" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round"/></svg> },
    { id: 'skills', label: 'Skills', icon: <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><path d="M2 4h12M2 8h8M2 12h5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/></svg> },
    { id: 'settings', label: 'Setup', icon: <svg width="13" height="13" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="2" stroke="currentColor" strokeWidth="1.3"/><path d="M8 1v2M8 13v2M1 8h2M13 8h2M3.22 3.22l1.42 1.42M11.36 11.36l1.42 1.42M11.36 4.64l1.42-1.42M3.22 12.78l1.42-1.42" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/></svg> },
  ];

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', fontFamily: 'Segoe UI, system-ui, sans-serif', background: '#f8f8f8' }}>

      <div style={{ background: brandColor, flexShrink: 0 }}>
        <div style={{ padding: '10px 14px 0 14px', display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
          <div style={{ width: 22, height: 22, background: 'rgba(255,255,255,0.2)', borderRadius: 3, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, color: 'white' }}>
            {host === 'Excel' ? 'X' : 'W'}
          </div>
          <span style={{ fontSize: 13, fontWeight: 600, color: 'white' }}>AI Assistant</span>
          <span style={{ marginLeft: 'auto', fontSize: 10.5, background: 'rgba(255,255,255,0.18)', borderRadius: 3, padding: '1px 6px', color: 'rgba(255,255,255,0.9)', fontWeight: 500 }}>
            {host}
          </span>
        </div>

        <div style={{ display: 'flex' }}>
          {tabs.map(tab => (
            <button key={tab.id} onClick={() => setSelectedTab(tab.id)} style={{
              flex: 1, background: 'none', border: 'none',
              borderBottom: selectedTab === tab.id ? '2px solid white' : '2px solid transparent',
              color: selectedTab === tab.id ? 'white' : 'rgba(255,255,255,0.65)',
              padding: '8px 4px 9px', cursor: 'pointer',
              fontSize: 11.5, fontFamily: 'Segoe UI, system-ui, sans-serif',
              fontWeight: selectedTab === tab.id ? 600 : 400,
              display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3,
            }}>
              {tab.icon}
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
        <div style={{ display: selectedTab === 'chat' ? 'flex' : 'none', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
          <ChatPane ref={chatRef} host={host} />
        </div>
        <div style={{ display: selectedTab === 'analyze' ? 'flex' : 'none', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
          <AnalysisPane host={host} />
        </div>
        <div style={{ display: selectedTab === 'skills' ? 'flex' : 'none', flexDirection: 'column', flex: 1, overflow: 'auto' }}>
          <Skill onRunSkill={handleRunSkill} />
        </div>
        <div style={{ display: selectedTab === 'settings' ? 'flex' : 'none', flexDirection: 'column', flex: 1, overflow: 'auto' }}>
          <Settings />
        </div>
      </div>
    </div>
  );
};

export default App;