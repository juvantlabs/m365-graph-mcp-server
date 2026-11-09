/* global React */
const { useState } = React;

const AppShell = ({ active, onNav, children }) => {
  const items = [
    { id: 'room',   label: 'The Room',   icon: 'M4 7a2 2 0 012-2h12a2 2 0 012 2v10a2 2 0 01-2 2H6a2 2 0 01-2-2V7zm3 3h10M7 14h6' },
    { id: 'mirror', label: 'The Mirror', icon: 'M4 5h16v14H4z M4 12h16' },
    { id: 'lens',   label: 'The Lens',   icon: 'M4 14h3l3-8 4 14 3-10h3' },
    { id: 'folio',  label: 'The Folio',  icon: 'M5 5a2 2 0 012-2h7l5 5v11a2 2 0 01-2 2H7a2 2 0 01-2-2V5z M14 3v5h5' },
    { id: 'thread', label: 'The Thread', icon: 'M6 3v18 M18 3c-4 4-4 14 0 18' },
    { id: 'bridge', label: 'The Bridge', icon: 'M3 12h18 M5 12v-4 M19 12v-4 M9 12v-2 M15 12v-2' },
  ];
  return (
    <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr', minHeight: '100vh',
      background: '#0B0F1A', color: '#E8EEFA' }}>
      {/* Sidebar */}
      <aside style={{ background: '#080B14', borderRight: '1px solid rgba(232,238,250,0.05)',
        padding: '18px 14px', display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 8px' }}>
          <HardysMark size={28}/>
          <div>
            <div style={{ fontFamily: 'var(--font-wordmark)', fontWeight: 700, fontSize: 16,
              letterSpacing: 1.5,
              background: 'linear-gradient(90deg,#A855F7,#22D3EE)',
              WebkitBackgroundClip: 'text', backgroundClip: 'text', color: 'transparent' }}>
              HARDYS
            </div>
            <div style={{ fontSize: 10, color: '#475569', letterSpacing: '0.14em',
              textTransform: 'uppercase', marginTop: 1 }}>Teaching OS</div>
          </div>
        </div>
        <nav style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          {items.map(it => {
            const on = it.id === active;
            return (
              <button key={it.id} onClick={() => onNav(it.id)} style={{
                display: 'flex', alignItems: 'center', gap: 12,
                padding: '10px 12px', borderRadius: 8, border: 0, cursor: 'pointer',
                background: on ? 'rgba(168,85,247,0.12)' : 'transparent',
                color: on ? '#E8EEFA' : '#94A3B8',
                fontFamily: 'var(--font-sans)', fontSize: 14, fontWeight: on ? 600 : 500,
                textAlign: 'left', position: 'relative',
              }}>
                {on && <span style={{ position: 'absolute', left: 0, top: 8, bottom: 8, width: 2,
                  background: 'linear-gradient(180deg,#A855F7,#22D3EE)', borderRadius: 2 }}/>}
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                  strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round">
                  <path d={it.icon}/>
                </svg>
                {it.label}
              </button>
            );
          })}
        </nav>
        <div style={{ marginTop: 'auto', padding: 12, borderTop: '1px solid rgba(232,238,250,0.05)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 32, height: 32, borderRadius: 999,
              background: 'linear-gradient(135deg,#A855F7,#22D3EE)', display: 'flex',
              alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 13, color: '#fff' }}>
              AG
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: '#E8EEFA' }}>Prof. A. Gatti</div>
              <div style={{ fontSize: 11, color: '#64748b' }}>Bocconi · MSc · Sem I</div>
            </div>
          </div>
        </div>
      </aside>

      {/* Main */}
      <main style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
        <header style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: '16px 32px', borderBottom: '1px solid rgba(232,238,250,0.05)', background: '#080B14' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{ fontFamily: 'var(--font-sans)', fontSize: 12, color: '#64748b' }}>
              Fall 2026 / <span style={{ color: '#E8EEFA' }}>Econometrics II</span>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 12px',
              background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)', borderRadius: 8,
              color: '#94A3B8', fontSize: 13 }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/>
              </svg>
              <span style={{ fontFamily: 'var(--font-sans)' }}>Search anything</span>
              <span style={{ fontFamily: 'var(--font-mono)', fontSize: 10, marginLeft: 24,
                padding: '1px 6px', background: '#0B0F1A', borderRadius: 4, color: '#475569' }}>⌘ K</span>
            </div>
            <HdPill tone="cyan">● Live · 142 students</HdPill>
          </div>
        </header>
        <div style={{ flex: 1 }}>{children}</div>
      </main>
    </div>
  );
};

Object.assign(window, { AppShell });
