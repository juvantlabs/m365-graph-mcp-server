/* global React */
const { useState } = React;

// ------------- THE ROOM -------------
const RoomScreen = () => {
  const [input, setInput] = useState('');
  const [msgs, setMsgs] = useState([
    { who: 'hardys', text: 'Good morning. We left off at heteroskedasticity-robust standard errors. Would you like to start with the Wednesday problem set, or the section you flagged on Tuesday?' },
    { who: 'you', text: 'Start with the flagged section. I think the class got lost on the White test.' },
    { who: 'hardys', text: 'Understood. 42 of 142 students showed confusion at the 3:14 mark — three different misconceptions. I\'ll walk the class through each, then check for residual issues. Ready when you are.' },
  ]);
  const send = () => {
    if (!input.trim()) return;
    setMsgs([...msgs, { who: 'you', text: input }, { who: 'hardys', text: 'Noted — I\'ll build a short exercise around that and surface results in The Mirror when it runs.' }]);
    setInput('');
  };
  return (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 320px', height: 'calc(100vh - 64px)' }}>
      <div style={{ display: 'flex', flexDirection: 'column', padding: '28px 32px', minHeight: 0 }}>
        <div>
          <HdEyebrow>THE ROOM · LIVE SESSION</HdEyebrow>
          <h1 style={{ fontFamily: 'var(--font-display)', fontSize: 34, lineHeight: 1.1,
            margin: '12px 0 4px', color: '#E8EEFA', fontWeight: 400 }}>
            Wednesday · Heteroskedasticity & robust SE
          </h1>
          <div style={{ color: '#64748b', fontSize: 13 }}>142 students in the room · 38 min elapsed</div>
        </div>
        <div style={{ flex: 1, marginTop: 24, overflowY: 'auto', display: 'flex',
          flexDirection: 'column', gap: 16, paddingRight: 8 }}>
          {msgs.map((m, i) => (
            <div key={i} style={{ display: 'flex', gap: 12, alignItems: 'flex-start',
              maxWidth: 720, alignSelf: m.who === 'you' ? 'flex-end' : 'flex-start',
              flexDirection: m.who === 'you' ? 'row-reverse' : 'row' }}>
              {m.who === 'hardys' ? (
                <HardysMark size={32}/>
              ) : (
                <div style={{ width: 32, height: 32, borderRadius: 999,
                  background: 'linear-gradient(135deg,#A855F7,#22D3EE)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: 12, fontWeight: 700, color: '#fff' }}>AG</div>
              )}
              <div style={{
                background: m.who === 'you' ? 'rgba(168,85,247,0.12)' : '#161D2C',
                border: `1px solid ${m.who === 'you' ? 'rgba(168,85,247,0.25)' : 'rgba(232,238,250,0.06)'}`,
                borderRadius: 12, padding: '14px 18px', fontSize: 15, lineHeight: 1.6, color: '#E8EEFA',
              }}>{m.text}</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop: 16, display: 'flex', gap: 10, alignItems: 'center',
          background: '#161D2C', border: '1px solid rgba(232,238,250,0.08)',
          borderRadius: 12, padding: '10px 10px 10px 18px' }}>
          <input value={input} onChange={e => setInput(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && send()}
            placeholder="Ask Hardys, or cue a section…"
            style={{ flex: 1, background: 'transparent', border: 0, outline: 0,
              color: '#E8EEFA', fontFamily: 'var(--font-sans)', fontSize: 15 }}/>
          <HdButton size="sm" onClick={send}>Send →</HdButton>
        </div>
      </div>
      {/* Live roster */}
      <aside style={{ borderLeft: '1px solid rgba(232,238,250,0.05)', background: '#080B14',
        padding: '28px 24px', overflowY: 'auto' }}>
        <HdEyebrow>ATTENTION MAP</HdEyebrow>
        <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 14 }}>
          {[
            { name: 'Chen, M.', signal: 'engaged', val: 0.94 },
            { name: 'Rossi, G.', signal: 'engaged', val: 0.88 },
            { name: 'Al-Farsi, L.', signal: 'lost', val: 0.32 },
            { name: 'Bianchi, E.', signal: 'engaged', val: 0.81 },
            { name: 'Okoye, D.', signal: 'ahead', val: 1.0 },
            { name: 'Park, S.', signal: 'lost', val: 0.41 },
          ].map(s => {
            const c = s.signal === 'lost' ? '#A855F7' : s.signal === 'ahead' ? '#22D3EE' : '#94A3B8';
            return (
              <div key={s.name} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <div style={{ width: 24, height: 24, borderRadius: 999,
                  background: '#161D2C', border: `1px solid ${c}`,
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: 9, color: c, fontWeight: 700 }}>
                  {s.name.split(',')[0][0]}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, color: '#E8EEFA' }}>{s.name}</div>
                  <div style={{ height: 2, marginTop: 4, background: 'rgba(232,238,250,0.06)',
                    borderRadius: 999, overflow: 'hidden' }}>
                    <div style={{ height: '100%', width: `${s.val * 100}%`, background: c }}/>
                  </div>
                </div>
                <div style={{ fontFamily: 'var(--font-mono)', fontSize: 10, color: c,
                  textTransform: 'uppercase', letterSpacing: '0.1em' }}>{s.signal}</div>
              </div>
            );
          })}
        </div>
        <HdHair opacity={0.25} style={{ margin: '28px 0 20px' }}/>
        <HdEyebrow color="#22D3EE">THE MIRROR</HdEyebrow>
        <div style={{ fontSize: 13, lineHeight: 1.6, color: '#94A3B8', marginTop: 10 }}>
          Two misconception clusters detected around <span style={{ color: '#E8EEFA' }}>White test interpretation</span>.
          I can run a 3-min clarification now, or queue it for Friday.
        </div>
        <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
          <HdButton size="sm">Run now</HdButton>
          <HdButton size="sm" variant="ghost">Queue Friday</HdButton>
        </div>
      </aside>
    </div>
  );
};

// ------------- THE MIRROR -------------
const MirrorScreen = () => {
  const bars = [68, 72, 61, 55, 42, 38, 52, 71, 84, 78, 70, 63, 58, 66, 79, 88, 82, 75];
  return (
    <div style={{ padding: '28px 32px' }}>
      <HdEyebrow>THE MIRROR · REAL-TIME TEACHING INTELLIGENCE</HdEyebrow>
      <h1 style={{ fontFamily: 'var(--font-display)', fontSize: 36, lineHeight: 1.1,
        margin: '14px 0 26px', color: '#E8EEFA', fontWeight: 400 }}>
        What you cannot see.
      </h1>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 20 }}>
        {[
          { k: '142', l: 'In the room', d: '+6 vs last week', tone: '#22D3EE' },
          { k: '87%', l: 'Engagement', d: 'above cohort median', tone: '#22D3EE' },
          { k: '12', l: 'Students lost', d: 'flagged for follow-up', tone: '#A855F7' },
          { k: '3', l: 'Misconception clusters', d: 'across today\'s session', tone: '#A855F7' },
        ].map(m => (
          <div key={m.l} style={{ background: '#161D2C',
            border: '1px solid rgba(232,238,250,0.06)', borderRadius: 12, padding: 20 }}>
            <div style={{ fontFamily: 'var(--font-display)', fontSize: 38, lineHeight: 1,
              color: m.tone }}>{m.k}</div>
            <div style={{ fontSize: 13, color: '#E8EEFA', marginTop: 10 }}>{m.l}</div>
            <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{m.d}</div>
          </div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16 }}>
        <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
          borderRadius: 12, padding: 24 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between' }}>
            <div>
              <HdEyebrow>SESSION ENGAGEMENT · MIN-BY-MIN</HdEyebrow>
              <div style={{ fontFamily: 'var(--font-display)', fontSize: 22, color: '#E8EEFA',
                marginTop: 8 }}>Wednesday · 09:00 – 10:30</div>
            </div>
            <HdPill tone="cyan">Live</HdPill>
          </div>
          <div style={{ marginTop: 28, display: 'flex', alignItems: 'flex-end',
            gap: 4, height: 140, paddingBottom: 8,
            borderBottom: '1px solid rgba(232,238,250,0.06)' }}>
            {bars.map((h, i) => (
              <div key={i} style={{ flex: 1, height: `${h}%`,
                background: h < 45 ? 'rgba(168,85,247,0.55)' :
                  'linear-gradient(180deg, rgba(34,211,238,0.7), rgba(34,211,238,0.2))',
                borderRadius: '3px 3px 0 0' }}/>
            ))}
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 10,
            fontSize: 11, color: '#475569', fontFamily: 'var(--font-mono)' }}>
            <span>09:00</span><span>09:20</span><span>09:40</span><span>10:00</span><span>10:20</span>
          </div>
        </div>
        <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
          borderRadius: 12, padding: 24 }}>
          <HdEyebrow>MISCONCEPTION CLUSTERS</HdEyebrow>
          <div style={{ marginTop: 16, display: 'flex', flexDirection: 'column', gap: 14 }}>
            {[
              { t: 'White test vs Breusch–Pagan', n: 17, p: 'high' },
              { t: 'Robust SE under small n', n: 14, p: 'med' },
              { t: 'Clustered errors rationale', n: 11, p: 'med' },
            ].map((c, i) => (
              <div key={i}>
                <div style={{ display: 'flex', justifyContent: 'space-between',
                  fontSize: 13, color: '#E8EEFA' }}>
                  <span>{c.t}</span>
                  <span style={{ color: c.p === 'high' ? '#A855F7' : '#22D3EE',
                    fontFamily: 'var(--font-mono)', fontSize: 11 }}>{c.n} students</span>
                </div>
                <div style={{ height: 3, marginTop: 6, background: 'rgba(232,238,250,0.05)',
                  borderRadius: 999, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${c.n * 5}%`,
                    background: c.p === 'high' ? '#A855F7' : '#22D3EE' }}/>
                </div>
              </div>
            ))}
          </div>
          <HdHair opacity={0.25} style={{ margin: '22px 0 16px' }}/>
          <HdButton size="sm" variant="outline">Draft a clarification →</HdButton>
        </div>
      </div>
    </div>
  );
};

// ------------- THE LENS -------------
const LensScreen = () => (
  <div style={{ padding: '28px 32px' }}>
    <HdEyebrow>THE LENS · INSTITUTIONAL VIEW</HdEyebrow>
    <h1 style={{ fontFamily: 'var(--font-display)', fontSize: 36, lineHeight: 1.1,
      margin: '14px 0 26px', color: '#E8EEFA', fontWeight: 400 }}>
      Fall 2026 · 14 courses · 2,410 students.
    </h1>
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
      <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
        borderRadius: 12, padding: 24 }}>
        <HdEyebrow>COURSE HEALTH</HdEyebrow>
        <div style={{ marginTop: 18, display: 'flex', flexDirection: 'column', gap: 10 }}>
          {[
            { c: 'Econometrics II', p: 'Gatti', h: 92, s: 'strong' },
            { c: 'Macro Theory', p: 'Moretti', h: 78, s: 'strong' },
            { c: 'Applied Statistics', p: 'Verdi', h: 64, s: 'watch' },
            { c: 'Advanced Corporate Finance', p: 'Russo', h: 54, s: 'watch' },
            { c: 'Game Theory', p: 'Bianchi', h: 41, s: 'at-risk' },
          ].map(r => {
            const c = r.s === 'at-risk' ? '#A855F7' : r.s === 'watch' ? '#E8EEFA' : '#22D3EE';
            return (
              <div key={r.c} style={{ display: 'grid', gridTemplateColumns: '1.8fr 1fr 120px 80px',
                alignItems: 'center', gap: 12, padding: '10px 0',
                borderBottom: '1px solid rgba(232,238,250,0.04)' }}>
                <div style={{ fontSize: 14, color: '#E8EEFA' }}>{r.c}</div>
                <div style={{ fontSize: 13, color: '#94A3B8' }}>Prof. {r.p}</div>
                <div style={{ height: 4, background: 'rgba(232,238,250,0.05)',
                  borderRadius: 999, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${r.h}%`, background: c }}/>
                </div>
                <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, color: c,
                  textTransform: 'uppercase', letterSpacing: '0.12em', textAlign: 'right' }}>{r.s}</div>
              </div>
            );
          })}
        </div>
      </div>
      <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
        borderRadius: 12, padding: 24 }}>
        <HdEyebrow>LONGITUDINAL COHORT · MSc ECON 2024 → 2026</HdEyebrow>
        <svg width="100%" height="220" viewBox="0 0 400 220" style={{ marginTop: 12 }}>
          <defs>
            <linearGradient id="lcurve" x1="0" y1="0" x2="1" y2="0">
              <stop offset="0%" stopColor="#A855F7"/>
              <stop offset="100%" stopColor="#22D3EE"/>
            </linearGradient>
            <linearGradient id="lfill" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#22D3EE" stopOpacity="0.25"/>
              <stop offset="100%" stopColor="#22D3EE" stopOpacity="0"/>
            </linearGradient>
          </defs>
          {[40, 80, 120, 160, 200].map(y =>
            <line key={y} x1="30" y1={y} x2="390" y2={y} stroke="rgba(232,238,250,0.04)"/>
          )}
          <path d="M30,170 C80,150 120,120 170,130 S260,80 310,70 390,40 390,40"
            stroke="url(#lcurve)" strokeWidth="2.5" fill="none"/>
          <path d="M30,170 C80,150 120,120 170,130 S260,80 310,70 390,40 390,40 L390,210 L30,210 Z"
            fill="url(#lfill)"/>
          <circle cx="170" cy="130" r="3" fill="#A855F7"/>
          <circle cx="310" cy="70" r="3" fill="#22D3EE"/>
          <text x="30" y="208" fill="#64748b" fontSize="10" fontFamily="JetBrains Mono">'24 Q3</text>
          <text x="175" y="208" fill="#64748b" fontSize="10" fontFamily="JetBrains Mono">'25 Q3</text>
          <text x="340" y="208" fill="#64748b" fontSize="10" fontFamily="JetBrains Mono">'26 Q3</text>
        </svg>
        <div style={{ fontSize: 13, color: '#94A3B8', lineHeight: 1.6, marginTop: 8 }}>
          Retention in advanced electives up <span style={{ color: '#22D3EE' }}>+18pt</span> in
          cohorts that used Hardys from Y1.
        </div>
      </div>
    </div>
  </div>
);

// ------------- THE FOLIO -------------
const FolioScreen = () => (
  <div style={{ padding: '28px 32px' }}>
    <HdEyebrow>THE FOLIO · STUDENT</HdEyebrow>
    <div style={{ display: 'flex', alignItems: 'center', gap: 18, margin: '14px 0 28px' }}>
      <div style={{ width: 64, height: 64, borderRadius: 999,
        background: 'linear-gradient(135deg,#A855F7,#22D3EE)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: '#fff', fontWeight: 700, fontSize: 22 }}>LA</div>
      <div>
        <h1 style={{ fontFamily: 'var(--font-display)', fontSize: 34, lineHeight: 1.1,
          margin: 0, color: '#E8EEFA', fontWeight: 400 }}>Leila Al-Farsi</h1>
        <div style={{ color: '#94A3B8', fontSize: 13, marginTop: 4 }}>
          MSc Economics · Year 2 · 8 courses · Thread open since Sep 2024
        </div>
      </div>
      <div style={{ marginLeft: 'auto', display: 'flex', gap: 10 }}>
        <HdButton size="sm" variant="outline">Open The Thread</HdButton>
        <HdButton size="sm">Draft office-hours note</HdButton>
      </div>
    </div>
    <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: 16 }}>
      <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
        borderRadius: 12, padding: 24 }}>
        <HdEyebrow>INTELLECTUAL PORTRAIT</HdEyebrow>
        <p style={{ fontFamily: 'Georgia, serif', fontStyle: 'italic', fontSize: 17,
          lineHeight: 1.6, color: '#E8EEFA', margin: '14px 0 0' }}>
          "Exceptional at causal reasoning when given concrete settings; freezes on abstract derivations.
          Leila consistently returns to her earlier misconceptions before resolving them — a slow, deep learner.
          Pairs well with Chen, M. (complementary style)."
        </p>
        <div style={{ fontSize: 12, color: '#64748b', marginTop: 14 }}>
          — Hardys summary · updated this morning · 17 sessions analysed
        </div>
        <HdHair opacity={0.25} style={{ margin: '24px 0 18px' }}/>
        <HdEyebrow>RECENT TURNING POINTS</HdEyebrow>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginTop: 14 }}>
          {[
            { d: 'Oct 12', t: 'First independent application of IV estimation — unprompted.', tone: '#22D3EE' },
            { d: 'Sep 24', t: 'Breakthrough on confounding vs mediation after 3 prior stalls.', tone: '#22D3EE' },
            { d: 'Sep 08', t: 'Flagged lost during White test — 3 clarification cycles before resolved.', tone: '#A855F7' },
          ].map((e, i) => (
            <div key={i} style={{ display: 'flex', gap: 14 }}>
              <div style={{ width: 44, fontFamily: 'var(--font-mono)', fontSize: 11, color: '#64748b' }}>{e.d}</div>
              <div style={{ width: 2, background: e.tone, borderRadius: 2 }}/>
              <div style={{ fontSize: 13, color: '#E8EEFA', lineHeight: 1.6 }}>{e.t}</div>
            </div>
          ))}
        </div>
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
          borderRadius: 12, padding: 24 }}>
          <HdEyebrow>SIGNAL</HdEyebrow>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginTop: 14 }}>
            {[
              { k: 'Causal reasoning', v: 88, tone: '#22D3EE' },
              { k: 'Formal derivation', v: 42, tone: '#A855F7' },
              { k: 'Retention (4-wk)', v: 76, tone: '#22D3EE' },
              { k: 'Engagement', v: 83, tone: '#22D3EE' },
            ].map(x => (
              <div key={x.k}>
                <div style={{ fontSize: 12, color: '#94A3B8' }}>{x.k}</div>
                <div style={{ fontFamily: 'var(--font-display)', fontSize: 24, color: x.tone,
                  marginTop: 4 }}>{x.v}</div>
                <div style={{ height: 2, background: 'rgba(232,238,250,0.05)',
                  borderRadius: 999, overflow: 'hidden', marginTop: 6 }}>
                  <div style={{ height: '100%', width: `${x.v}%`, background: x.tone }}/>
                </div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ background: 'linear-gradient(135deg, rgba(168,85,247,0.08), rgba(34,211,238,0.04))',
          border: '1px solid rgba(168,85,247,0.18)', borderRadius: 12, padding: 24 }}>
          <HdEyebrow color="#22D3EE">HARDYS SUGGESTS</HdEyebrow>
          <p style={{ fontSize: 14, lineHeight: 1.6, color: '#E8EEFA', margin: '12px 0 16px' }}>
            A 10-minute office-hours conversation on Friday could consolidate the IV breakthrough
            before the midterm. I can prepare a 3-problem warm-up sequence.
          </p>
          <HdButton size="sm" variant="outline">Accept & draft</HdButton>
        </div>
      </div>
    </div>
  </div>
);

Object.assign(window, { RoomScreen, MirrorScreen, LensScreen, FolioScreen });
