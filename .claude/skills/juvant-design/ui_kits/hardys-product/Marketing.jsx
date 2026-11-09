/* global React */
const { useState } = React;

const MarketingHero = () => (
  <section style={{ position: 'relative', minHeight: 720, overflow: 'hidden' }}>
    <DotField height={720}/>
    <div style={{ position: 'relative', maxWidth: 1100, margin: '0 auto',
      padding: '180px 32px 120px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 28 }}>
        <HardysMark size={40} animate/>
        <HdPill tone="cyan">Hardys · A Juvant product</HdPill>
      </div>
      <h1 style={{
        fontFamily: 'var(--font-display)', fontSize: 'clamp(52px, 7vw, 92px)',
        lineHeight: 1.02, letterSpacing: '-0.015em', margin: 0, color: '#E8EEFA', fontWeight: 400,
        maxWidth: 920,
      }}>
        Every student deserves<br/>a <HdGradText>Hardy</HdGradText>.
      </h1>
      <p style={{ fontFamily: 'var(--font-sans)', fontSize: 20, lineHeight: 1.65,
        color: '#94A3B8', margin: '26px 0 0', maxWidth: 640 }}>
        The Teaching Intelligence OS for the AI era. A <em>second professor</em> — in the room,
        directly with your students. Precise, patient, and never tired.
      </p>
      <div style={{ display: 'flex', gap: 14, marginTop: 42 }}>
        <HdButton size="lg" href="#contact">See Hardys in action →</HdButton>
        <HdButton size="lg" variant="ghost" href="#thesis">Read the rationale</HdButton>
      </div>
      <div style={{ marginTop: 80, display: 'flex', gap: 48, color: '#64748b',
        fontFamily: 'var(--font-mono)', fontSize: 11, letterSpacing: '0.08em' }}>
        <span>USED BY</span>
        {['UNIVERSITÀ BOCCONI', 'POLITECNICO DI MILANO', 'IE UNIVERSITY', 'CAMBRIDGE JUDGE'].map(u =>
          <span key={u} style={{ opacity: 0.6 }}>{u}</span>)}
      </div>
    </div>
  </section>
);

const Features = () => {
  const items = [
    { name: 'The Room', tag: '01', body: 'A virtual classroom where Hardys meets your students directly — precise, patient, and never tired. Voice, text, and whiteboard, all in one surface.' },
    { name: 'The Thread', tag: '02', body: 'Longitudinal learning memory. Hardys remembers every student across every session, and brings that context into the next one.' },
    { name: 'The Mirror', tag: '03', body: 'Real-time Teaching Intelligence for faculty. Shows you what you cannot see: which students are lost, which are ahead, and why.' },
    { name: 'The Lens', tag: '04', body: 'Institutional dashboard for program directors. Cohort-level signals, course health, longitudinal outcomes.' },
    { name: 'The Bridge', tag: '05', body: 'Connects non-didactic experts — alumni, guests, industry — directly to the learning experience, mediated by Hardys.' },
    { name: 'The Folio', tag: '06', body: 'Each student\'s evolving learning portrait. Not a transcript. A longitudinal intellectual biography.' },
  ];
  return (
    <section id="features" style={{ padding: '120px 32px', background: '#0B0F1A' }}>
      <div style={{ maxWidth: 1200, margin: '0 auto' }}>
        <HdEyebrow>THE PLATFORM</HdEyebrow>
        <h2 style={{ fontFamily: 'var(--font-display)', fontSize: 48, lineHeight: 1.08,
          margin: '20px 0 64px', color: '#E8EEFA', fontWeight: 400, letterSpacing: '-0.01em',
          maxWidth: 820 }}>
          Six surfaces. One continuous<br/><HdGradText>Teaching Intelligence</HdGradText>.
        </h2>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 20 }}>
          {items.map(it => (
            <div key={it.name} style={{
              background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
              borderRadius: 12, padding: 28, position: 'relative',
            }}>
              <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, color: '#64748b' }}>
                {it.tag}
              </div>
              <h3 style={{ fontFamily: 'var(--font-display)', fontSize: 26, lineHeight: 1.15,
                color: '#E8EEFA', margin: '16px 0 0', fontWeight: 400 }}>{it.name}</h3>
              <HdHair opacity={0.3} style={{ margin: '16px 0' }}/>
              <p style={{ fontFamily: 'var(--font-sans)', fontSize: 14, lineHeight: 1.65,
                color: '#94A3B8', margin: 0 }}>{it.body}</p>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
};

const QuoteBlock = () => (
  <section style={{ padding: '140px 32px', position: 'relative', overflow: 'hidden' }}>
    <DotField height={520}/>
    <div style={{ position: 'relative', maxWidth: 900, margin: '0 auto', textAlign: 'center' }}>
      <HdEyebrow>THE PROMISE</HdEyebrow>
      <blockquote style={{
        fontFamily: 'var(--font-display)', fontSize: 'clamp(36px, 4.6vw, 58px)',
        lineHeight: 1.15, color: '#E8EEFA', margin: '28px 0 0', fontStyle: 'italic',
        fontWeight: 400, letterSpacing: '-0.005em',
      }}>
        "Hardys is not an assistant. It is a <HdGradText>second professor</HdGradText> —
        in the room, directly with your students, building a longitudinal relationship with each of them."
      </blockquote>
      <div style={{ fontFamily: 'var(--font-sans)', fontSize: 13, color: '#94A3B8',
        marginTop: 32, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
        Antonio Gatti · Founder, Juvant Srls
      </div>
    </div>
  </section>
);

const MarketingCTA = () => (
  <section id="contact" style={{ padding: '100px 32px', background: '#0B0F1A', borderTop: '1px solid rgba(232,238,250,0.05)' }}>
    <div style={{ maxWidth: 1000, margin: '0 auto', display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 48, alignItems: 'center' }}>
      <div>
        <HdEyebrow>NEXT STEPS</HdEyebrow>
        <h2 style={{ fontFamily: 'var(--font-display)', fontSize: 44, lineHeight: 1.1,
          margin: '20px 0 16px', color: '#E8EEFA', fontWeight: 400 }}>
          Bring Hardys into your program.
        </h2>
        <p style={{ fontFamily: 'var(--font-sans)', fontSize: 17, lineHeight: 1.7,
          color: '#94A3B8', margin: 0, maxWidth: 560 }}>
          Built for senior faculty and program directors. We start with a single seminar,
          observe, and expand — never a generic rollout.
        </p>
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <HdButton size="lg">Talk to us</HdButton>
        <HdButton size="lg" variant="ghost">Download the paper</HdButton>
      </div>
    </div>
  </section>
);

const MarketingFooter = () => (
  <footer style={{ padding: '40px 32px', borderTop: '1px solid rgba(232,238,250,0.06)' }}>
    <div style={{ maxWidth: 1200, margin: '0 auto', display: 'flex',
      justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 20 }}>
        <img src="../../assets/hardys/hardys-logo-dark.png" alt="Hardys" style={{ height: 22 }}/>
        <span style={{ color: '#64748b', fontSize: 13 }}>· A Juvant product · hardys.ai</span>
      </div>
      <div style={{ color: '#475569', fontSize: 12 }}>© 2026 Juvant Srls</div>
    </div>
  </footer>
);

const Marketing = () => (
  <div style={{ background: '#080B14', minHeight: '100vh' }}>
    <MarketingHero/>
    <Features/>
    <QuoteBlock/>
    <MarketingCTA/>
    <MarketingFooter/>
  </div>
);

Object.assign(window, { Marketing });
