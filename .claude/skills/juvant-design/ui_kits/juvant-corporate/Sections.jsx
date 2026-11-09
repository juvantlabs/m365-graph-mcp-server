/* global React */

const Thesis = () => (
  <section id="thesis" style={{ position: 'relative', padding: '120px 32px', maxWidth: 1200, margin: '0 auto' }}>
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: 80, alignItems: 'start' }}>
      <div>
        <Eyebrow>01 — THE THESIS</Eyebrow>
        <h2 style={{
          fontFamily: 'var(--font-display)', fontSize: 44, lineHeight: 1.1,
          margin: '22px 0 0', color: '#E8EEFA', fontWeight: 400, letterSpacing: '-0.01em',
        }}>
          Name the enabler,<br/>not the genius.
        </h2>
      </div>
      <div>
        <p style={{ fontFamily: 'var(--font-sans)', fontSize: 18, lineHeight: 1.75, color: '#E8EEFA', margin: 0 }}>
          Every Juvant product is named after a real historical figure — not the brilliant one,
          but the person who stood beside greatness and <em>made space for it</em>.
        </p>
        <p style={{ fontFamily: 'var(--font-sans)', fontSize: 16, lineHeight: 1.75, color: '#94A3B8', marginTop: 20 }}>
          This is not a branding flourish. It is our product philosophy, compressed into a single word.
          An agent's job is not to replace the human in the room — it is to recognise what is already there,
          and make it bigger.
        </p>
        <GradientHairline opacity={0.35} style={{ margin: '40px 0' }}/>
        <blockquote style={{
          fontFamily: 'Georgia, serif', fontStyle: 'italic', fontSize: 26, lineHeight: 1.3,
          color: '#E8EEFA', margin: 0, maxWidth: 640,
        }}>
          "Intelligence already exists. Our agents help it emerge."
        </blockquote>
      </div>
    </div>
  </section>
);

const Portfolio = () => {
  const products = [
    {
      status: 'active', name: 'Hardys', domain: 'hardys.ai',
      namesake: 'G. H. Hardy · recognised Ramanujan',
      desc: 'The Teaching Intelligence OS. A second professor — in the room, with your students.',
      href: '../hardys-product/index.html',
    },
    { status: 'pipeline', name: 'Vogts', namesake: 'Heinrich Vogt · mentored Cecilia Payne', desc: 'Research / Science amplification.' },
    { status: 'pipeline', name: 'Ruskins', namesake: 'John Ruskin · championed Turner', desc: 'Creative / Design intelligence.' },
    { status: 'pipeline', name: 'Flexners', namesake: 'Abraham Flexner · gave Einstein the IAS', desc: 'Environments where thought happens.' },
  ];
  return (
    <section id="portfolio" style={{ position: 'relative', padding: '120px 32px', background: '#0B0F1A' }}>
      <div style={{ maxWidth: 1200, margin: '0 auto' }}>
        <Eyebrow>02 — PORTFOLIO</Eyebrow>
        <h2 style={{
          fontFamily: 'var(--font-display)', fontSize: 44, lineHeight: 1.1,
          margin: '22px 0 64px', color: '#E8EEFA', fontWeight: 400, maxWidth: 680,
        }}>
          A platform of agents — each<br/>named after an <GradientText>enabler</GradientText>.
        </h2>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {products.map((p) => (
            <a key={p.name} href={p.href || '#'} style={{
              display: 'block', textDecoration: 'none',
              background: '#161D2C', border: '1px solid rgba(232,238,250,0.06)',
              borderRadius: 12, padding: 32, position: 'relative',
              gridColumn: p.status === 'active' ? 'span 2' : 'span 1',
              opacity: p.status === 'pipeline' ? 0.7 : 1,
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start' }}>
                <div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                    {p.status === 'active' ? (
                      <img src="../../assets/hardys/hardys-logo-transparent.png"
                        alt={p.name}
                        style={{ height: 48, width: 'auto', display: 'block' }}/>
                    ) : (
                      <span className="jv-wordmark" style={{ fontSize: 32 }}>
                        {p.name}
                      </span>
                    )}
                    {p.domain && (
                      <span style={{ fontFamily: 'var(--font-mono)', fontSize: 12, color: '#22D3EE' }}>
                        {p.domain}
                      </span>
                    )}
                  </div>
                  <div style={{ fontFamily: 'var(--font-sans)', fontSize: 12, color: '#94A3B8',
                    marginTop: 10, letterSpacing: '0.02em' }}>{p.namesake}</div>
                </div>
                <span style={{
                  fontFamily: 'var(--font-sans)', fontSize: 10, fontWeight: 600,
                  letterSpacing: '0.18em', textTransform: 'uppercase',
                  padding: '4px 10px', borderRadius: 999,
                  color: p.status === 'active' ? '#22D3EE' : '#94A3B8',
                  border: `1px solid ${p.status === 'active' ? 'rgba(34,211,238,0.35)' : 'rgba(148,163,184,0.3)'}`,
                }}>{p.status === 'active' ? '● Active' : 'Pipeline'}</span>
              </div>
              <p style={{ fontFamily: 'var(--font-sans)', fontSize: p.status === 'active' ? 18 : 15,
                lineHeight: 1.55, color: '#E8EEFA', margin: '24px 0 0', maxWidth: 560 }}>
                {p.desc}
              </p>
              {p.status === 'active' && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 28,
                  color: '#22D3EE', fontSize: 14, fontWeight: 600 }}>
                  Visit hardys.ai <span>→</span>
                </div>
              )}
            </a>
          ))}
        </div>
      </div>
    </section>
  );
};

const Philosophy = () => (
  <section id="research" style={{ padding: '120px 32px', maxWidth: 1200, margin: '0 auto' }}>
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 32 }}>
      {[
        { n: 'I', title: 'Complexity is the default', body: 'We build for senior professionals. We do not simplify — we make the intricate legible.' },
        { n: 'II', title: 'Story before data', body: 'Every feature begins with a scene or a claim. Numbers come after the claim, never as the lede.' },
        { n: 'III', title: 'Calm over spectacle', body: 'No hype, no buzzwords, no hollow AI rhetoric. A brilliant colleague explaining the future calmly.' },
      ].map((p) => (
        <div key={p.n}>
          <div style={{
            fontFamily: 'Georgia, serif', fontSize: 48, fontStyle: 'italic',
            background: 'linear-gradient(135deg,#A855F7,#22D3EE)',
            WebkitBackgroundClip: 'text', backgroundClip: 'text', color: 'transparent', lineHeight: 1,
          }}>{p.n}</div>
          <h3 style={{ fontFamily: 'var(--font-display)', fontSize: 26, lineHeight: 1.2,
            color: '#E8EEFA', margin: '16px 0 12px', fontWeight: 400 }}>{p.title}</h3>
          <p style={{ fontFamily: 'var(--font-sans)', fontSize: 15, lineHeight: 1.7, color: '#94A3B8', margin: 0 }}>
            {p.body}
          </p>
        </div>
      ))}
    </div>
  </section>
);

const Contact = () => (
  <section id="contact" style={{ position: 'relative', padding: '140px 32px', textAlign: 'center', overflow: 'hidden' }}>
    <DotField height={520}/>
    <div style={{ position: 'relative', maxWidth: 720, margin: '0 auto' }}>
      <SymbolMark size={80}/>
      <h2 style={{
        fontFamily: 'var(--font-display)', fontSize: 56, lineHeight: 1.08, letterSpacing: '-0.01em',
        margin: '28px 0 0', color: '#E8EEFA', fontWeight: 400,
      }}>
        <GradientText>Talk to us.</GradientText>
      </h2>
      <p style={{ fontFamily: 'var(--font-sans)', fontSize: 18, lineHeight: 1.7, color: '#94A3B8',
        margin: '20px 0 36px' }}>
        We read every message. Investors, faculty, institutions — all welcome.
      </p>
      <div style={{ display: 'flex', gap: 14, justifyContent: 'center' }}>
        <Button href="mailto:hello@juvant.io">hello@juvant.io</Button>
        <Button variant="ghost" href="#">Download the paper</Button>
      </div>
    </div>
  </section>
);

const Footer = () => (
  <footer style={{ padding: '48px 32px', borderTop: '1px solid rgba(232,238,250,0.06)' }}>
    <div style={{ maxWidth: 1200, margin: '0 auto', display: 'flex',
      justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 24, color: '#94A3B8', fontSize: 13 }}>
        <img src="../../assets/juvant/juvant-logo-dark.png" alt="Juvant" style={{ height: 24 }}/>
        <span>· Milan, Italy ·</span>
        <a href="mailto:hello@juvant.io" style={{ color: '#A855F7', textDecoration: 'none' }}>hello@juvant.io</a>
      </div>
      <div style={{ color: '#475569', fontSize: 12 }}>
        © 2026 Juvant Srls · Confidential
      </div>
    </div>
  </footer>
);

Object.assign(window, { Thesis, Portfolio, Philosophy, Contact, Footer });
