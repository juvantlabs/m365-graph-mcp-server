/* global React */
const { useState } = React;

const Nav = () => {
  const [scrolled, setScrolled] = useState(false);
  React.useEffect(() => {
    const onScroll = () => setScrolled(window.scrollY > 20);
    window.addEventListener('scroll', onScroll);
    return () => window.removeEventListener('scroll', onScroll);
  }, []);
  return (
    <div style={{
      position: 'fixed', inset: '0 0 auto 0', zIndex: 50,
      background: scrolled ? 'rgba(8,11,20,0.82)' : 'transparent',
      backdropFilter: scrolled ? 'blur(12px)' : 'none',
      borderBottom: scrolled ? '1px solid rgba(232,238,250,0.06)' : '1px solid transparent',
      transition: 'background 200ms ease, border 200ms ease',
    }}>
      {scrolled && <div style={{ position: 'absolute', inset: 'auto 0 -1px 0', height: 1,
        background: 'linear-gradient(90deg, rgba(168,85,247,0.4), rgba(34,211,238,0.4))' }}/>}
      <nav style={{
        maxWidth: 1280, margin: '0 auto', padding: '18px 32px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
      }}>
        <a href="#" style={{ display: 'flex', alignItems: 'center', gap: 8, textDecoration: 'none' }}>
          <img src="../../assets/juvant/juvant-logo-dark.png" alt="Juvant" style={{ height: 32 }}/>
        </a>
        <div style={{ display: 'flex', gap: 32, alignItems: 'center' }}>
          {['Thesis', 'Portfolio', 'Research', 'Contact'].map((item) => (
            <a key={item} href={`#${item.toLowerCase()}`} style={{
              color: '#E8EEFA', textDecoration: 'none', fontFamily: 'var(--font-sans)',
              fontSize: 14, opacity: 0.82,
            }}>{item}</a>
          ))}
          <Button variant="outline" href="#contact">Talk to us</Button>
        </div>
      </nav>
    </div>
  );
};

const Hero = () => (
  <section style={{ position: 'relative', minHeight: 760, overflow: 'hidden' }}>
    <DotField height={760}/>
    <div style={{ position: 'relative', maxWidth: 1080, margin: '0 auto',
      padding: '200px 32px 140px', textAlign: 'left' }}>
      <Eyebrow>JUVANT SRLS · MILAN</Eyebrow>
      <h1 style={{
        fontFamily: 'var(--font-display)', fontSize: 'clamp(52px, 7.2vw, 96px)',
        lineHeight: 1.02, letterSpacing: '-0.015em', margin: '22px 0 0', color: '#E8EEFA',
        fontWeight: 400, maxWidth: 900,
      }}>
        Augmenting human<br/><GradientText>ingenuity.</GradientText>
      </h1>
      <p style={{
        fontFamily: 'var(--font-sans)', fontSize: 20, lineHeight: 1.65,
        color: '#94A3B8', margin: '28px 0 0', maxWidth: 640,
      }}>
        We build AI agents that see what no one else sees — and make space for the people around them.
        Intelligence already exists. Our agents help it emerge.
      </p>
      <div style={{ display: 'flex', gap: 14, marginTop: 44 }}>
        <Button href="#portfolio">See the portfolio →</Button>
        <Button variant="ghost" href="#thesis">Read the thesis</Button>
      </div>
    </div>
    <GradientHairline opacity={0.25} style={{ position: 'absolute', insetInline: 0, bottom: 0 }}/>
  </section>
);

Object.assign(window, { Nav, Hero });
