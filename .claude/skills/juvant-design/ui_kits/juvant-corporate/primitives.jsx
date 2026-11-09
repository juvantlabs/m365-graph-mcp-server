/* global React */
const { useState } = React;

// ---------- Primitive building blocks ----------

const Eyebrow = ({ children, color = 'violet' }) => (
  <div className="jv-eyebrow" style={{ color: color === 'cyan' ? '#22D3EE' : '#A855F7' }}>
    {children}
  </div>
);

const GradientText = ({ as: Tag = 'span', children, style, ...rest }) => (
  <Tag className="jv-gradient-text" style={style} {...rest}>{children}</Tag>
);

const Button = ({ variant = 'primary', href, children, onClick }) => {
  const base = {
    fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: 14,
    padding: '12px 24px', borderRadius: 8, border: 0, cursor: 'pointer',
    textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 8,
    transition: 'filter 160ms ease, box-shadow 160ms ease, transform 160ms ease',
  };
  const styles = {
    primary: { ...base, background: 'linear-gradient(90deg,#A855F7,#22D3EE)', color: '#fff' },
    outline: {
      ...base, color: '#E8EEFA', padding: '11px 23px',
      backgroundImage: 'linear-gradient(#080B14,#080B14), linear-gradient(90deg,#A855F7,#22D3EE)',
      backgroundOrigin: 'border-box', backgroundClip: 'padding-box, border-box',
      border: '1px solid transparent',
    },
    ghost: { ...base, background: 'transparent', color: '#E8EEFA', border: '1px solid rgba(232,238,250,0.12)' },
  };
  const Tag = href ? 'a' : 'button';
  return <Tag href={href} onClick={onClick} style={styles[variant]}>{children}</Tag>;
};

const GradientHairline = ({ opacity = 0.4, style }) => (
  <div style={{
    height: 1,
    background: `linear-gradient(90deg, rgba(168,85,247,${opacity}), rgba(34,211,238,${opacity}))`,
    ...style,
  }}/>
);

// Ambient dot-field background — reusable hero backdrop.
const DotField = ({ height = 560 }) => {
  const dots = React.useMemo(() => {
    const arr = [];
    const seed = (i) => Math.abs(Math.sin(i * 9301 + 49297) * 233280 % 1);
    for (let i = 0; i < 90; i++) {
      const x = seed(i) * 100;
      const y = seed(i + 100) * 100;
      const r = seed(i + 200) * 1.4 + 0.6;
      const hue = seed(i + 300);
      const color = hue < 0.4 ? '#A855F7' : hue < 0.8 ? '#22D3EE' : '#E8EEFA';
      arr.push({ x, y, r, color, op: seed(i + 400) * 0.5 + 0.15 });
    }
    return arr;
  }, []);
  return (
    <svg width="100%" height={height} style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }}
         preserveAspectRatio="none" viewBox="0 0 100 100">
      <defs>
        <radialGradient id="dfg1" cx="30%" cy="40%" r="50%">
          <stop offset="0%" stopColor="#A855F7" stopOpacity="0.2"/>
          <stop offset="100%" stopColor="#A855F7" stopOpacity="0"/>
        </radialGradient>
        <radialGradient id="dfg2" cx="75%" cy="70%" r="55%">
          <stop offset="0%" stopColor="#22D3EE" stopOpacity="0.15"/>
          <stop offset="100%" stopColor="#22D3EE" stopOpacity="0"/>
        </radialGradient>
      </defs>
      <rect width="100" height="100" fill="url(#dfg1)"/>
      <rect width="100" height="100" fill="url(#dfg2)"/>
      {dots.map((d, i) => (
        <circle key={i} cx={d.x} cy={d.y} r={d.r * 0.15} fill={d.color} opacity={d.op}/>
      ))}
    </svg>
  );
};

const SymbolMark = ({ size = 80 }) => (
  // Real Juvant symbol — extracted transparent crop (no dark card background).
  <img
    src="../../assets/juvant/juvant-symbol-transparent.png?v=3"
    alt="Juvant symbol"
    width={size * (210/158)}
    height={size}
    style={{ display: 'inline-block' }}
  />
);

Object.assign(window, { Eyebrow, GradientText, Button, GradientHairline, DotField, SymbolMark });
