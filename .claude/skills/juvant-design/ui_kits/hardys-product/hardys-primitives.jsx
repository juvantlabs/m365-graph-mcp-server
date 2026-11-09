/* global React */

const HdButton = ({ variant = 'primary', href, children, onClick, size = 'md', icon }) => {
  const pad = size === 'sm' ? '8px 14px' : size === 'lg' ? '14px 28px' : '12px 22px';
  const fz = size === 'sm' ? 13 : size === 'lg' ? 15 : 14;
  const base = {
    fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: fz,
    padding: pad, borderRadius: 8, border: 0, cursor: 'pointer',
    textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 8,
    transition: 'filter 160ms ease, box-shadow 160ms ease',
  };
  const styles = {
    primary: { ...base, background: 'linear-gradient(90deg,#A855F7,#22D3EE)', color: '#fff' },
    outline: { ...base, padding: size === 'sm' ? '7px 13px' : '11px 21px', color: '#E8EEFA',
      backgroundImage: 'linear-gradient(#080B14,#080B14), linear-gradient(90deg,#A855F7,#22D3EE)',
      backgroundOrigin: 'border-box', backgroundClip: 'padding-box, border-box',
      border: '1px solid transparent' },
    ghost: { ...base, background: 'transparent', color: '#E8EEFA',
      border: '1px solid rgba(232,238,250,0.12)' },
    soft: { ...base, background: 'rgba(168,85,247,0.12)', color: '#A855F7',
      border: '1px solid rgba(168,85,247,0.25)' },
  };
  const Tag = href ? 'a' : 'button';
  return <Tag href={href} onClick={onClick} style={styles[variant]}>{icon}{children}</Tag>;
};

const HdEyebrow = ({ children, color = '#A855F7' }) => (
  <div style={{
    fontFamily: 'var(--font-sans)', fontSize: 11, fontWeight: 600,
    textTransform: 'uppercase', letterSpacing: '0.22em', color,
  }}>{children}</div>
);

const HdPill = ({ children, tone = 'violet' }) => {
  const colors = {
    violet: { c: '#A855F7', b: 'rgba(168,85,247,0.35)' },
    cyan:   { c: '#22D3EE', b: 'rgba(34,211,238,0.35)' },
    slate:  { c: '#94A3B8', b: 'rgba(148,163,184,0.3)' },
  }[tone];
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 6,
      padding: '3px 10px', fontSize: 10, fontWeight: 600,
      letterSpacing: '0.16em', textTransform: 'uppercase',
      color: colors.c, border: `1px solid ${colors.b}`, borderRadius: 999,
    }}>{children}</span>
  );
};

const HdGradText = ({ children }) => (
  <span style={{
    background: 'linear-gradient(90deg,#A855F7,#22D3EE)',
    WebkitBackgroundClip: 'text', backgroundClip: 'text', color: 'transparent',
  }}>{children}</span>
);

const HdHair = ({ opacity = 0.35, style }) => (
  <div style={{ height: 1, background:
    `linear-gradient(90deg, rgba(168,85,247,${opacity}), rgba(34,211,238,${opacity}))`, ...style }}/>
);

// Orbit symbol — inline version of the Hardys mark
const HardysMark = ({ size = 56, animate = false }) => (
  <svg width={size} height={size} viewBox="0 0 80 80" fill="none"
    style={{ display: 'block' }}>
    <defs>
      <linearGradient id={`hm-g-${size}`} x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stopColor="#A855F7"/>
        <stop offset="100%" stopColor="#22D3EE"/>
      </linearGradient>
      <filter id={`hm-glow-${size}`}><feGaussianBlur stdDeviation="1.6"/></filter>
    </defs>
    <circle cx="40" cy="40" r="22" stroke="url(#hm-g-${size})" strokeWidth="1"
      fill="none" opacity="0.45" strokeDasharray="2 3"/>
    <circle cx="40" cy="40" r="30" stroke={`url(#hm-g-${size})`} strokeWidth="0.8"
      fill="none" opacity="0.35"/>
    <g filter={`url(#hm-glow-${size})`}>
      <circle cx="40" cy="40" r="8" fill="#22D3EE" opacity="0.3"/>
    </g>
    <circle cx="40" cy="40" r="4.5" fill="#A855F7"/>
    <circle cx="40" cy="40" r="2" fill="#ffffff"/>
    {/* Orbiting node with comet trail */}
    <g style={animate ? { transformOrigin: '40px 40px', animation: 'hdOrbit 6s linear infinite' } : {}}>
      <circle cx="70" cy="40" r="3" fill="#22D3EE" filter={`url(#hm-glow-${size})`}/>
      <circle cx="70" cy="40" r="1.6" fill="#ffffff"/>
      <circle cx="66" cy="40" r="0.8" fill="#22D3EE" opacity="0.7"/>
      <circle cx="62" cy="40" r="0.6" fill="#22D3EE" opacity="0.4"/>
      <circle cx="58" cy="40" r="0.5" fill="#22D3EE" opacity="0.2"/>
    </g>
    {animate && <style>{`@keyframes hdOrbit { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>}
  </svg>
);

Object.assign(window, { HdButton, HdEyebrow, HdPill, HdGradText, HdHair, HardysMark });
