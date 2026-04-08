const PptxGenJS = require("pptxgenjs");

// ── Themes ──────────────────────────────────────────────────────────────────
const THEMES = {
  midnight:     { dark:"0D1B2A", light:"F4F6F9", accent:"2563EB", accent2:"22D3EE", text:"1E293B", card:"162032" },
  terracotta:   { dark:"B85042", light:"F2EAE1", accent:"B85042", accent2:"7A9E7E", text:"2B2118", card:"E8D9CC" },
  forest:       { dark:"1A3C2B", light:"F0F5F1", accent:"2D6A4F", accent2:"95D5B2", text:"1A3C2B", card:"D8EAE0" },
  slate:        { dark:"1E293B", light:"F8FAFC", accent:"6366F1", accent2:"A5B4FC", text:"1E293B", card:"E2E8F0" },
  rose:         { dark:"881337", light:"FFF1F2", accent:"E11D48", accent2:"FB7185", text:"1F0A10", card:"FFE4E6" },
  charcoal:     { dark:"212121", light:"FAFAFA", accent:"212121", accent2:"9E9E9E", text:"212121", card:"F5F5F5" },
};

function getTheme(style, brandOn, brandColors) {
  if (brandOn && brandColors) {
    return {
      dark:    brandColors.primary.replace("#",""),
      light:   "FFFFFF",
      accent:  brandColors.accent.replace("#",""),
      accent2: brandColors.secondary ? brandColors.secondary.replace("#","") : brandColors.accent.replace("#",""),
      text:    brandColors.text.replace("#",""),
      card:    "F1F5F9",
    };
  }
  return THEMES[style] || THEMES.midnight;
}

// ── Colour helpers ───────────────────────────────────────────────────────────
function toHex(c) {
  if (!c || c === "transparent") return null;
  c = String(c).trim().replace("#","");
  if (/^[0-9A-Fa-f]{6}$/.test(c)) return c.toUpperCase();
  const m = ("rgb("+c).match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) < 0.06) return null;
    return [m[1],m[2],m[3]].map(v=>("0"+parseInt(v).toString(16)).slice(-2)).join("").toUpperCase();
  }
  return null;
}

function alpha(hex, pct) {
  // Returns hex + 2-char alpha (00-FF), pct = 0-100 opacity
  const a = Math.round(pct * 2.55).toString(16).padStart(2,"0").toUpperCase();
  return hex + a;
}

// ── Image helpers ────────────────────────────────────────────────────────────
async function fetchB64(url) {
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const buf = await r.arrayBuffer();
    const ct = r.headers.get("content-type") || "image/jpeg";
    return ct + ";base64," + Buffer.from(buf).toString("base64");
  } catch(e) { return null; }
}

async function prepImg(src) {
  if (!src) return null;
  if (src.startsWith("data:")) return src.slice(5);
  if (src.includes(";base64,")) return src;
  return await fetchB64(src);
}

async function getImgDims(d) {
  try {
    const b64 = d.includes(";base64,") ? d.split(";base64,")[1] : d;
    const buf = Buffer.from(b64,"base64");
    if (buf[0]===0x89&&buf[1]===0x50&&buf[2]===0x4E&&buf[3]===0x47)
      return { w:buf.readUInt32BE(16), h:buf.readUInt32BE(20) };
    let i=2;
    while(i<buf.length-8){
      if(buf[i]===0xFF&&(buf[i+1]===0xC0||buf[i+1]===0xC2))
        return { w:buf.readUInt16BE(i+7), h:buf.readUInt16BE(i+5) };
      const seg=buf.readUInt16BE(i+2);
      if(seg<2)break;
      i+=seg+2;
    }
  } catch(e) {}
  return null;
}

async function addLogo(s, pres, data, pos, withBg, isCover) {
  if (!data) return;
  const isBot = pos === "bottom-left";
  const maxW = isCover ? 1.6 : 0.72, maxH = isCover ? 0.7 : 0.3;
  let fw = maxW, fh = maxH;
  const dims = await getImgDims(data);
  if (dims && dims.w > 0) {
    const r = dims.w / dims.h;
    if (r > maxW/maxH) { fw=maxW; fh=maxW/r; } else { fh=maxH; fw=maxH*r; }
  }
  const x=0.22, y=isBot?(5.625-fh-0.15):0.13;
  if (withBg) {
    const bh=Math.max(fh+0.15,isCover?0.9:0.48);
    s.addShape(pres.shapes.RECTANGLE,{x:0,y:isBot?5.625-bh:0,w:10,h:bh,fill:{color:"FFFFFF"},line:{type:"none"}});
  }
  const imgData = data.startsWith("data:") ? data.slice(5) : data;
  try { s.addImage({data:imgData,x,y,w:fw,h:fh}); } catch(e) {}
}

// ── Design system ────────────────────────────────────────────────────────────
// All measurements in inches (10 x 5.625 slide)
const DS = {
  W: 10, H: 5.625,
  ml: 0.55,           // left margin
  mr: 0.55,           // right margin
  get cw() { return this.W - this.ml - this.mr; }, // content width = 8.9"
  ty: 0.36,           // title y
  sy: 0.88,           // subtitle y
  cy: 1.18,           // content start y
  bh: 0.62,           // bullet row height
  // font sizes (pt)
  fCover:  44,
  fTitle:  26,
  fHead:   18,
  fBody:   13,
  fCap:    10,
  fStat:   56,
  fQuote:  24,
  fEye:     9,
};

const mk = () => ({ type:"outer", color:"000000", blur:8, offset:2, angle:135, opacity:0.12 });

// ── PPTX slide builder ───────────────────────────────────────────────────────
async function buildSlide(s, pres, sd, theme, logoImg, logoPos, logoWb, isFirst) {
  const lay    = sd.layout || "default";
  const isDark = !!sd.dark;
  const BG     = isDark ? theme.dark  : theme.light;
  const TC     = isDark ? "FFFFFF"    : theme.text;
  const BC     = isDark ? "E2E8F0"    : "475569";
  const AC     = theme.accent;
  const AC2    = theme.accent2;

  s.background = { color: BG };
  await addLogo(s, pres, logoImg, logoPos||"top-left", logoWb, isFirst);

  // ── Helpers ────────────────────────────────────────────────────────────
  const t = (text, x, y, w, h, o={}) => {
    if (!text || !String(text).trim()) return;
    s.addText(String(text), {
      x, y, w, h,
      fontSize:  o.sz  || DS.fBody,
      fontFace:  o.ff  || "Calibri",
      bold:      !!o.bold,
      italic:    !!o.it,
      color:     o.c   || TC,
      align:     o.al  || "left",
      valign:    o.va  || "top",
      lineSpacingMultiple: o.lh || 1.3,
      charSpacing: o.cs || 0,
      wrap: true,
      margin: [0,0,0,0],
    });
  };

  const r = (x, y, w, h, color, transp=0, opts={}) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h,
      fill:  { color: toHex(color)||"CCCCCC", transparency: transp },
      line:  opts.line || { type:"none" },
      shadow: opts.shadow,
      rectRadius: opts.radius,
    });
  };

  const oval = (x, y, w, h, color, transp=0) => {
    s.addShape(pres.shapes.OVAL, {
      x, y, w, h,
      fill: { color: toHex(color)||"CCCCCC", transparency: transp },
      line: { type:"none" },
    });
  };

  const img = async (src, x, y, w, h) => {
    if (!src) return;
    const d = await prepImg(src);
    if (d) { try { s.addImage({data:d,x,y,w,h,sizing:{type:"cover",w,h}}); } catch(e){} }
  };

  // bullets helper — each bullet on its own text box for reliable layout
  const bullets = (items, x, y, w, gap=DS.bh, sz=DS.fBody, col) => {
    (items||[]).forEach((b,i) => {
      if (!b) return;
      s.addText([{text:"• "+b}], {
        x, y: y+i*gap, w, h: gap,
        fontSize: sz, fontFace:"Calibri",
        color: col||BC, valign:"middle", wrap:true, margin:[0,0,0,0],
      });
    });
  };

  const { ml, cw, ty, sy, cy, bh, fCover, fTitle, fHead, fBody, fCap, fStat, fQuote } = DS;

  // ── Layouts ────────────────────────────────────────────────────────────

  // ════ COVERS ════════════════════════════════════════════════════════════

  if (lay === "cover-center") {
    // Full-bleed image, dark gradient, headline + sub overlaid
    await img(sd.images?.[0], 0, 0, 10, 5.625);
    r(0, 3.0, 10, 2.625, "000000", 25);  // dark gradient bottom
    t(sd.heading, ml, 3.1, cw, 1.4, {sz:fCover, bold:true, c:"FFFFFF", al:"center", lh:1.08});
    if (sd.subheading) t(sd.subheading, ml, 4.3, cw, 0.6, {sz:fBody, c:"E2E8F0", al:"center"});

  } else if (lay === "cover-split") {
    // Image left half, theme panel right
    await img(sd.images?.[0], 0, 0, 5, 5.625);
    r(5, 0, 5, 5.625, isDark ? theme.dark : theme.light);
    t(sd.eyebrow||"", 5.3, 1.2, 4.3, 0.35, {sz:DS.fEye, bold:true, c:AC, cs:3});
    t(sd.heading, 5.3, 1.6, 4.2, 1.8, {sz:40, bold:true, c:TC, lh:1.1});
    if (sd.subheading) t(sd.subheading, 5.3, 3.6, 4.2, 0.8, {sz:fBody, c:BC, lh:1.45});

  } else if (lay === "cover-dark") {
    // Pure dark, giant centred text, subtle circle
    oval(3.0, -0.5, 4, 4, AC, 88);
    t(sd.eyebrow||"", ml, 1.3, cw, 0.35, {sz:DS.fEye, bold:true, c:AC, al:"center", cs:4});
    t(sd.heading, ml, 1.7, cw, 2.0, {sz:fCover, bold:true, c:"FFFFFF", al:"center", lh:1.08});
    if (sd.subheading) t(sd.subheading, 1.0, 3.9, 8.0, 0.7, {sz:fBody+1, c:"CBD5E1", al:"center"});

  // ════ TEXT ═══════════════════════════════════════════════════════════════

  } else if (lay === "default") {
    t(sd.heading, ml, ty, cw, 0.7, {sz:fTitle, bold:true, c:TC});
    bullets(sd.bullets, ml+0.1, cy, cw-0.1, bh, fBody, BC);

  } else if (lay === "title-body") {
    t(sd.heading, ml, ty, cw, 0.7, {sz:fTitle+2, bold:true, c:TC});
    if (sd.paragraph) t(sd.paragraph, ml, 1.0, cw, 1.4, {sz:fBody+1, c:BC, lh:1.6});
    bullets(sd.bullets, ml+0.1, 2.6, cw-0.1, bh, fBody, BC);

  } else if (lay === "two-col") {
    t(sd.heading, ml, ty, cw, 0.7, {sz:fTitle, bold:true, c:TC});
    const half = Math.ceil((sd.bullets||[]).length/2);
    bullets((sd.bullets||[]).slice(0,half),   ml,      cy, cw/2-0.15, bh, fBody, BC);
    bullets((sd.bullets||[]).slice(half), ml+cw/2+0.15, cy, cw/2-0.15, bh, fBody, BC);

  } else if (lay === "quote") {
    t("\u201C", 0.3, 0.1, 1.5, 1.4, {sz:110, bold:true, c:AC, ff:"Georgia"});
    t(sd.quote||(sd.bullets||[])[0]||sd.heading, ml, 1.1, cw, 1.8,
      {sz:fQuote, it:true, c:TC, al:"center", lh:1.45, ff:"Georgia"});
    r(4.1, 3.4, 1.8, 0.04, AC);
    if (sd.author) t("— "+sd.author, ml, 3.55, cw, 0.4, {sz:fBody, bold:true, c:AC, al:"center"});

  } else if (lay === "big-statement") {
    t(sd.heading, ml, 0.9, cw, 1.8, {sz:fTitle+10, bold:true, c:TC, al:"center", lh:1.15, ff:"Georgia"});
    if (sd.paragraph) t(sd.paragraph, 1.0, 3.0, 8.0, 0.8, {sz:fBody+1, c:BC, al:"center", lh:1.5});

  } else if (lay === "agenda") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    (sd.bullets||[]).slice(0,4).forEach((b,i) => {
      const y = 1.1 + i * 1.1;
      oval(ml, y+0.05, 0.44, 0.44, AC, 0);
      t(String(i+1), ml, y+0.04, 0.44, 0.44, {sz:fBody, bold:true, c:"FFFFFF", al:"center", va:"middle"});
      t(b, ml+0.6, y+0.08, cw-0.6, 0.38, {sz:fBody+1, bold:true, c:TC});
    });

  } else if (lay === "closing") {
    await img(sd.images?.[0], 0, 0, 10, 5.625);
    r(0, 0, 10, 5.625, "000000", 45);
    t(sd.heading, ml, 1.6, cw, 1.6, {sz:fCover, bold:true, c:"FFFFFF", al:"center", lh:1.1});
    if (sd.subheading) t(sd.subheading, 1.0, 3.4, 8.0, 0.7, {sz:fBody+1, c:"CBD5E1", al:"center"});
    if (sd.cta) t(sd.cta, 3.0, 4.3, 4.0, 0.5, {sz:fBody, c:AC2, al:"center", bold:true});

  // ════ INFOGRAPHIC ════════════════════════════════════════════════════════

  } else if (lay === "stat") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC, al:"center"});
    r(2.0, 0.9, 6.0, 2.5, AC, 88, {shadow:mk()});
    r(2.0, 0.9, 6.0, 0.06, AC);
    t(sd.stat||"—", 2.0, 1.0, 6.0, 1.8, {sz:fStat, bold:true, c:AC, al:"center", lh:1.0});
    if ((sd.bullets||[]).length) {
      t((sd.bullets||[]).join("   ·   "), ml, 3.6, cw, 0.45, {sz:fBody, c:BC, al:"center"});
    }

  } else if (lay === "three-stats") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const cw3=2.73, gap3=0.255, cx3s=[ml, ml+cw3+gap3, ml+(cw3+gap3)*2];
    (sd.stats||[]).slice(0,3).forEach((st,i) => {
      const x=cx3s[i];
      r(x, 0.85, cw3, 3.6, AC, isDark?88:92, {shadow:mk()});
      r(x, 0.85, cw3, 0.06, AC);
      t(st.value||"—", x, 1.05, cw3, 1.3, {sz:fStat, bold:true, c:AC, al:"center", lh:1.0});
      t(st.label||"", x+0.1, 2.45, cw3-0.2, 0.5, {sz:fHead, bold:true, c:TC, al:"center"});
      if (st.sub) t(st.sub, x+0.1, 3.0, cw3-0.2, 0.55, {sz:fCap, c:BC, al:"center", lh:1.35});
    });

  } else if (lay === "timeline") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const items=sd.bullets||[], n=Math.max(items.length,1);
    const lx=0.7, rx=9.3, span=rx-lx, lY=2.6;
    r(lx, lY-0.02, span, 0.04, AC, 55);
    items.forEach((item,i) => {
      const cx = lx + i*(span/(n-1||1));
      oval(cx-0.22, lY-0.22, 0.44, 0.44, BG, 0);
      s.addShape(pres.shapes.OVAL,{x:cx-0.22,y:lY-0.22,w:0.44,h:0.44,
        fill:{type:"none"},line:{color:AC,width:2.5}});
      t(String(i+1), cx-0.22, lY-0.22, 0.44, 0.44, {sz:fCap, bold:true, c:AC, al:"center", va:"middle"});
      const above = i%2===0;
      const tw = Math.min(span/n*1.1, 1.8);
      t(item, cx-tw/2, above?lY-1.1:lY+0.4, tw, 0.9, {sz:fCap, c:BC, al:"center", lh:1.35});
    });

  } else if (lay === "process") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const items=sd.bullets||[], n=Math.max(items.length,1);
    const pw=(cw-(n-1)*0.18)/n;
    items.forEach((item,i) => {
      const x=ml+i*(pw+0.18);
      r(x, 0.82, pw, 3.8, AC, isDark?88:92, {shadow:mk()});
      r(x, 0.82, pw, 0.65, AC, isDark?40:20);
      t(String(i+1), x, 0.84, pw, 0.62, {sz:fHead+4, bold:true, c:isDark?"FFFFFF":AC, al:"center", lh:1.0});
      t(item, x+0.12, 1.6, pw-0.24, 2.8, {sz:fBody, c:BC, al:"center", lh:1.5});
      if(i<n-1) r(x+pw, 2.55, 0.18, 0.04, AC, 50);
    });

  } else if (lay === "comparison") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const pw=(cw-0.18)/2;
    const sides=[
      {items:(sd.left||[]), x:ml},
      {items:(sd.right||[]), x:ml+pw+0.18},
    ];
    sides.forEach(({items,x},si) => {
      r(x, 0.82, pw, 4.4, AC, isDark?88:92, {shadow:mk()});
      r(x, 0.82, pw, 0.52, AC, si===0?25:45);
      t(items[0]||"", x+0.12, 0.86, pw-0.24, 0.44, {sz:fHead, bold:true, c:TC, al:"center"});
      (items.slice(1)||[]).forEach((b,i) => {
        t("• "+b, x+0.14, 1.46+i*bh, pw-0.28, bh, {sz:fBody, c:BC, lh:1.4});
      });
    });

  } else if (lay === "four-icons") {
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const iw=cw/4, ir=0.52;
    (sd.bullets||[]).slice(0,4).forEach((b,i) => {
      const cx=ml+i*iw+iw/2;
      oval(cx-ir, 1.55, ir*2, ir*2, AC, 82);
      s.addShape(pres.shapes.OVAL,{x:cx-ir,y:1.55,w:ir*2,h:ir*2,
        fill:{type:"none"},line:{color:AC,width:1.5}});
      t(b, ml+i*iw, 2.75, iw-0.05, 1.6, {sz:fCap, bold:true, c:BC, al:"center", lh:1.4});
    });

  } else if (lay === "two-cols-cards") {
    // Two feature cards side by side — from deck1 design
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    const pw=(cw-0.22)/2;
    (sd.cards||[]).slice(0,2).forEach((card,i) => {
      const x=ml+i*(pw+0.22);
      r(x, 0.82, pw, 4.4, AC, isDark?88:92, {shadow:mk()});
      r(x, 0.82, pw, 0.06, AC);
      t(card.title||"", x+0.18, 1.0, pw-0.36, 0.55, {sz:fHead, bold:true, c:TC});
      t(card.body||"", x+0.18, 1.65, pw-0.36, 3.0, {sz:fBody, c:BC, lh:1.5});
    });

  } else if (lay === "three-cards") {
    // Three feature cards — exact deck1 pattern
    t(sd.heading, ml, ty, cw, 0.6, {sz:fTitle, bold:true, c:TC});
    if (sd.subheading) t(sd.subheading, ml, sy, cw, 0.38, {sz:fCap, c:BC});
    const cw3c=2.73, gap3c=0.255, cx3cs=[ml, ml+cw3c+gap3c, ml+(cw3c+gap3c)*2];
    (sd.cards||[]).slice(0,3).forEach((card,i) => {
      const x=cx3cs[i];
      r(x, 0.96, cw3c, 3.98, AC, isDark?88:92, {shadow:mk()});
      r(x, 0.96, cw3c, 0.06, AC);
      // Number circle
      oval(x+0.16, 1.1, 0.55, 0.55, AC, 0);
      t(String(i+1), x+0.16, 1.1, 0.55, 0.55, {sz:fBody, bold:true, c:"FFFFFF", al:"center", va:"middle"});
      t(card.title||"", x+0.14, 1.82, cw3c-0.28, 0.52, {sz:fHead, bold:true, c:TC});
      t(card.body||"", x+0.14, 2.44, cw3c-0.28, 2.2, {sz:fBody-1, c:BC, lh:1.45});
    });

  // ════ IMAGE ══════════════════════════════════════════════════════════════

  } else if (lay === "img-right") {
    t(sd.heading, ml, ty, cw/2-0.1, 0.7, {sz:fTitle, bold:true, c:TC});
    bullets(sd.bullets, ml, cy, cw/2-0.1, bh, fBody, BC);
    await img(sd.images?.[0], 5.0, 0, 5.0, 5.625);

  } else if (lay === "img-left") {
    await img(sd.images?.[0], 0, 0, 5.0, 5.625);
    t(sd.heading, 5.3, ty, cw/2-0.3, 0.7, {sz:fTitle, bold:true, c:TC});
    bullets(sd.bullets, 5.3, cy, cw/2-0.3, bh, fBody, BC);

  } else if (lay === "img-full") {
    await img(sd.images?.[0], 0, 0, 10, 5.625);
    r(0, 3.1, 10, 2.525, "000000", 25);
    t(sd.heading, ml, 3.2, cw, 1.1, {sz:fTitle+4, bold:true, c:"FFFFFF", lh:1.1});
    if ((sd.bullets||[])[0]) t(sd.bullets[0], ml, 4.5, cw, 0.5, {sz:fBody, c:"E2E8F0"});

  } else if (lay === "img-hero") {
    await img(sd.images?.[0], 0, 0, 5.6, 5.625);
    r(5.6, 0, 4.4, 5.625, isDark?theme.dark:theme.light);
    t(sd.eyebrow||"", 5.9, 0.9, 3.8, 0.35, {sz:DS.fEye, bold:true, c:AC, cs:3});
    t(sd.heading, 5.9, 1.3, 3.8, 1.4, {sz:fTitle+2, bold:true, c:TC, lh:1.1});
    bullets(sd.bullets, 5.9, 2.9, 3.7, bh*0.95, fBody, BC);

  } else if (lay === "two-images") {
    t(sd.heading, ml, 0.12, cw, 0.46, {sz:fHead+2, bold:true, c:TC});
    await img(sd.images?.[0], 0.1, 0.66, 4.8, 3.94);
    await img(sd.images?.[1], 5.1, 0.66, 4.8, 3.94);
    if ((sd.bullets||[])[0]) t(sd.bullets[0], 0.1, 4.66, 4.8, 0.44, {sz:fCap, bold:true, c:BC});
    if ((sd.bullets||[])[1]) t(sd.bullets[1], 5.1, 4.66, 4.8, 0.44, {sz:fCap, bold:true, c:BC});

  } else if (lay === "img-mosaic") {
    // 4 images fill the entire slide — no text
    await img(sd.images?.[0], 0,     0,     5.1, 3.25);
    await img(sd.images?.[1], 5.18,  0,     4.82, 1.56);
    await img(sd.images?.[2], 5.18,  1.64,  4.82, 1.61);
    await img(sd.images?.[3], 0,     3.33,  10,   2.295);

  } else {
    // Fallback — clean default
    t(sd.heading, ml, ty, cw, 0.7, {sz:fTitle, bold:true, c:TC});
    bullets(sd.bullets, ml+0.1, cy, cw-0.1, bh, fBody, BC);
  }
}

// ── Main PPTX builder ────────────────────────────────────────────────────────
async function buildPptx(deck, pres, logoImg, logoPos, logoWb) {
  pres.layout = "LAYOUT_16x9";
  pres.title  = deck.title || "Presentation";

  for (let i=0; i<deck.slides.length; i++) {
    const sd = deck.slides[i];
    const s  = pres.addSlide();
    await buildSlide(s, pres, sd, deck.theme, logoImg, logoPos, logoWb, i===0);
  }
}

// ── AI prompt ─────────────────────────────────────────────────────────────────
function buildPrompt(input, count, style) {
  return `You are a world-class presentation designer. Create a ${count}-slide deck.

USER INPUT: ${input}
STYLE: ${style}

SLIDE SCHEMA — return ONLY a raw JSON object with this exact structure, no markdown:
{
  "title": "Deck title",
  "slides": [
    {
      "layout": "cover-center|cover-split|cover-dark|default|title-body|two-col|quote|big-statement|agenda|closing|stat|three-stats|timeline|process|comparison|four-icons|two-cols-cards|three-cards|img-right|img-left|img-full|img-hero|two-images|img-mosaic",
      "dark": true,
      "heading": "MAX 8 words — sharp, specific, data-driven",
      "subheading": "MAX 10 words — cover/closing only, else empty string",
      "eyebrow": "MAX 4 words ALL CAPS — optional label above heading",
      "bullets": ["MAX 7 words each","MAX 3 bullets total"],
      "paragraph": "MAX 20 words — only on title-body/two-col, empty on all image/stat/cover/closing layouts",
      "quote": "MAX 15 words — only for quote layout",
      "author": "First Last, Title — only for quote layout",
      "stat": "single number with unit e.g. $4.2B or 89% or 3.2× — only for stat layout",
      "stats": [{"value":"41%","label":"Revenue Growth","sub":"vs 12% industry avg"}],
      "cards": [{"title":"Card heading","body":"2-3 sentence description"}],
      "left": ["header item","bullet 1","bullet 2"],
      "right": ["header item","bullet 1","bullet 2"],
      "cta": "Call to action text — only for closing",
      "images": ["3-5 word Pexels keyword for each image slot"],
      "imageKeyword": "3-5 words for main image",
      "speakerNote": "One sentence for presenter"
    }
  ]
}

STRICT RULES:
1. Slide 1: layout "cover-center" OR "cover-split", dark:true
2. Slide 2: use "stat" or "three-stats" if there are numbers to highlight
3. Last slide: layout "closing", dark:true
4. Use image layouts (img-right, img-left, img-full, img-hero) for at least 30% of content slides
5. Use "three-cards" or "two-cols-cards" for feature/benefit slides — NOT plain bullets
6. Use "timeline" or "process" for sequential steps
7. NEVER use the same layout twice in a row
8. "three-stats" requires stats array: [{"value":"...","label":"...","sub":"..."}] × 3
9. "three-cards" requires cards array: [{"title":"...","body":"..."}] × 3
10. "comparison" requires left array and right array, first item of each is the column header
11. images array: provide one keyword per image slot the layout needs
12. paragraph MUST be empty string "" on: cover, closing, stat, three-stats, quote, big-statement, img-*, two-images, img-mosaic, four-icons, timeline, process
13. bullets MUST be empty array [] on: stat, three-stats, three-cards, two-cols-cards, comparison, quote, img-mosaic`;
}

// ── Request handler ──────────────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST")    return res.status(405).json({error:"Method not allowed"});
  if (!req.body || typeof req.body !== "object")
    return res.status(400).json({error:"Request body missing or not JSON"});

  try {
    const { action, input, slideCount, style, deck,
            logoData, logoPos, logoWhiteBg, brandOn, brandColors } = req.body;

    const apiKey = process.env.ANTHROPIC_API_KEY;
    const model  = process.env.CLAUDE_MODEL || "claude-sonnet-4-20250514";

    // ── Generate deck outline ──────────────────────────────────────────────
    if (action === "generate") {
      const count = Math.max(4, Math.min(20, parseInt(slideCount)||8));
      const prompt = buildPrompt(input, count, style||"midnight");

      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01"},
        body: JSON.stringify({model, max_tokens:8000, messages:[{role:"user",content:prompt}]}),
      });
      const d = await r.json();
      if (!r.ok) return res.status(r.status).json({error: d.error?.message||"API error"});

      const text  = d.content.map(b=>b.text||"").join("");
      const clean = text.replace(/```json|```/g,"").trim();
      const s     = clean.indexOf("{"), e = clean.lastIndexOf("}");
      if (s===-1||e===-1) return res.status(500).json({error:"No JSON in response"});

      let result;
      try { result = JSON.parse(clean.slice(s,e+1)); }
      catch(pe) { return res.status(500).json({error:"Invalid JSON: "+pe.message}); }

      return res.status(200).json({deck: result});
    }

    // ── Export PPTX ────────────────────────────────────────────────────────
    if (action === "export") {
      if (!deck || !deck.slides) return res.status(400).json({error:"No deck data"});

      const theme = getTheme(deck.style||style, brandOn, brandColors);
      deck.theme  = theme;

      const pres = new PptxGenJS();

      let logoImg = null;
      if (logoData) {
        logoImg = logoData.startsWith("data:") || logoData.includes(";base64,")
          ? logoData : await fetchB64(logoData);
      }

      try { await buildPptx(deck, pres, logoImg, logoPos, logoWhiteBg); }
      catch(e) {
        console.error("buildPptx:", e);
        return res.status(500).json({error:"Build failed: "+e.message});
      }

      const buf  = await pres.write({outputType:"nodebuffer"});
      const safe = ((deck.title||"presentation").trim()||"presentation").replace(/[^a-z0-9]/gi,"_");

      res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition",`attachment; filename="${safe}.pptx"`);
      res.setHeader("Content-Length", buf.length);
      return res.status(200).send(buf);
    }

    return res.status(400).json({error:"Invalid action"});

  } catch(err) {
    console.error(err);
    return res.status(500).json({error: err.message});
  }
};
