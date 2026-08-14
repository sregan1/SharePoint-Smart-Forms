'use strict';
// Generates docs/screenshots/*.png for the README and USER-GUIDE — pure HTML/CSS
// mockups rendered with puppeteer-core, matching the real designer/filler/results
// UI (see components under src/webparts/smartForms/components/). No live
// SharePoint connection, no gulp serve.
// Run: node docs/generate-screenshots.js

const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME = fs.existsSync('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
  ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  : 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

const F = `'Segoe UI', Arial, sans-serif`;

// ── Theme tokens (light fallback, matches src/webparts/smartForms/utils/theme.ts) ──
const T = {
  bg: '#f3f2f1',
  card: '#ffffff',
  surface: '#faf9f8',
  surface2: '#f3f2f1',
  fg: '#323130',
  fgMuted: '#605e5c',
  fgSubtle: '#8a8886',
  border: '#edebe9',
  borderStrong: '#d2d0ce',
  accent: '#0078d4',
  accentHover: '#106ebe',
  accentTint: '#eff6fc',
  accentTintStrong: '#deecf9',
  error: '#a4262c',
  errorTint: '#fde7e9',
  success: '#0b6a0b',
  successTint: '#dff6dd',
  warning: '#8a5700',
  warningTint: '#fff4ce',
  shadowSm: '0 1px 2px rgba(0,0,0,0.08)',
  shadowMd: '0 2px 8px rgba(0,0,0,0.10)',
  shadowLg: '0 4px 16px rgba(0,0,0,0.14)',
};

const CATEGORICAL = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7'];

// ── Shared primitives ────────────────────────────────────────────────────────
function icon(pathD, opts = {}) {
  const { size = 16, color = T.fgMuted, viewBox = '0 0 24 24' } = opts;
  return `<svg width="${size}" height="${size}" viewBox="${viewBox}" fill="${color}" style="flex-shrink:0;"><path d="${pathD}"/></svg>`;
}
const ICONS = {
  grip: 'M9 4a1 1 0 110 2 1 1 0 010-2zm6 0a1 1 0 110 2 1 1 0 010-2zM9 9a1 1 0 110 2 1 1 0 010-2zm6 0a1 1 0 110 2 1 1 0 010-2zM9 14a1 1 0 110 2 1 1 0 010-2zm6 0a1 1 0 110 2 1 1 0 010-2zM9 19a1 1 0 110 2 1 1 0 010-2zm6 0a1 1 0 110 2 1 1 0 010-2z',
  bulb: 'M9 21c0 .55.45 1 1 1h4c.55 0 1-.45 1-1v-1H9v1zm3-19C8.14 2 5 5.14 5 9c0 2.38 1.19 4.47 3 5.74V17c0 .55.45 1 1 1h6c.55 0 1-.45 1-1v-2.26c1.81-1.27 3-3.36 3-5.74 0-3.86-3.14-7-7-7z',
  gear: 'M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.61l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96a7.02 7.02 0 0 0-1.62-.94l-.36-2.54A.484.484 0 0 0 14 2h-4a.484.484 0 0 0-.48.41l-.36 2.54a7.36 7.36 0 0 0-1.62.94l-2.39-.96a.48.48 0 0 0-.59.22L2.74 8.87a.47.47 0 0 0 .12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58a.49.49 0 0 0-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.36 1.04.67 1.62.94l.36 2.54c.05.24.27.41.49.41h4c.22 0 .44-.17.47-.41l.36-2.54a7.36 7.36 0 0 0 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32a.47.47 0 0 0-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z',
  eye: 'M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8a3 3 0 100 6 3 3 0 000-6z',
  send: 'M2 21l21-9L2 3v7l15 2-15 2v7z',
  check: 'M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z',
  refresh: 'M17.65 6.35A7.958 7.958 0 0012 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08a5.99 5.99 0 01-5.65 4c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L14 11h7V4l-3.35 2.35z',
  info: 'M11 7h2v2h-2V7zm0 4h2v6h-2v-6zm1-9C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z',
  flow: 'M4 4h6v6H4V4zm10 10h6v6h-6v-6zM10 7h4v2h5v6h-2v-4h-3v2H8V9h2V7z',
  warn: 'M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z',
  lock: 'M12 17a2 2 0 002-2 2 2 0 00-2-2 2 2 0 00-2 2 2 2 0 002 2zm6-9h-1V6a5 5 0 00-10 0v2H6a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V10a2 2 0 00-2-2zM8.9 6c0-1.71 1.39-3.1 3.1-3.1s3.1 1.39 3.1 3.1v2H8.9V6z',
  clock: 'M12 2C6.47 2 2 6.47 2 12s4.47 10 10 10 10-4.47 10-10S17.53 2 12 2zm.5 5H11v6l5.25 3.15.75-1.23-4.5-2.67V7z',
  print: 'M19 8H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zm-3 11H8v-5h8v5zm3-7c-.55 0-1-.45-1-1s.45-1 1-1 1 .45 1 1-.45 1-1 1zm-1-9H6v4h12V3z',
  chevronL: 'M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z',
  chevronR: 'M8.59 16.59L10 18l6-6-6-6-1.41 1.41L13.17 12z',
  contact: 'M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z',
  attach: 'M16.5 6v11.5c0 2.21-1.79 4-4 4s-4-1.79-4-4V5c0-1.38 1.12-2.5 2.5-2.5S13.5 3.62 13.5 5v10.5c0 .55-.45 1-1 1s-1-.45-1-1V6H10v9.5c0 1.38 1.12 2.5 2.5 2.5s2.5-1.12 2.5-2.5V5c0-2.21-1.79-4-4-4S7 2.79 7 5v12.5c0 3.04 2.46 5.5 5.5 5.5s5.5-2.46 5.5-5.5V6h-1.5z',
};

function chip(label, opts = {}) {
  const { bg = T.surface2, color = T.fgMuted, weight = 600, border } = opts;
  return `<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 9px;border-radius:10px;
    font-size:11px;font-weight:${weight};background:${bg};color:${color};font-family:${F};white-space:nowrap;
    ${border ? `border:1px solid ${border};` : ''}">${label}</span>`;
}

function btn(label, opts = {}) {
  const { primary = false, iconD = null, size = 13 } = opts;
  const bg = primary ? T.accent : T.card;
  const color = primary ? '#fff' : T.fg;
  const border = primary ? T.accent : T.borderStrong;
  return `<div style="display:inline-flex;align-items:center;gap:6px;padding:6px 14px;border-radius:2px;
    border:1px solid ${border};background:${bg};color:${color};font-size:${size}px;font-family:${F};
    font-weight:600;cursor:pointer;white-space:nowrap;">
    ${iconD ? icon(iconD, { size: 14, color }) : ''}${label}</div>`;
}

function iconBtn(iconD, opts = {}) {
  const { active = false, size = 30 } = opts;
  return `<div style="width:${size}px;height:${size}px;border-radius:2px;display:flex;align-items:center;
    justify-content:center;border:1px solid ${active ? T.accent : T.borderStrong};
    background:${active ? T.accentTint : T.card};cursor:pointer;flex-shrink:0;">
    ${icon(iconD, { size: 16, color: active ? T.accent : T.fgMuted })}</div>`;
}

function pageShell(bodyHtml, opts = {}) {
  const { width = 1280 } = opts;
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>*{margin:0;padding:0;box-sizing:border-box;}body{-webkit-font-smoothing:antialiased;}</style>
</head><body style="background:${T.bg};width:${width}px;font-family:${F};color:${T.fg};">
${bodyHtml}
</body></html>`;
}

function card(innerHtml, opts = {}) {
  const { padding = '0', radius = 8 } = opts;
  return `<div style="background:${T.card};border-radius:${radius}px;box-shadow:${T.shadowMd};
    border:1px solid ${T.border};overflow:hidden;padding:${padding};">${innerHtml}</div>`;
}

// ── Owner bar (shared across designer/filler/results mockups) ────────────────
function ownerBar(opts = {}) {
  const {
    activeTab = 'design', questionCount = 9, saveState = 'saved', shareLabel = 'Collect responses',
  } = opts;

  const saveChip = {
    saved: chip('&#10003; Saved', { bg: T.successTint, color: T.success }),
    saving: chip('&#8635; Saving&hellip;', { bg: T.warningTint, color: T.warning }),
  }[saveState];

  function pill(label, active) {
    return `<div style="padding:6px 16px;border-radius:16px;font-size:13px;font-family:${F};font-weight:600;
      cursor:pointer;white-space:nowrap;
      ${active ? `background:${T.card};color:${T.accent};box-shadow:${T.shadowSm};` : `color:${T.fgMuted};`}">${label}</div>`;
  }

  return `
  <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 18px;
    border-bottom:1px solid ${T.border};background:${T.surface};gap:14px;flex-wrap:wrap;">
    <div style="display:flex;align-items:center;gap:2px;background:${T.surface2};border-radius:18px;padding:3px;">
      ${pill(`Questions (${questionCount})`, activeTab === 'design')}
      ${pill('Responses', activeTab === 'responses')}
    </div>
    <div style="display:flex;align-items:center;gap:8px;">
      ${saveChip}
      ${iconBtn(ICONS.bulb, { size: 30 })}
      ${iconBtn(ICONS.gear, { size: 30 })}
      ${btn('Preview', { iconD: ICONS.eye })}
      ${btn(shareLabel, { primary: true, iconD: ICONS.send })}
    </div>
  </div>`;
}

function previewBanner() {
  return `<div style="display:flex;align-items:center;gap:10px;padding:9px 18px;background:${T.accentTint};
    border-bottom:1px solid ${T.accentTintStrong};font-size:12.5px;color:${T.fg};font-family:${F};">
    ${icon(ICONS.eye, { size: 15, color: T.accent })}
    <span style="flex:1;">Preview &mdash; nothing you submit here is saved. Validation and branching behave exactly as they will for respondents.</span>
    ${btn('Close preview')}
  </div>`;
}

// ── 01: Form designer canvas ──────────────────────────────────────────────────
function designerCanvasPage() {
  function questionCard({ num, title, type, required = true, branching = false, live = false, expanded = false, bodyHtml = '' }) {
    const badges = [
      required ? `<span style="color:${T.error};font-weight:700;">*</span>` : '',
      branching ? chip('&#128257; Branching', { bg: T.accentTint, color: T.accent }) : '',
      live ? chip('&#128274; Live', { bg: T.surface2, color: T.fgMuted }) : '',
    ].filter(Boolean).join(' ');

    return `<div style="display:flex;gap:10px;background:${T.card};border:1px solid ${expanded ? T.accent : T.border};
      border-radius:6px;padding:14px 16px;margin-bottom:10px;${expanded ? `box-shadow:${T.shadowMd};` : ''}">
      <div style="padding-top:2px;">${icon(ICONS.grip, { size: 16, color: T.fgSubtle })}</div>
      <div style="flex:1;min-width:0;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:${expanded ? '10px' : '6px'};">
          <span style="font-size:13px;color:${T.fgSubtle};font-weight:600;">${num}.</span>
          <span style="font-size:14.5px;font-weight:600;color:${T.fg};">${title}</span>
          ${badges}
          <span style="margin-left:auto;font-size:11px;color:${T.fgSubtle};">${type}</span>
          <div style="width:26px;height:26px;border-radius:3px;display:flex;align-items:center;justify-content:center;
            color:${T.fgMuted};font-size:16px;cursor:pointer;">&#8942;</div>
        </div>
        ${bodyHtml}
        ${expanded ? `<div style="display:flex;align-items:center;gap:14px;margin-top:12px;padding-top:10px;
          border-top:1px solid ${T.border};font-size:12px;color:${T.fgMuted};">
          <span>Width: <strong style="color:${T.fg};">Full</strong></span>
          <span style="display:flex;align-items:center;gap:5px;">
            <div style="width:30px;height:16px;border-radius:8px;background:${T.accent};position:relative;">
              <div style="width:12px;height:12px;border-radius:50%;background:#fff;position:absolute;top:2px;right:2px;"></div>
            </div>Required</span>
          <span style="margin-left:auto;color:${T.accent};font-weight:600;cursor:pointer;">Branching, validation &amp; more</span>
        </div>` : ''}
      </div>
    </div>`;
  }

  function choicePreview(options, disabled = true) {
    return `<div style="display:flex;flex-direction:column;gap:6px;margin-top:2px;">
      ${options.map(o => `<label style="display:flex;align-items:center;gap:8px;font-size:13px;color:${disabled ? T.fgMuted : T.fg};">
        <div style="width:16px;height:16px;border-radius:50%;border:1.5px solid ${T.borderStrong};flex-shrink:0;"></div>${o}
      </label>`).join('')}
    </div>`;
  }

  const outlineRows = [
    { n: 1, t: 'Full name', flag: false },
    { n: 2, t: 'Work email address', flag: false },
    { n: 3, t: 'Which department are you in?', flag: false },
    { n: 4, t: 'How satisfied are you with IT support?', branch: true },
    { n: 5, t: 'What could we improve?', branch: true },
    { n: 6, t: 'Would you recommend us to a colleague?', flag: false },
    { n: 7, t: 'Upload a screenshot (optional)', flag: false },
    { n: '&mdash;', t: 'Thanks for your time!', content: true },
  ];

  const body = `
  <div style="padding:18px 22px;display:flex;gap:18px;align-items:flex-start;">
    <div style="flex:1;min-width:0;">
      <div style="background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:16px 18px;margin-bottom:16px;">
        <div style="font-size:22px;font-weight:700;color:${T.fg};margin-bottom:6px;">IT Support Feedback</div>
        <div style="font-size:13px;color:${T.fgMuted};">Tell us about your recent support experience &mdash; it takes under two minutes.</div>
      </div>

      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
        ${chip('Section 1 of 2', { bg: T.surface2, color: T.fgMuted })}
        <span style="font-size:13px;font-weight:600;color:${T.fg};">About you</span>
      </div>

      ${questionCard({ num: 1, title: 'Full name', type: 'Short text', bodyHtml: `<div style="height:32px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};margin-top:2px;"></div>` })}
      ${questionCard({ num: 2, title: 'Work email address', type: 'Email', bodyHtml: `<div style="height:32px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};margin-top:2px;"></div>` })}
      ${questionCard({ num: 3, title: 'Which department are you in?', type: 'Choice', required: false, bodyHtml: choicePreview(['Sales', 'Engineering', 'Marketing', 'Operations']) })}

      <div style="display:flex;align-items:center;gap:8px;margin:18px 0 10px;">
        ${chip('Section 2 of 2', { bg: T.surface2, color: T.fgMuted })}
        <span style="font-size:13px;font-weight:600;color:${T.fg};">Your feedback</span>
      </div>

      ${questionCard({
        num: 4, title: 'How satisfied are you with IT support?', type: 'Opinion scale', branching: true, expanded: true,
        bodyHtml: `<div style="display:flex;align-items:center;gap:6px;margin-top:2px;">
          ${[0,1,2,3,4,5,6,7,8,9,10].map(n => `<div style="width:24px;height:24px;border-radius:4px;border:1px solid ${T.borderStrong};
            display:flex;align-items:center;justify-content:center;font-size:11px;color:${T.fgMuted};
            ${n === 9 ? `background:${T.accent};color:#fff;border-color:${T.accent};` : ''}">${n}</div>`).join('')}
        </div>
        <div style="display:flex;justify-content:space-between;font-size:11px;color:${T.fgSubtle};margin-top:4px;">
          <span>Not at all likely</span><span>Extremely likely</span>
        </div>`
      })}
      ${questionCard({ num: 5, title: 'What could we improve?', type: 'Long text', required: false, branching: true, bodyHtml: `<div style="height:56px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};margin-top:2px;"></div>` })}
      ${questionCard({ num: 6, title: 'Would you recommend us to a colleague?', type: 'Yes / No', live: true, bodyHtml: `<div style="display:flex;gap:8px;margin-top:2px;">
        ${btn('Yes')}${btn('No')}</div>` })}
      ${questionCard({ num: 7, title: 'Upload a screenshot (optional)', type: 'File upload', required: false, bodyHtml: `<div style="border:1.5px dashed ${T.borderStrong};border-radius:4px;padding:14px;text-align:center;font-size:12px;color:${T.fgSubtle};margin-top:2px;">Drag a file here or browse</div>` })}

      <div style="display:flex;justify-content:center;margin:16px 0;">
        ${btn('+ Add new')}
      </div>
      <div style="display:flex;justify-content:center;">
        <span style="font-size:12.5px;color:${T.accent};font-weight:600;cursor:pointer;">+ Add section</span>
      </div>
    </div>

    <!-- Outline rail -->
    <div style="width:260px;flex-shrink:0;background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:12px;">
      <div style="display:flex;align-items:center;gap:6px;font-size:12.5px;font-weight:700;color:${T.fg};margin-bottom:8px;">
        8 questions ${icon(ICONS.warn, { size:13, color: T.warning })}
      </div>
      <div style="height:28px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};margin-bottom:10px;
        display:flex;align-items:center;padding:0 8px;font-size:11.5px;color:${T.fgSubtle};">Find a question&hellip;</div>
      ${outlineRows.map(r => `<div style="display:flex;align-items:center;gap:6px;padding:6px 6px;border-radius:3px;font-size:12px;
        ${r.n === 4 ? `background:${T.accentTint};` : ''}">
        <span style="width:16px;color:${T.fgSubtle};text-align:right;flex-shrink:0;">${r.n}</span>
        <span style="flex:1;color:${T.fg};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${r.t}</span>
        ${r.branch ? icon(ICONS.flow, { size: 12, color: T.accent }) : ''}
      </div>`).join('')}
    </div>
  </div>`;

  return pageShell(`<div style="padding:18px;">${card(ownerBar({ questionCount: 8 }) + body)}</div>`, { width: 1300 });
}

// ── 02: Field editor panel — Branching tab ────────────────────────────────────
function fieldEditorPage() {
  const canvasBg = designerCanvasPage();
  // Reuse the canvas markup dimmed behind the panel by re-rendering just the body shell (not the outer <html>).
  const dimmedCanvas = `<div style="filter:blur(1px) brightness(0.97);pointer-events:none;">${canvasBg.replace(/^[\s\S]*?<body[^>]*>/,'').replace(/<\/body>[\s\S]*$/,'')}</div>`;

  function conditionRow(first, field, op, value) {
    return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
      <span style="font-size:11.5px;color:${T.fgMuted};width:36px;flex-shrink:0;">${first ? 'When' : 'And'}</span>
      <div style="flex:1;min-width:0;height:30px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.card};
        display:flex;align-items:center;padding:0 8px;font-size:12.5px;color:${T.fg};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${field}</div>
      <div style="width:90px;flex-shrink:0;height:30px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.card};
        display:flex;align-items:center;padding:0 8px;font-size:12.5px;color:${T.fg};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${op}</div>
      <div style="width:80px;flex-shrink:0;height:30px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.card};
        display:flex;align-items:center;padding:0 8px;font-size:12.5px;color:${T.fg};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${value}</div>
      <span style="color:${T.fgSubtle};font-size:14px;cursor:pointer;">&#10005;</span>
    </div>`;
  }

  const panel = `
  <div style="position:absolute;top:0;right:0;bottom:0;width:440px;background:${T.card};box-shadow:${T.shadowLg};
    border-left:1px solid ${T.border};display:flex;flex-direction:column;">
    <div style="padding:16px 20px;border-bottom:1px solid ${T.border};display:flex;align-items:center;justify-content:space-between;">
      <div style="font-size:15px;font-weight:700;color:${T.fg};">What could we improve?</div>
      <span style="font-size:18px;color:${T.fgSubtle};cursor:pointer;">&#10005;</span>
    </div>
    <div style="display:flex;border-bottom:1px solid ${T.border};padding:0 20px;">
      ${['Options', 'Branching', 'Validation'].map(t => `<div style="padding:10px 14px;font-size:13px;font-weight:600;
        color:${t === 'Branching' ? T.accent : T.fgMuted};border-bottom:2px solid ${t === 'Branching' ? T.accent : 'transparent'};">${t}</div>`).join('')}
    </div>
    <div style="padding:18px 20px;flex:1;overflow-y:auto;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
        <div style="width:34px;height:18px;border-radius:9px;background:${T.accent};position:relative;flex-shrink:0;">
          <div style="width:14px;height:14px;border-radius:50%;background:#fff;position:absolute;top:2px;right:2px;"></div>
        </div>
        <span style="font-size:13px;color:${T.fg};">Only show this question when other answers match</span>
      </div>

      <div style="font-size:11.5px;font-weight:600;color:${T.fgMuted};text-transform:uppercase;letter-spacing:.03em;margin-bottom:8px;">
        Match <strong style="color:${T.fg};">all</strong> of these conditions
      </div>
      ${conditionRow(true, 'How satisfied are you with IT support?', 'is between', '0 &ndash; 6')}

      <div style="margin:4px 0 16px;">
        <span style="font-size:12.5px;color:${T.accent};font-weight:600;cursor:pointer;">+ Add condition</span>
      </div>

      <div style="background:${T.accentTint};border-radius:4px;padding:10px 12px;font-size:12.5px;color:${T.fg};line-height:1.5;">
        Shown when &ldquo;How satisfied are you with IT support?&rdquo; is between 0 and 6.
      </div>
    </div>
    <div style="padding:14px 20px;border-top:1px solid ${T.border};display:flex;justify-content:flex-end;gap:8px;">
      ${btn('Cancel')}${btn('Apply', { primary: true })}
    </div>
  </div>`;

  return pageShell(`<div style="position:relative;padding:18px;">${card(dimmedCanvas)}${panel}</div>`, { width: 1300 });
}

// ── 03: Form settings panel — Appearance tab ──────────────────────────────────
function formSettingsPage() {
  const canvasBg = designerCanvasPage();
  const dimmedCanvas = `<div style="filter:blur(1px) brightness(0.97);pointer-events:none;">${canvasBg.replace(/^[\s\S]*?<body[^>]*>/,'').replace(/<\/body>[\s\S]*$/,'')}</div>`;

  const swatches = [
    { c: '#0078d4', active: true }, { c: '#03787c' }, { c: '#498205' }, { c: '#8764b8' },
    { c: '#881798' }, { c: '#d13438' }, { c: '#ca5010' }, { c: '#69797e' },
  ];
  const headerIcons = ['&#128203;', '&#128172;', '&#128202;', '&#128101;', '&#128197;', '&#127903;', '&#128176;', '&#128722;',
    '&#127973;', '&#128737;', '&#128161;', '&#127942;', '&#9992;', '&#127968;', '&#128092;', '&#127891;', '&#128295;', '&#128227;', '&#128522;', '&#128203;'];

  const panel = `
  <div style="position:absolute;top:0;right:0;bottom:0;width:460px;background:${T.card};box-shadow:${T.shadowLg};
    border-left:1px solid ${T.border};display:flex;flex-direction:column;">
    <div style="padding:16px 20px;border-bottom:1px solid ${T.border};display:flex;align-items:center;justify-content:space-between;">
      <div style="font-size:15px;font-weight:700;color:${T.fg};">Form settings</div>
      <span style="font-size:18px;color:${T.fgSubtle};cursor:pointer;">&#10005;</span>
    </div>
    <div style="display:flex;border-bottom:1px solid ${T.border};padding:0 20px;">
      ${['Basics', 'Appearance', 'After submit', 'Notifications', 'Access'].map(t => `<div style="padding:10px 10px;font-size:12.5px;font-weight:600;
        color:${t === 'Appearance' ? T.accent : T.fgMuted};border-bottom:2px solid ${t === 'Appearance' ? T.accent : 'transparent'};white-space:nowrap;">${t}</div>`).join('')}
    </div>
    <div style="padding:18px 20px;flex:1;overflow-y:auto;">
      <div style="font-size:12.5px;font-weight:600;color:${T.fg};margin-bottom:8px;">Accent color</div>
      <div style="display:flex;gap:10px;margin-bottom:6px;">
        ${swatches.map(s => `<div style="width:28px;height:28px;border-radius:50%;background:${s.c};cursor:pointer;
          ${s.active ? `box-shadow:0 0 0 2px #fff, 0 0 0 3.5px ${s.c};` : ''}"></div>`).join('')}
        <div style="width:28px;height:28px;border-radius:50%;border:1.5px dashed ${T.borderStrong};display:flex;
          align-items:center;justify-content:center;font-size:14px;color:${T.fgSubtle};cursor:pointer;">+</div>
      </div>
      <div style="font-size:11.5px;color:${T.fgSubtle};margin-bottom:18px;">
        Automatically adjusted for contrast on light and dark sites.
      </div>

      <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
        <div style="width:34px;height:18px;border-radius:9px;background:${T.accent};position:relative;flex-shrink:0;">
          <div style="width:14px;height:14px;border-radius:50%;background:#fff;position:absolute;top:2px;right:2px;"></div>
        </div>
        <span style="font-size:13px;color:${T.fg};">Show the form header</span>
      </div>

      <div style="font-size:12.5px;font-weight:600;color:${T.fg};margin-bottom:8px;">Header icon</div>
      <div style="display:grid;grid-template-columns:repeat(8,1fr);gap:6px;">
        ${headerIcons.slice(0, 16).map((e, i) => `<div style="aspect-ratio:1;border-radius:4px;border:1px solid ${i === 1 ? T.accent : T.border};
          background:${i === 1 ? T.accentTint : T.surface};display:flex;align-items:center;justify-content:center;font-size:15px;">${e}</div>`).join('')}
      </div>
    </div>
    <div style="padding:14px 20px;border-top:1px solid ${T.border};display:flex;justify-content:flex-end;gap:8px;">
      ${btn('Cancel')}${btn('Apply', { primary: true })}
    </div>
  </div>`;

  return pageShell(`<div style="position:relative;padding:18px;">${card(dimmedCanvas)}${panel}</div>`, { width: 1300 });
}

// ── 04: Templates gallery ─────────────────────────────────────────────────────
function templatesGalleryPage() {
  const templates = [
    { name: 'Blank form', desc: 'Start from nothing', meta: 'Empty', icon: '&#128196;', color: '#0078d4' },
    { name: 'Customer feedback', desc: 'NPS, ratings and open comments', meta: '6 questions', icon: '&#128172;', color: '#0078d4' },
    { name: 'Event registration', desc: 'Attendee details, sessions and dietary needs', meta: '9 questions &middot; wizard', icon: '&#128197;', color: '#8764b8' },
    { name: 'IT service request', desc: 'Categorized requests with priority and attachments', meta: '7 questions', icon: '&#128295;', color: '#03787c' },
    { name: 'Expense claim', desc: 'Line items, receipts and a calculated total', meta: '10 questions', icon: '&#128176;', color: '#498205' },
    { name: 'Employee pulse survey', desc: 'Likert grid, engagement scale and anonymity note', meta: '4 questions', icon: '&#128522;', color: '#881798' },
    { name: 'Safety inspection', desc: 'Checklist, photos and a signature', meta: '8 questions', icon: '&#128737;', color: '#ca5010' },
    { name: 'Room or resource booking', desc: 'Lookup-driven resource picker with a time slot', meta: '8 questions', icon: '&#127915;', color: '#69797e' },
  ];

  const body = `
  <div style="padding:22px 26px;">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
      <div style="font-size:19px;font-weight:700;color:${T.fg};">Start from a template</div>
      <span style="font-size:18px;color:${T.fgSubtle};">&#10005;</span>
    </div>
    <div style="font-size:13px;color:${T.fgMuted};margin-bottom:18px;">Pick a starting point &mdash; every template can be fully edited afterwards.</div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;">
      ${templates.map(t => `<div style="border:1px solid ${T.border};border-radius:6px;padding:16px;cursor:pointer;background:${T.card};">
        <div style="width:36px;height:36px;border-radius:6px;background:${t.color}1a;display:flex;align-items:center;
          justify-content:center;font-size:18px;margin-bottom:10px;">${t.icon}</div>
        <div style="font-size:13.5px;font-weight:700;color:${T.fg};margin-bottom:4px;">${t.name}</div>
        <div style="font-size:11.5px;color:${T.fgMuted};line-height:1.4;margin-bottom:10px;min-height:32px;">${t.desc}</div>
        <div style="font-size:11px;color:${T.fgSubtle};font-weight:600;">${t.meta}</div>
      </div>`).join('')}
    </div>
  </div>`;

  return pageShell(`<div style="padding:18px;">${card(body)}</div>`, { width: 1180 });
}

// ── 05: Collect responses / Share dialog ──────────────────────────────────────
function shareDialogPage() {
  const body = `
  <div style="padding:24px 28px;">
    <div style="font-size:17px;font-weight:700;color:${T.fg};margin-bottom:4px;">Your form is ready to share</div>
    <div style="font-size:12.5px;color:${T.fgMuted};margin-bottom:16px;">Anyone with this link sees just the form &mdash; not the designer or results.</div>
    <div style="display:flex;gap:8px;margin-bottom:18px;">
      <div style="flex:1;height:34px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};
        display:flex;align-items:center;padding:0 10px;font-size:12px;color:${T.fgMuted};font-family:monospace;
        overflow:hidden;white-space:nowrap;">https://contoso.sharepoint.com/sites/Ops/SitePages/IT-Feedback.aspx?sfview=fill</div>
      ${btn('Copy', { primary: true })}
    </div>
    <div style="display:flex;gap:10px;padding:10px 12px;background:${T.accentTint};border-radius:4px;margin-bottom:10px;">
      ${icon(ICONS.info, { size: 15, color: T.accent })}
      <span style="font-size:12px;color:${T.fg};line-height:1.5;">Respondents need <strong>Add Items</strong> permission on the response list to submit the form.</span>
    </div>
    <div style="display:flex;gap:10px;padding:10px 12px;background:${T.warningTint};border-radius:4px;">
      ${icon(ICONS.bulb, { size: 15, color: T.warning })}
      <span style="font-size:12px;color:${T.fg};line-height:1.5;">Add <code style="background:#fff;padding:1px 4px;border-radius:2px;">&amp;SFDepartment=Marketing</code> to the link to pre-answer a question.</span>
    </div>
    <div style="display:flex;justify-content:flex-end;margin-top:18px;">${btn('Done', { primary: true })}</div>
  </div>`;

  return pageShell(`<div style="padding:40px;display:flex;justify-content:center;">${card(body, { radius: 8 })}</div>`, { width: 820 });
}

// ── 06: Form filler — single page ─────────────────────────────────────────────
function fillerSinglePage() {
  function question(num, title, required, bodyHtml, opts = {}) {
    return `<div style="margin-bottom:22px;">
      <div style="font-size:14px;font-weight:600;color:${T.fg};margin-bottom:8px;">
        ${opts.number !== false ? `${num}. ` : ''}${title} ${required ? `<span style="color:${T.error};">*</span>` : ''}
      </div>
      ${bodyHtml}
    </div>`;
  }

  const likertRows = ['My requests are resolved quickly', 'Support staff are knowledgeable', 'I would recommend IT support to a colleague'];
  const likertCols = ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree'];

  const body = `
  <div style="padding:26px 30px;">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:6px;">
      <div style="width:40px;height:40px;border-radius:50%;background:${T.accent};display:flex;align-items:center;
        justify-content:center;color:#fff;font-size:18px;flex-shrink:0;">&#128172;</div>
      <div style="font-size:21px;font-weight:700;color:${T.fg};">IT Support Feedback</div>
    </div>
    <div style="font-size:13px;color:${T.fgMuted};margin-bottom:22px;padding-left:52px;">Tell us about your recent support experience &mdash; it takes under two minutes.</div>

    ${question(1, 'How satisfied are you with IT support?', true, `
      <div style="display:flex;align-items:center;gap:6px;">
        ${[0,1,2,3,4,5,6,7,8,9,10].map(n => `<div style="width:30px;height:30px;border-radius:4px;border:1px solid ${T.borderStrong};
          display:flex;align-items:center;justify-content:center;font-size:12px;color:${T.fg};cursor:pointer;
          ${n === 9 ? `background:${T.accent};color:#fff;border-color:${T.accent};` : ''}">${n}</div>`).join('')}
      </div>
      <div style="display:flex;justify-content:space-between;font-size:11.5px;color:${T.fgSubtle};margin-top:5px;">
        <span>Not at all likely</span><span>Extremely likely</span>
      </div>`)}

    ${question(2, 'Rate your most recent support interaction', true, `
      <div style="display:flex;align-items:center;gap:4px;">
        ${[1,2,3,4,5].map(n => icon('M12 2l2.9 6.6 7.1.6-5.4 4.7 1.6 7-6.2-3.7L5.8 21l1.6-7L2 9.2l7.1-.6z', { size: 26, color: n <= 4 ? '#f0b400' : T.border })).join('')}
        <span style="margin-left:8px;font-size:13px;color:${T.fgMuted};">4 / 5</span>
      </div>`)}

    ${question(3, 'Tell us how we could improve', false, `
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr>
          <td></td>
          ${likertCols.map(c => `<td style="text-align:center;color:${T.fgMuted};font-size:11px;padding:4px;">${c}</td>`).join('')}
        </tr></thead>
        <tbody>
          ${likertRows.map((r, ri) => `<tr style="border-top:1px solid ${T.border};">
            <td style="padding:8px 8px 8px 0;color:${T.fg};">${r}</td>
            ${likertCols.map((c, ci) => `<td style="text-align:center;padding:8px;">
              <div style="width:15px;height:15px;border-radius:50%;border:1.5px solid ${T.borderStrong};margin:0 auto;
                ${ri === 0 && ci === 3 ? `background:${T.accent};border-color:${T.accent};` : ''}
                ${ri === 1 && ci === 4 ? `background:${T.accent};border-color:${T.accent};` : ''}
                ${ri === 2 && ci === 3 ? `background:${T.accent};border-color:${T.accent};` : ''}"></div>
            </td>`).join('')}
          </tr>`).join('')}
        </tbody>
      </table>`)}

    ${question(4, 'Would you recommend us to a colleague?', true, `
      <div style="display:flex;gap:10px;">
        <div style="padding:8px 22px;border-radius:3px;border:1.5px solid ${T.accent};background:${T.accentTint};
          color:${T.accent};font-weight:600;font-size:13px;">Yes</div>
        <div style="padding:8px 22px;border-radius:3px;border:1.5px solid ${T.borderStrong};color:${T.fg};font-size:13px;">No</div>
      </div>`)}

    <div style="display:flex;align-items:center;justify-content:space-between;margin-top:28px;padding-top:16px;
      border-top:1px solid ${T.border};">
      <span style="font-size:12.5px;color:${T.fgMuted};display:flex;align-items:center;gap:6px;cursor:pointer;">Save and finish later</span>
      ${btn('Submit', { primary: true })}
    </div>
  </div>`;

  return pageShell(`<div style="padding:24px;display:flex;justify-content:center;">
    <div style="width:760px;">${card(body)}</div></div>`, { width: 860 });
}

// ── 07: Form filler — wizard mode ─────────────────────────────────────────────
function fillerWizardPage() {
  const steps = ['Your details', 'Session &amp; dietary', 'Confirm'];
  const stepDots = steps.map((s, i) => `
    <div style="display:flex;flex-direction:column;align-items:center;gap:6px;flex:1;">
      <div style="width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;
        font-size:12px;font-weight:700;
        ${i < 1 ? `background:${T.accent};color:#fff;` : i === 1 ? `background:${T.accent};color:#fff;box-shadow:0 0 0 4px ${T.accentTint};` : `background:${T.surface2};color:${T.fgMuted};`}">
        ${i < 1 ? '&#10003;' : i + 1}</div>
      <span style="font-size:11.5px;color:${i <= 1 ? T.fg : T.fgSubtle};font-weight:${i === 1 ? 700 : 500};text-align:center;">${s}</span>
    </div>`).join(`<div style="flex:2;height:2px;background:${T.border};margin-top:14px;"></div>`);

  const body = `
  <div style="padding:26px 30px;">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:22px;">
      <div style="width:40px;height:40px;border-radius:50%;background:#8764b8;display:flex;align-items:center;
        justify-content:center;color:#fff;font-size:18px;flex-shrink:0;">&#128197;</div>
      <div style="font-size:21px;font-weight:700;color:${T.fg};">Annual Conference &mdash; Registration</div>
    </div>

    <div style="display:flex;align-items:flex-start;margin-bottom:8px;">${stepDots}</div>
    <div style="height:4px;background:${T.surface2};border-radius:2px;margin-bottom:26px;overflow:hidden;">
      <div style="width:50%;height:100%;background:${T.accent};"></div>
    </div>

    <div style="font-size:16px;font-weight:700;color:${T.fg};margin-bottom:4px;">Session &amp; dietary</div>
    <div style="font-size:12.5px;color:${T.fgMuted};margin-bottom:18px;">Pick the sessions you'd like to attend and let us know about any dietary needs.</div>

    <div style="margin-bottom:20px;">
      <div style="font-size:14px;font-weight:600;color:${T.fg};margin-bottom:8px;">Which sessions will you attend? <span style="color:${T.error};">*</span></div>
      ${['Keynote &mdash; The future of collaboration', 'Workshop &mdash; SPFx deep dive', 'Panel &mdash; Security &amp; compliance', 'Networking lunch'].map((o, i) => `
        <label style="display:flex;align-items:center;gap:9px;padding:8px 0;font-size:13px;color:${T.fg};">
          <div style="width:16px;height:16px;border:1.5px solid ${i < 2 ? T.accent : T.borderStrong};border-radius:3px;
            background:${i < 2 ? T.accent : '#fff'};display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;flex-shrink:0;">
            ${i < 2 ? '&#10003;' : ''}</div>${o}
        </label>`).join('')}
    </div>

    <div style="margin-bottom:24px;">
      <div style="font-size:14px;font-weight:600;color:${T.fg};margin-bottom:8px;">Dietary requirements (optional)</div>
      <div style="height:34px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.surface};max-width:340px;"></div>
    </div>

    <div style="display:flex;align-items:center;justify-content:space-between;padding-top:16px;border-top:1px solid ${T.border};">
      ${btn('Back', { iconD: ICONS.chevronL })}
      ${btn('Next', { primary: true, iconD: ICONS.chevronR })}
    </div>
  </div>`;

  return pageShell(`<div style="padding:24px;display:flex;justify-content:center;">
    <div style="width:720px;">${card(body)}</div></div>`, { width: 820 });
}

// ── 08: Confirmation screen ────────────────────────────────────────────────────
function confirmationPage() {
  const body = `
  <div style="padding:60px 40px;text-align:center;">
    <div style="width:64px;height:64px;border-radius:50%;background:${T.successTint};display:flex;align-items:center;
      justify-content:center;margin:0 auto 18px;">${icon(ICONS.check, { size: 30, color: T.success })}</div>
    <div style="font-size:20px;font-weight:700;color:${T.fg};margin-bottom:8px;">Thanks for your feedback!</div>
    <div style="font-size:13.5px;color:${T.fgMuted};max-width:380px;margin:0 auto 22px;line-height:1.6;">
      Your response has been recorded. Our team reviews every submission weekly.</div>
    ${btn('Submit another response', { iconD: ICONS.refresh })}
  </div>`;

  return pageShell(`<div style="padding:40px;display:flex;justify-content:center;">
    <div style="width:560px;">${card(body)}</div></div>`, { width: 640 });
}

// ── 09: Results dashboard ──────────────────────────────────────────────────────
function statTile(label, value, opts = {}) {
  const { delta = null, caption = null, sparkline = true } = opts;
  const pts = [4,7,5,9,6,10,8,12,9,13,11,15];
  const max = Math.max(...pts);
  const poly = pts.map((p, i) => `${(i / (pts.length - 1)) * 100},${28 - (p / max) * 24}`).join(' ');
  return `<div style="flex:1;min-width:150px;background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:14px 16px;">
    <div style="font-size:11.5px;color:${T.fgMuted};margin-bottom:6px;">${label}</div>
    <div style="display:flex;align-items:flex-end;gap:8px;">
      <div style="font-size:24px;font-weight:700;color:${T.fg};">${value}</div>
      ${delta ? `<span style="font-size:11.5px;font-weight:700;color:${delta.up ? T.success : T.error};margin-bottom:4px;">
        ${delta.up ? '&#9650;' : '&#9660;'} ${delta.text}</span>` : ''}
    </div>
    ${sparkline ? `<svg width="100%" height="28" viewBox="0 0 100 28" preserveAspectRatio="none" style="margin-top:6px;">
      <polyline points="${poly}" fill="none" stroke="${T.accent}" stroke-width="2"/>
    </svg>` : ''}
    ${caption ? `<div style="font-size:10.5px;color:${T.fgSubtle};margin-top:2px;">${caption}</div>` : ''}
  </div>`;
}

function timelineChart() {
  const days = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun','Mon','Tue','Wed','Thu','Fri'];
  const vals = [3,5,4,8,6,2,1,7,9,6,10,12];
  const max = Math.max(...vals);
  const w = 640, h = 130, pad = 6;
  const pts = vals.map((v, i) => `${pad + (i / (vals.length - 1)) * (w - pad * 2)},${h - pad - (v / max) * (h - pad * 2)}`).join(' ');
  const areaPts = `${pad},${h - pad} ${pts} ${w - pad},${h - pad}`;
  return `<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">
    <polygon points="${areaPts}" fill="${T.accent}" opacity="0.12"/>
    <polyline points="${pts}" fill="none" stroke="${T.accent}" stroke-width="2.5"/>
  </svg>
  <div style="display:flex;justify-content:space-between;font-size:10px;color:${T.fgSubtle};margin-top:2px;">
    ${days.filter((_, i) => i % 2 === 0).map(d => `<span>${d}</span>`).join('')}
  </div>`;
}

function summaryCardShell(title, meta, chartHtml, opts = {}) {
  const { span2 = false } = opts;
  return `<div style="background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:14px 16px;
    ${span2 ? 'grid-column: span 2;' : ''}">
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:2px;">
      <span style="font-size:13px;font-weight:700;color:${T.fg};">${title}</span>
      <span style="margin-left:auto;font-size:14px;color:${T.fgSubtle};cursor:pointer;">&#8942;</span>
    </div>
    <div style="font-size:11px;color:${T.fgSubtle};margin-bottom:10px;">${meta}</div>
    ${chartHtml}
  </div>`;
}

function barsChart(rows, opts = {}) {
  const max = Math.max(...rows.map(r => r.v));
  return `<div style="display:flex;flex-direction:column;gap:7px;">
    ${rows.map((r, i) => `<div style="display:flex;align-items:center;gap:8px;">
      <span style="width:110px;font-size:11.5px;color:${T.fg};flex-shrink:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;">${r.l}</span>
      <div style="flex:1;background:${T.surface2};border-radius:3px;height:16px;overflow:hidden;">
        <div style="width:${(r.v/max)*100}%;height:100%;background:${CATEGORICAL[i % CATEGORICAL.length]};border-radius:3px;"></div>
      </div>
      <span style="width:26px;font-size:11px;color:${T.fgMuted};text-align:right;">${r.v}</span>
    </div>`).join('')}
  </div>`;
}

function gaugeChart(score) {
  const pct = (score + 100) / 200;
  const angle = Math.PI * (1 - pct);
  const cx = 90, cy = 80, r = 70;
  const x = cx + r * Math.cos(angle), y = cy - r * Math.sin(angle);
  return `<svg width="180" height="95" viewBox="0 0 180 95" style="display:block;margin:0 auto;">
    <path d="M20 80 A70 70 0 0 1 65 15" stroke="${T.error}" stroke-width="12" fill="none"/>
    <path d="M65 15 A70 70 0 0 1 115 15" stroke="${T.warning}" stroke-width="12" fill="none"/>
    <path d="M115 15 A70 70 0 0 1 160 80" stroke="${T.success}" stroke-width="12" fill="none"/>
    <line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="${T.fg}" stroke-width="3"/>
    <circle cx="${cx}" cy="${cy}" r="4" fill="${T.fg}"/>
    <text x="90" y="92" text-anchor="middle" font-size="16" font-weight="700" fill="${T.fg}" font-family="${F}">${score}</text>
  </svg>`;
}

function dashboardPage() {
  const body = `
  <div style="padding:20px 24px;">
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">
      ${statTile('Total responses', '482', { caption: 'Latest 2 min ago' })}
      ${statTile('Last 7 days', '96', { delta: { up: true, text: '18%' } })}
      ${statTile('Today', '14', { caption: 'Since midnight', sparkline: false })}
      ${statTile('Unique respondents', '441', { caption: '31 repeat', sparkline: false })}
      ${statTile('Completion rate', '87%', { delta: { up: true, text: '3%' } })}
      ${statTile('Median time to complete', '2m 14s', { sparkline: false })}
    </div>

    <div style="background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:14px 16px;margin-bottom:16px;">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
        <span style="font-size:13.5px;font-weight:700;color:${T.fg};">Responses over time</span>
        <div style="font-size:11.5px;color:${T.fgMuted};border:1px solid ${T.borderStrong};border-radius:3px;padding:4px 10px;">By day &#9662;</div>
      </div>
      ${timelineChart()}
    </div>

    <div style="font-size:13.5px;font-weight:700;color:${T.fg};margin-bottom:10px;">Highlights</div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">
      ${summaryCardShell('Would you recommend us to a colleague?', '482 of 482 answered', gaugeChart(42))}
      ${summaryCardShell('Which department are you in?', '470 of 482 answered', barsChart([
        { l: 'Engineering', v: 168 }, { l: 'Sales', v: 122 }, { l: 'Operations', v: 96 }, { l: 'Marketing', v: 84 },
      ]))}
      ${summaryCardShell('Rate your most recent support interaction', '478 of 482 answered', `
        <div style="display:flex;align-items:baseline;gap:8px;">
          <span style="font-size:28px;font-weight:700;color:${T.fg};">4.3</span>
          <span style="font-size:12px;color:${T.fgMuted};">/ 5 average</span>
        </div>`)}
    </div>
  </div>`;

  return pageShell(`<div style="padding:18px;">${card(ownerBar({ activeTab: 'responses', shareLabel: 'Share' }) + previewBannerOff() + body)}</div>`, { width: 1300 });
}
function previewBannerOff() { return ''; }

// ── 10: Summary view — Compare by ─────────────────────────────────────────────
function summaryComparePage() {
  function segmentBlock(label, count, rows) {
    return `<div style="margin-bottom:12px;">
      <div style="font-size:11.5px;font-weight:700;color:${T.fgMuted};margin-bottom:6px;">${label} &middot; ${count} response(s)</div>
      ${barsChart(rows)}
    </div>`;
  }

  const body = `
  <div style="padding:20px 24px;">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
      <div style="font-size:12px;color:${T.fgMuted};border:1px solid ${T.borderStrong};border-radius:3px;padding:6px 12px;
        background:${T.card};">Compare by&hellip; &#9662;</div>
      ${chip('&#128257; Split by Which department are you in?', { bg: T.accentTint, color: T.accent })}
    </div>

    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;">
      <div style="background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:14px 16px;">
        <div style="font-size:13px;font-weight:700;color:${T.fg};margin-bottom:2px;">Would you recommend us to a colleague?</div>
        <div style="font-size:11px;color:${T.fgSubtle};margin-bottom:10px;">482 of 482 answered &middot; Yes / No</div>
        ${segmentBlock('Engineering', 168, [{ l: 'Yes', v: 151 }, { l: 'No', v: 17 }])}
        ${segmentBlock('Sales', 122, [{ l: 'Yes', v: 89 }, { l: 'No', v: 33 }])}
        ${segmentBlock('Operations', 96, [{ l: 'Yes', v: 80 }, { l: 'No', v: 16 }])}
      </div>
      <div style="background:${T.card};border:1px solid ${T.border};border-radius:6px;padding:14px 16px;">
        <div style="font-size:13px;font-weight:700;color:${T.fg};margin-bottom:2px;">How satisfied are you with IT support?</div>
        <div style="font-size:11px;color:${T.fgSubtle};margin-bottom:10px;">482 of 482 answered &middot; NPS (0&ndash;10)</div>
        ${segmentBlock('Engineering', 168, [{ l: 'Promoters', v: 120 }, { l: 'Passives', v: 32 }, { l: 'Detractors', v: 16 }])}
        ${segmentBlock('Sales', 122, [{ l: 'Promoters', v: 61 }, { l: 'Passives', v: 40 }, { l: 'Detractors', v: 21 }])}
        ${segmentBlock('Operations', 96, [{ l: 'Promoters', v: 74 }, { l: 'Passives', v: 15 }, { l: 'Detractors', v: 7 }])}
      </div>
    </div>
  </div>`;

  return pageShell(`<div style="padding:18px;">${card(ownerBar({ activeTab: 'responses', shareLabel: 'Share' }) + body)}</div>`, { width: 1300 });
}

// ── 11: Response table ────────────────────────────────────────────────────────
function responseTablePage() {
  const cols = ['Submitted', 'Submitted by', 'How satisfied are you with IT support?', 'Would you recommend us?', 'Department'];
  const rows = [
    ['Today 10:42 AM', 'Alice Chen', '9', 'Yes', 'Engineering'],
    ['Today 9:15 AM', 'Bob Martinez', '6', 'No', 'Sales'],
    ['Yesterday 4:03 PM', 'Priya Nair', '10', 'Yes', 'Engineering'],
    ['Yesterday 2:47 PM', 'Jane Smith', '7', 'Yes', 'Operations'],
    ['Yesterday 11:20 AM', 'Tom Fields', '3', 'No', 'Marketing'],
  ];

  const body = `
  <div style="padding:0 0 8px;">
    <div style="display:flex;align-items:center;gap:8px;padding:12px 20px;border-bottom:1px solid ${T.border};">
      ${btn('Columns (5/9)')}${btn('Reorder')}
      <div style="margin-left:auto;height:30px;border:1px solid ${T.borderStrong};border-radius:3px;background:${T.card};
        display:flex;align-items:center;padding:0 10px;font-size:12px;color:${T.fgSubtle};width:200px;">Search responses&hellip;</div>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="border-bottom:1px solid ${T.border};background:${T.surface};">
          ${cols.map((c, i) => `<th style="text-align:left;padding:9px 16px;font-size:12px;font-weight:700;color:${T.fg};
            white-space:nowrap;">${c} ${i === 0 ? '&#9650;' : ''}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${rows.map((r, ri) => `<tr style="border-bottom:1px solid ${T.border};${ri % 2 ? `background:${T.surface};` : ''}">
          ${r.map((v, ci) => `<td style="padding:9px 16px;font-size:12.5px;color:${ci === 0 ? T.accent : T.fg};
            ${ci === 0 ? 'text-decoration:underline;cursor:pointer;' : ''}">${v}</td>`).join('')}
        </tr>`).join('')}
      </tbody>
    </table>
  </div>`;

  return pageShell(`<div style="padding:18px;">${card(ownerBar({ activeTab: 'responses', shareLabel: 'Share' }) + body)}</div>`, { width: 1300 });
}

// ── 12: Response detail panel ─────────────────────────────────────────────────
function responseDetailPage() {
  const canvasBg = dashboardPage();
  const dimmed = `<div style="filter:blur(1px) brightness(0.97);pointer-events:none;">${canvasBg.replace(/^[\s\S]*?<body[^>]*>/,'').replace(/<\/body>[\s\S]*$/,'')}</div>`;

  const answers = [
    ['Full name', 'Alice Chen'],
    ['Work email address', 'alice.chen@contoso.com'],
    ['Which department are you in?', 'Engineering'],
    ['How satisfied are you with IT support?', '9'],
    ['What could we improve?', '"Faster response time on P2 tickets."'],
    ['Would you recommend us to a colleague?', 'Yes'],
  ];

  const panel = `
  <div style="position:absolute;top:0;right:0;bottom:0;width:480px;background:${T.card};box-shadow:${T.shadowLg};
    border-left:1px solid ${T.border};display:flex;flex-direction:column;">
    <div style="padding:16px 20px;border-bottom:1px solid ${T.border};">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
        <div style="font-size:15px;font-weight:700;color:${T.fg};">Response 3 of 482</div>
        <span style="font-size:18px;color:${T.fgSubtle};cursor:pointer;">&#10005;</span>
      </div>
      <div style="display:flex;align-items:center;gap:14px;font-size:11.5px;color:${T.fgMuted};flex-wrap:wrap;">
        <span style="display:flex;align-items:center;gap:4px;">${icon(ICONS.clock, { size: 13, color: T.fgSubtle })} Today, 10:42 AM</span>
        <span style="display:flex;align-items:center;gap:4px;">${icon(ICONS.contact, { size: 13, color: T.fgSubtle })} Alice Chen</span>
        <span>took 1m 48s</span>
      </div>
    </div>
    <div style="display:flex;border-bottom:1px solid ${T.border};padding:0 20px;">
      ${['Answers', 'As the form'].map((t, i) => `<div style="padding:10px 14px;font-size:13px;font-weight:600;
        color:${i === 1 ? T.accent : T.fgMuted};border-bottom:2px solid ${i === 1 ? T.accent : 'transparent'};">${t}</div>`).join('')}
    </div>
    <div style="padding:18px 20px;flex:1;overflow-y:auto;">
      <div style="font-size:19px;font-weight:700;color:${T.fg};margin-bottom:14px;">IT Support Feedback</div>
      ${answers.map(([q, a]) => `<div style="margin-bottom:14px;">
        <div style="font-size:12.5px;font-weight:600;color:${T.fg};margin-bottom:4px;">${q}</div>
        <div style="font-size:13px;color:${T.fgMuted};">${a}</div>
      </div>`).join('')}
    </div>
    <div style="padding:12px 20px;border-top:1px solid ${T.border};display:flex;align-items:center;gap:8px;">
      ${iconBtn(ICONS.chevronL, { size: 30 })}${iconBtn(ICONS.chevronR, { size: 30 })}
      ${btn('Print', { iconD: ICONS.print })}
      <span style="margin-left:auto;"></span>
    </div>
  </div>`;

  return pageShell(`<div style="position:relative;padding:18px;">${card(dimmed)}${panel}</div>`, { width: 1300 });
}

// ── Runner ────────────────────────────────────────────────────────────────────
async function main() {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--font-render-hinting=none'],
  });

  const shots = [
    ['01_designer_canvas.png', designerCanvasPage, { width: 1300, height: 1080 }],
    ['02_field_editor_branching.png', fieldEditorPage, { width: 1300, height: 900 }],
    ['03_form_settings_appearance.png', formSettingsPage, { width: 1300, height: 900 }],
    ['04_templates_gallery.png', templatesGalleryPage, { width: 1180, height: 620 }],
    ['05_collect_responses_share.png', shareDialogPage, { width: 820, height: 480 }],
    ['06_form_filler_singlepage.png', fillerSinglePage, { width: 860, height: 1080 }],
    ['07_form_filler_wizard.png', fillerWizardPage, { width: 820, height: 720 }],
    ['08_form_confirmation.png', confirmationPage, { width: 640, height: 420 }],
    ['09_dashboard.png', dashboardPage, { width: 1300, height: 960 }],
    ['10_summary_compare.png', summaryComparePage, { width: 1300, height: 820 }],
    ['11_response_table.png', responseTablePage, { width: 1300, height: 560 }],
    ['12_response_detail.png', responseDetailPage, { width: 1300, height: 900 }],
  ];

  for (const [filename, htmlFn, vp] of shots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'load' });
    await pg.screenshot({ path: path.join(OUT, filename) });
    await pg.close();
    console.log('✓', filename);
  }

  await browser.close();
  console.log('\nAll screenshots saved to docs/screenshots/');
}

main().catch(e => { console.error(e); process.exit(1); });
