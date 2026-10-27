// Tempo — tüm veri tabletin localStorage'ında tutulur.
const KEY = 'etut-sayaci-v1';
let studies = load();
let openId = null; // detay ekranındaki etüt

function load() {
  try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch { return []; }
}
function save() { localStorage.setItem(KEY, JSON.stringify(studies)); }
const now = () => Date.now();
const uid = () => Math.random().toString(36).slice(2, 10);
const byId = id => studies.find(s => s.id === id);

// Her işlem gerçek saatiyle olay kaydına yazılır (ekranda görünmez, Excel'e çıkar).
function log(s, type, detail = '') { s.events.push({ type, at: now(), detail }); }

function elapsed(s) {
  return s.segments.reduce((t, g) => t + ((g.end ?? now()) - g.start), 0);
}
const visibleLaps = s => s.laps.filter(l => !l.deleted);
// Tur süresi: bir önceki silinmemiş bayraktan (yoksa başlangıçtan) bu yana
function lapRows(s) {
  let prev = 0;
  return visibleLaps(s).map((l, i) => { const r = { ...l, sira: i + 1, tur: l.elapsed - prev }; prev = l.elapsed; return r; });
}

function fmt(ms) {
  const t = Math.floor(ms / 100), ds = t % 10, sec = Math.floor(t / 10) % 60, min = Math.floor(t / 600) % 60, h = Math.floor(t / 36000);
  const p = n => String(n).padStart(2, '0');
  return (h ? h + ':' : '') + p(min) + ':' + p(sec) + ',' + ds;
}
const clock = ts => ts ? new Date(ts).toLocaleTimeString('tr-TR') : '—';
const dateTime = ts => ts ? new Date(ts).toLocaleString('tr-TR') : '';
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// --- İşlemler ---
function newStudy() {
  const name = prompt('Etüt adı:', 'Etüt ' + (studies.length + 1));
  if (name === null) return;
  const s = { id: uid(), name: name.trim() || 'Etüt', createdAt: now(), status: 'idle', segments: [], laps: [], events: [], startedAt: null, finishedAt: null };
  log(s, 'oluşturuldu', s.name);
  studies.unshift(s); save(); render();
}
function start(s) {
  if (s.status === 'idle') { s.startedAt = now(); log(s, 'başlat'); }
  else if (s.status === 'finished') { s.finishedAt = null; log(s, 'bitmişten devam'); }
  else log(s, 'devam');
  s.segments.push({ start: now(), end: null }); s.status = 'running'; save(); render();
}
function pause(s) {
  s.segments.at(-1).end = now(); s.status = 'paused'; log(s, 'duraklat', fmt(elapsed(s))); save(); render();
}
function flag(s) {
  const l = { id: uid(), no: s.laps.length + 1, elapsed: elapsed(s), at: now(), note: s.pendingNote || '', deleted: false };
  if (l.note) l.loggedNote = l.note;
  s.pendingNote = '';
  s.laps.push(l); log(s, 'bayrak', '#' + l.no + ' ' + fmt(l.elapsed) + (l.note ? ' · ' + l.note : '')); save(); render();
}
function finish(s) {
  if (s.status === 'running') s.segments.at(-1).end = now();
  s.status = 'finished'; s.finishedAt = now(); log(s, 'bitir', fmt(elapsed(s))); save(); render();
}
function delLap(s, lapId) {
  const l = s.laps.find(x => x.id === lapId);
  l.deleted = true; l.deletedAt = now(); log(s, 'bayrak sil', '#' + l.no + ' ' + fmt(l.elapsed)); save(); render();
}
function toggleTag(s, lapId) {
  const l = s.laps.find(x => x.id === lapId);
  l.tag = !l.tag; log(s, l.tag ? 'etiketle' : 'etiket kaldır', '#' + l.no); save(); render();
}
function setNote(s, lapId, note, final) {
  const l = s.laps.find(x => x.id === lapId);
  if (final) { if (l.loggedNote !== note) { l.loggedNote = note; l.noteAt = now(); log(s, 'not', '#' + l.no + ': ' + note); save(); } return; }
  l.note = note; save();
}
function rename(s) {
  const name = prompt('Yeni ad:', s.name);
  if (!name || !name.trim()) return;
  log(s, 'yeniden adlandır', s.name + ' → ' + name.trim()); s.name = name.trim(); save(); render();
}
function removeStudy(s) {
  if (!confirm('"' + s.name + '" ve tüm kayıtları SİLİNSİN mi? Geri alınamaz.')) return;
  studies = studies.filter(x => x !== s); openId = null; save(); render();
}

// --- Excel ---
function sheetName(name, used) {
  let base = name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 28) || 'Etüt', n = base, i = 2;
  while (used.has(n)) n = base.slice(0, 26) + ' ' + i++;
  used.add(n); return n;
}
function exportXlsx(list) {
  if (!list.length) return alert('Dışa aktarılacak etüt yok.');
  const wb = XLSX.utils.book_new(), used = new Set(['Özet', 'Olay Kaydı']);
  const ozet = [['Etüt', 'Durum', 'Başlangıç', 'Bitiş', 'Toplam süre', 'Toplam (sn)', 'Bayrak sayısı', 'Silinen bayrak']];
  const olay = [['Etüt', 'Gerçek zaman', 'İşlem', 'Detay']];
  for (const s of list) {
    ozet.push([s.name, durum(s), dateTime(s.startedAt), dateTime(s.finishedAt), fmt(elapsed(s)), +(elapsed(s) / 1000).toFixed(1), visibleLaps(s).length, s.laps.length - visibleLaps(s).length]);
    const rows = [['Etüt', s.name], ['Başlangıç', dateTime(s.startedAt)], ['Bitiş', dateTime(s.finishedAt)], ['Yeşil etiket anlamı', s.tagMeaning || ''], [],
      ['Sıra', 'Tur süresi', 'Tur (sn)', 'Toplam', 'Toplam (sn)', 'Gerçek saat', 'Not', 'Etiket']];
    for (const r of lapRows(s)) rows.push([r.sira, fmt(r.tur), +(r.tur / 1000).toFixed(1), fmt(r.elapsed), +(r.elapsed / 1000).toFixed(1), dateTime(r.at), r.note, r.tag ? (s.tagMeaning || 'Yeşil') : '']);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 11 }, { wch: 20 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, ws, sheetName(s.name, used));
    for (const e of s.events) olay.push([s.name, dateTime(e.at), e.type, e.detail]);
  }
  const wo = XLSX.utils.aoa_to_sheet(ozet); wo['!cols'] = [{ wch: 24 }, { wch: 10 }, { wch: 20 }, { wch: 20 }, { wch: 12 }, { wch: 11 }, { wch: 12 }, { wch: 14 }];
  const we = XLSX.utils.aoa_to_sheet(olay); we['!cols'] = [{ wch: 24 }, { wch: 20 }, { wch: 16 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, wo, 'Özet', true);
  XLSX.utils.book_append_sheet(wb, we, 'Olay Kaydı');
  wb.SheetNames = ['Özet', ...wb.SheetNames.filter(n => n !== 'Özet')];
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const stamp = d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
  const file = (list.length === 1 ? list[0].name.replace(/[^\wçğıöşüÇĞİÖŞÜ -]/g, '') : 'Etutler') + '_' + stamp + '.xlsx';
  XLSX.writeFile(wb, file);
}

// --- Arayüz ---
const durum = s => ({ idle: 'Hazır', running: 'Çalışıyor', paused: 'Duraklatıldı', finished: 'Bitti' }[s.status]);
const app = document.getElementById('app');

function controls(s, big) {
  const r = s.status === 'running', f = s.status === 'finished';
  return `<div class="row">
      ${r ? `<button data-a="pause" data-id="${s.id}">Duraklat</button>`
          : `<button class="ok" data-a="start" data-id="${s.id}">${s.status === 'idle' ? 'Başlat' : 'Devam'}</button>`}
      <button class="bad" data-a="finish" data-id="${s.id}" ${s.status === 'idle' || f ? 'disabled' : ''}>Bitir</button>
    </div>
    <button class="acc flag" data-a="flag" data-id="${s.id}" ${r ? '' : 'disabled'}>⚑ Bayrak</button>`;
}

function render() {
  const f = document.activeElement, keep = (f?.dataset?.note || f?.dataset?.meaning || f?.dataset?.pend || f?.dataset?.name) ? { id: f.dataset.note || f.dataset.meaning || f.dataset.pend || f.dataset.name, a: f.selectionStart, b: f.selectionEnd } : null;
  draw();
  if (keep) { const el = app.querySelector(`[data-note="${keep.id}"],[data-meaning="${keep.id}"],[data-pend="${keep.id}"],[data-name="${keep.id}"]`); if (el) { el.focus(); el.setSelectionRange(keep.a, keep.b); } }
}
function draw() {
  const s = openId && byId(openId);
  document.getElementById('back').hidden = !s;
  document.getElementById('newStudy').hidden = !!s;
  document.getElementById('exportAll').textContent = s ? 'Excel' : 'Excel (tümü)';
  document.getElementById('title').textContent = s ? s.name : '⚑ Tempo';
  if (s) return renderDetail(s);
  if (!studies.length) { app.innerHTML = '<div class="empty">Henüz etüt yok. “+ Yeni etüt” ile başlayın.</div>'; return; }
  app.innerHTML = '<div class="grid">' + studies.map(s => {
    const last = lapRows(s).at(-1);
    return `<div class="card">
      <div class="row" style="justify-content:space-between;align-items:center">
        <input class="name" data-name="${s.id}" value="${esc(s.name)}" title="Adı düzenle">
        <button class="open" data-a="open" data-id="${s.id}" title="Büyük ekran">⤢</button>
        <span class="badge ${s.status}">${durum(s)}</span>
      </div>
      <div class="time" data-t="${s.id}">${fmt(elapsed(s))}</div>
      <div class="sub">Başlangıç ${clock(s.startedAt)} · Bitiş ${clock(s.finishedAt)} · ${visibleLaps(s).length} bayrak${last ? ' · son tur ' + fmt(last.tur) : ''}</div>
      ${controls(s)}
      ${lapList(s)}
    </div>`;
  }).join('') + '</div>';
}

// 03 | 00:12,4 | 00:41,7 [not] ✕ — en yeni üstte
function lapList(s) {
  const p = n => String(n).padStart(2, '0');
  const rows = lapRows(s).reverse();
  const next = s.status === 'finished' ? '' : `<tr class="next">
    <td></td><td class="no">${p(rows.length + 1)}</td><td data-tur="${s.id}">${fmt(elapsed(s) - (rows[0]?.elapsed || 0))}</td>
    <td class="sub" data-t="${s.id}">${fmt(elapsed(s))}</td>
    <td><input data-pend="${s.id}" value="${esc(s.pendingNote)}" placeholder="sıradaki not"></td><td></td></tr>`;
  if (!rows.length && !next) return '';
  return `<table class="laps"><tbody>${next}${rows.map(r => `<tr class="${r.tag ? 'tagged' : ''}">
    <td><button class="tag ${r.tag ? 'on' : ''}" data-a="tag" data-id="${s.id}" data-lap="${r.id}">●</button></td>
    <td class="no">${p(r.sira)}</td><td>${fmt(r.tur)}</td><td class="sub">${fmt(r.elapsed)}</td>
    <td><input data-note="${r.id}" data-id="${s.id}" value="${esc(r.note)}" placeholder="not"></td>
    <td><button class="del" data-a="del" data-id="${s.id}" data-lap="${r.id}">✕</button></td></tr>`).join('')}</tbody></table>
    <input class="meaning" data-meaning="${s.id}" value="${esc(s.tagMeaning)}" placeholder="● Yeşil etiket ne anlama geliyor?">`;
}

function renderDetail(s) {
  const rows = lapRows(s).reverse();
  app.innerHTML = `<div class="card big">
      <div class="time" data-t="${s.id}">${fmt(elapsed(s))}</div>
      <div class="sub" style="text-align:center">${durum(s)} · Başlangıç ${clock(s.startedAt)} · Bitiş ${clock(s.finishedAt)}</div>
      ${controls(s)}
    </div>
    <div style="margin-top:12px">${lapList(s)}</div>
    ${rows.length ? '' : '<div class="empty">Bayrak yok.</div>'}
    <div class="row" style="margin-top:24px">
      <button data-a="rename">Yeniden adlandır</button>
      <button class="bad" data-a="remove">Etüdü sil</button>
    </div>`;
}

app.addEventListener('click', e => {
  const b = e.target.closest('[data-a]'); if (!b) return;
  const s = byId(b.dataset.id || openId), a = b.dataset.a;
  if (a === 'open') { openId = s.id; render(); scrollTo(0, 0); }
  else if (a === 'start') start(s);
  else if (a === 'pause') pause(s);
  else if (a === 'flag') flag(s);
  else if (a === 'finish') finish(s);
  else if (a === 'tag') toggleTag(s, b.dataset.lap);
  else if (a === 'del') delLap(s, b.dataset.lap);
  else if (a === 'rename') rename(s);
  else if (a === 'remove') removeStudy(s);
});
app.addEventListener('input', e => { const d = e.target.dataset;
  if (d.pend) { byId(d.pend).pendingNote = e.target.value; save(); } if (d.note) setNote(byId(d.id), d.note, e.target.value, false); });
app.addEventListener('change', e => { const d = e.target.dataset;
  if (d.name) { const s = byId(d.name), n = e.target.value.trim(); if (n && n !== s.name) { log(s, 'yeniden adlandır', s.name + ' → ' + n); s.name = n; save(); } else e.target.value = s.name; }
  if (d.meaning) { const s = byId(d.meaning); s.tagMeaning = e.target.value; log(s, 'etiket anlamı', e.target.value); save(); } if (d.note) setNote(byId(d.id), d.note, e.target.value, true); });
app.addEventListener('keydown', e => { if (e.key === 'Enter' && (e.target.dataset.note || e.target.dataset.meaning || e.target.dataset.pend || e.target.dataset.name)) e.target.blur(); });
document.getElementById('back').onclick = () => { openId = null; render(); };
document.getElementById('newStudy').onclick = newStudy;
document.getElementById('exportAll').onclick = () => exportXlsx(openId ? [byId(openId)] : studies);

// Sadece süre yazılarını günceller; not yazarken input'lar yeniden çizilmez.
setInterval(() => {
  document.querySelectorAll('[data-t]').forEach(el => { const s = byId(el.dataset.t); if (s) el.textContent = fmt(elapsed(s)); });
  document.querySelectorAll('[data-tur]').forEach(el => { const s = byId(el.dataset.tur); if (s) el.textContent = fmt(elapsed(s) - (visibleLaps(s).at(-1)?.elapsed || 0)); });
}, 100);

// Ekran kapanmasın (destekleniyorsa)
async function wake() { try { await navigator.wakeLock?.request('screen'); } catch {} }
document.addEventListener('visibilitychange', () => document.visibilityState === 'visible' && wake());
wake();

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
render();
