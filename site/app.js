const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const msg = (el, cls, html) => { el.innerHTML = `<div class="msg ${cls}">${html}</div>`; };

let state = null; // {header, records, docxBuf, placeholders}

if (!window.showDirectoryPicker) $('dir').classList.add('hidden');

async function tryLoad() {
  const x = $('fx').files[0], d = $('fd').files[0];
  $('dx').classList.toggle('has', !!x); $('dd').classList.toggle('has', !!d);
  if (x) $('dx').querySelector('span').textContent = x.name;
  if (d) $('dd').querySelector('span').textContent = d.name;
  if (!x || !d) return;
  $('upmsg').innerHTML = '';
  try {
    const { header, records } = Core.readExcel(XLSX, await x.arrayBuffer());
    const docxBuf = await d.arrayBuffer();
    const placeholders = await Core.findPlaceholders(docxBuf);
    state = { header, records, docxBuf, placeholders };
    $('info').textContent = `Найдено строк: ${records.length}. Колонки: ${header.join(', ')}`;
    $('prev').innerHTML = '<tr>' + header.map(c => `<th>${esc(c)}</th>`).join('') + '</tr>' +
      records.slice(0, 5).map(r => '<tr>' + header.map(c => `<td>${esc(r[c] ?? '')}</td>`).join('') + '</tr>').join('');
    const unmatched = placeholders.filter(p => !header.includes(p));
    $('chips').innerHTML = placeholders.length
      ? placeholders.map(p => `<span class="${unmatched.includes(p) ? 'bad' : ''}">${esc(p)}</span>`).join('')
      : '<em>Меток &#123;&#123;...&#125;&#125; в шаблоне не найдено</em>';
    if (unmatched.length) msg($('upmsg'), 'warn', 'Нет колонки в Excel для меток: ' + unmatched.map(esc).join(', '));
    $('namecol').innerHTML = header.map(c => `<option>${esc(c)}</option>`).join('');
    const fio = header.find(c => /фио|ф\.и\.о|имя/i.test(c)); if (fio) $('namecol').value = fio;
    $('step2').classList.remove('hidden'); $('step3').classList.remove('hidden');
    $('result').innerHTML = '';
  } catch (e) {
    msg($('upmsg'), 'err', 'Не удалось прочитать файлы: ' + esc(e.message));
  }
}
$('fx').onchange = $('fd').onchange = tryLoad;

// Генерирует все документы: [{name, bytes}], warnings
async function generate() {
  const col = $('namecol').value, used = {}, files = [], warnings = [];
  for (let i = 0; i < state.records.length; i++) {
    const rec = state.records[i], row = i + 2;
    const base = rec[col] ? Core.safeName(rec[col]) : `строка ${row}`;
    const n = (used[base.toLowerCase()] || 0) + 1; used[base.toLowerCase()] = n;
    const name = n === 1 ? `${base}.docx` : `${base}_${n}.docx`;
    const { bytes, missing } = await Core.fillTemplate(state.docxBuf, rec);
    files.push({ name, bytes });
    const empty = Object.keys(rec).filter(k => rec[k] === '');
    if (empty.length) warnings.push(`Строка ${row} (${name}): пустые ячейки — ${empty.join(', ')}`);
    if (missing.size) warnings.push(`${name}: нет колонки для меток — ${[...missing].join(', ')}`);
  }
  return { files, warnings };
}

function report(count, extra, warnings) {
  let h = `<div class="msg ok">Создано документов: <b>${count}</b>. ${extra}</div>`;
  if (warnings.length) h += `<div class="msg warn"><b>Предупреждения:</b><ul>${warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></div>`;
  $('result').innerHTML = h;
}

async function run(btn, action) {
  const label = btn.textContent; btn.disabled = true; btn.textContent = 'Формирую…';
  try { await action(); }
  catch (e) { if (e.name !== 'AbortError') msg($('result'), 'err', esc(e.message)); }
  finally { btn.disabled = false; btn.textContent = label; }
}

$('zip').onclick = () => run($('zip'), async () => {
  const { files, warnings } = await generate();
  const zip = new JSZip();
  files.forEach(f => zip.file(f.name, f.bytes));
  const blob = await zip.generateAsync({ type: 'blob' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = 'documents.zip'; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  report(files.length, 'Архив скачан в папку загрузок.', warnings);
});

$('dir').onclick = () => run($('dir'), async () => {
  const dir = await showDirectoryPicker({ mode: 'readwrite' });
  const { files, warnings } = await generate();
  for (const f of files) {
    const h = await dir.getFileHandle(f.name, { create: true });
    const w = await h.createWritable(); await w.write(f.bytes); await w.close();
  }
  report(files.length, `Сохранено в папку «${esc(dir.name)}».`, warnings);
});
