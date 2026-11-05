// Ядро: подстановка {{Колонка}} в .docx. Работает в браузере (и в Node с jsdom для тестов).
(function (root) {
  const PH = /\{\{\s*(.+?)\s*\}\}/g;
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  // Части документа, где может быть текст
  const isTextPart = n => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n);

  // <w:t> элементы, принадлежащие именно этому параграфу (не вложенным, напр. в текстовых рамках)
  function ownTexts(p) {
    return Array.from(p.getElementsByTagNameNS(W, 't')).filter(t => {
      let a = t.parentNode;
      while (a && !(a.localName === 'p' && a.namespaceURI === W)) a = a.parentNode;
      return a === p;
    });
  }

  function replaceInParagraph(p, data, missing, found) {
    const ts = ownTexts(p);
    if (!ts.length) return;
    const full = ts.map(t => t.textContent).join('');
    if (!full.includes('{{')) return;
    const matches = Array.from(full.matchAll(PH));
    for (const m of matches.reverse()) {
      const key = m[1];
      if (found) found.add(key);
      if (!data) continue;
      if (!(key in data)) { missing && missing.add(key); continue; }
      const start = m.index, end = m.index + m[0].length;
      let pos = 0, first = true;
      for (const t of ts) {
        const txt = t.textContent, rs = pos, re = pos + txt.length;
        pos = re;
        if (re <= start || rs >= end) continue;
        const s = Math.max(start, rs) - rs, e = Math.min(end, re) - rs;
        t.textContent = txt.slice(0, s) + (first ? data[key] : '') + txt.slice(e);
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        first = false;
      }
    }
  }

  // Заменяет метки в XML-строке части документа. data=null — только поиск меток.
  function processXml(xml, data, missing, found) {
    const doc = new root.DOMParser().parseFromString(xml, 'application/xml');
    for (const p of Array.from(doc.getElementsByTagNameNS(W, 'p'))) replaceInParagraph(p, data, missing, found);
    return new root.XMLSerializer().serializeToString(doc);
  }

  async function findPlaceholders(zipBuf) {
    const zip = await root.JSZip.loadAsync(zipBuf);
    const found = new Set();
    for (const name of Object.keys(zip.files).filter(isTextPart)) {
      processXml(await zip.file(name).async('string'), null, null, found);
    }
    return Array.from(found);
  }

  // Возвращает {blob/uint8, missing:Set}
  async function fillTemplate(zipBuf, data) {
    const zip = await root.JSZip.loadAsync(zipBuf);
    const missing = new Set();
    for (const name of Object.keys(zip.files).filter(isTextPart)) {
      const xml = await zip.file(name).async('string');
      if (!xml.includes('{')) continue;
      zip.file(name, processXml(xml, data, missing, null));
    }
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { bytes, missing };
  }

  function safeName(s) {
    s = String(s).replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/^[. ]+|[. ]+$/g, '');
    return s.slice(0, 120) || 'документ';
  }

  // Excel -> {header, records}
  function readExcel(XLSX, buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const p2 = n => String(n).padStart(2, '0');
    for (const k of Object.keys(ws)) { // даты — всегда ДД.ММ.ГГГГ, независимо от формата ячейки
      const c = ws[k];
      if (k[0] !== '!' && c && c.t === 'd' && c.v instanceof Date)
      { const d = new Date(c.v.getTime() + 12 * 3600e3); // +12ч: гасит сдвиг часовых поясов у SheetJS
        c.w = `${p2(d.getDate())}.${p2(d.getMonth() + 1)}.${d.getFullYear()}`; }
    }
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: 'dd.mm.yyyy', defval: '', blankrows: false });
    if (!rows.length) return { header: [], records: [] };
    const header = rows[0].map(h => String(h).trim());
    const records = rows.slice(1)
      .filter(r => r.some(c => String(c).trim() !== ''))
      .map(r => { const o = {}; header.forEach((h, i) => { if (h) o[h] = String(r[i] ?? '').trim(); }); return o; });
    return { header: header.filter(Boolean), records };
  }

  const api = { findPlaceholders, fillTemplate, safeName, readExcel };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Core = api;
})(typeof window !== 'undefined' ? window : globalThis);
