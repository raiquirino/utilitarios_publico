function setupDropZone(zoneId, inputId, fileNameId) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  const nameEl = document.getElementById(fileNameId);
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', e => { e.preventDefault(); zone.classList.remove('drag-over'); });
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      showFileNames(nameEl, e.dataTransfer.files, zone);
      input.dispatchEvent(new Event('change'));
    }
  });
  input.addEventListener('change', () => showFileNames(nameEl, input.files, zone));
}
function showFileNames(nameEl, files, zone) {
  if (!files || !files.length) return;
  nameEl.textContent = files.length === 1 ? files[0].name : files.length + ' arquivos';
  zone.classList.add('has-file');
}

setupDropZone('dropZoneOFX', 'ofxInput', 'fileNameOFX');
setupDropZone('dropZoneExcel', 'csvInput', 'fileNameExcel');

// ========== Colar (Ctrl+V) nas zonas ==========
function parsePastedTable(text) {
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  if (!text) return null;
  const lines = text.split('\n').filter(l => l.trim() !== '');
  if (!lines.length) return null;
  const first = lines[0];
  let delim = '\t';
  if (!first.includes('\t')) {
    const semi = (first.match(/;/g) || []).length;
    const comma = (first.match(/,/g) || []).length;
    delim = semi > comma ? ';' : ',';
  }
  function splitLine(line) {
    if (delim === '\t') return line.split('\t');
    const cells = [];
    let cur = '', inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') { cur += '"'; i++; }
        else inQuotes = !inQuotes;
      } else if (ch === delim && !inQuotes) {
        cells.push(cur); cur = '';
      } else cur += ch;
    }
    cells.push(cur);
    return cells;
  }
  const rows = lines.map(splitLine);
  const maxCols = Math.max(...rows.map(r => r.length));
  rows.forEach(r => { while (r.length < maxCols) r.push(''); });
  return rows;
}

function looksLikeOFX(text) {
  const t = text.trim().toUpperCase();
  return t.includes('OFXHEADER') || t.includes('<OFX>') || t.includes('<STMTTRN>') || t.includes('<BANKTRANLIST>');
}

function assignFilesToInput(inputId, files) {
  const input = document.getElementById(inputId);
  const dt = new DataTransfer();
  files.forEach(f => dt.items.add(f));
  input.files = dt.files;
  input.dispatchEvent(new Event('change'));
}

function handlePasteOnOFX(text) {
  if (!looksLikeOFX(text)) return false;
  const blob = new Blob([text], { type: 'text/plain' });
  const file = new File([blob], 'colado.ofx', { type: 'text/plain' });
  assignFilesToInput('ofxInput', [file]);
  return true;
}

function handlePasteOnExcel(text) {
  if (looksLikeOFX(text)) return false;
  const rows = parsePastedTable(text);
  if (!rows || rows.length < 1) return false;
  // Heurística: precisa parecer tabela (múltiplas colunas ou várias linhas)
  if (rows.length < 2 && rows[0].length < 2) return false;

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Extrato');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], 'colado.xlsx', { type: blob.type });
  assignFilesToInput('csvInput', [file]);
  return true;
}

function onGlobalPaste(e) {
  const text = e.clipboardData && e.clipboardData.getData('text/plain');
  if (!text || !text.trim()) return;

  const tag = document.activeElement && document.activeElement.tagName;
  // Não roubar paste de campos de texto do formulário
  if (tag === 'INPUT' || tag === 'TEXTAREA') {
    const id = document.activeElement.id || '';
    if (id.startsWith('fld') || id === 'fldBankSearch') return;
  }

  // Prioridade: OFX se parecer OFX; senão tabela Excel
  if (looksLikeOFX(text)) {
    e.preventDefault();
    handlePasteOnOFX(text);
    return;
  }
  if (text.includes('\t') || text.split(/\r?\n/).filter(l => l.trim()).length >= 2) {
    e.preventDefault();
    handlePasteOnExcel(text);
  }
}

document.addEventListener('paste', onGlobalPaste);

// Clique na zona + Ctrl+V também funciona via paste global;
// torna as zonas focáveis para feedback visual
['dropZoneOFX', 'dropZoneExcel'].forEach(id => {
  const zone = document.getElementById(id);
  if (!zone) return;
  zone.setAttribute('tabindex', '0');
  zone.addEventListener('click', () => zone.focus());
});

// Códigos COMPE (principais bancos BR)
const BANCOS = {
  "001": "BB - Banco do Brasil S.A.",
  "003": "Banco da Amazônia S.A.",
  "004": "Banco do Nordeste do Brasil S.A.",
  "007": "BNDES",
  "012": "Banco Inbursa S.A.",
  "021": "Banestes S.A.",
  "025": "Banco Alfa S.A.",
  "029": "Banco Itaú Consignado S.A.",
  "033": "Banco Santander (Brasil) S.A.",
  "036": "Banco Bradesco BBI S.A.",
  "037": "Banpará S.A.",
  "041": "Banrisul S.A.",
  "047": "Banese S.A.",
  "062": "Hipercard Banco Múltiplo S.A.",
  "063": "Banco Bradescard S.A.",
  "069": "Banco Crefisa S.A.",
  "070": "BRB - Banco de Brasília S.A.",
  "074": "Banco J. Safra S.A.",
  "077": "Banco Inter S.A.",
  "082": "Banco Topázio S.A.",
  "083": "Banco da China Brasil S.A.",
  "102": "XP Investimentos S.A.",
  "104": "Caixa Econômica Federal",
  "120": "Banco Rodobens S.A.",
  "121": "Banco Agibank S.A.",
  "197": "Stone Pagamentos S.A.",
  "208": "Banco BTG Pactual S.A.",
  "212": "Banco Original S.A.",
  "218": "Banco BS2 S.A.",
  "237": "Banco Bradesco S.A.",
  "246": "Banco ABC Brasil S.A.",
  "260": "Nu Pagamentos S.A. (Nubank)",
  "280": "Will Bank",
  "290": "PagSeguro Internet S.A. (PagBank)",
  "318": "Banco BMG S.A.",
  "323": "Mercado Pago",
  "335": "Banco Digio S.A.",
  "336": "Banco C6 S.A. (C6 Bank)",
  "341": "Itaú Unibanco S.A.",
  "348": "Banco XP S.A.",
  "380": "PicPay",
  "389": "Banco Mercantil do Brasil S.A.",
  "422": "Banco Safra S.A.",
  "536": "Neon Pagamentos S.A.",
  "623": "Banco Pan S.A.",
  "637": "Banco Sofisa S.A.",
  "655": "Banco Votorantim S.A. (BV)",
  "707": "Banco Daycoval S.A.",
  "745": "Banco Citibank S.A.",
  "748": "Banco Cooperativo Sicredi S.A.",
  "756": "Banco Cooperativo Sicoob S.A. (Bancoob)"
};

function normalizeBankCode(code) {
  if (!code) return "";
  const c = String(code).replace(/\D/g, "");
  if (!c) return "";
  return c.padStart(3, "0").slice(-3);
}

function bankNameFromCode(code) {
  const c = normalizeBankCode(code);
  return BANCOS[c] || "";
}

function searchBanks(query) {
  const q = String(query || "").toLowerCase().trim();
  if (!q) return [];
  const results = [];
  for (const [code, name] of Object.entries(BANCOS)) {
    if (code.includes(q) || name.toLowerCase().includes(q)) {
      results.push({ code, name });
    }
  }
  return results.slice(0, 12);
}

function lastDayOfMonth(year, month) {
  return new Date(year, month, 0).getDate();
}
function parseBRDate(str) {
  if (!str) return null;
  const s = String(str).trim();
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return { d: +m[1], mo: +m[2], y: +m[3] };
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return { d: +m[3], mo: +m[2], y: +m[1] };
  return null;
}
function periodFromTransactions(transactions) {
  const parsed = [];
  for (const t of transactions) {
    const p = parseBRDate(t.date);
    if (p && p.y > 1900 && p.mo >= 1 && p.mo <= 12) parsed.push(p);
  }
  if (!parsed.length) return null;
  parsed.sort((a, b) => (a.y - b.y) || (a.mo - b.mo) || (a.d - b.d));
  const first = parsed[0];
  const last = parsed[parsed.length - 1];
  const start = '01/' + String(first.mo).padStart(2, '0') + '/' + first.y;
  const endDay = lastDayOfMonth(last.y, last.mo);
  const end = String(endDay).padStart(2, '0') + '/' + String(last.mo).padStart(2, '0') + '/' + last.y;
  return { dtStart: start, dtEnd: end };
}
function applyBankCodeToForm(code) {
  const c = normalizeBankCode(code);
  if (!c) return;
  document.getElementById("fldBankId").value = c;
  const name = bankNameFromCode(c);
  if (name) document.getElementById("fldOrg").value = name;
  const hint = document.getElementById("bankHint");
  if (hint) hint.textContent = name ? ("→ " + name) : "Código não encontrado na lista";
}

// Busca por nome → preenche ID
document.addEventListener("DOMContentLoaded", () => {});


function formatDate(raw) {
  if (!raw) return '';
  let clean = String(raw).replace(/\[.*$/, '').replace(/\s+/g, '').trim();
  if (clean.length > 8) clean = clean.substring(0, 8);
  if (clean.length < 8) return String(raw).trim();
  return clean.substring(6, 8) + '/' + clean.substring(4, 6) + '/' + clean.substring(0, 4);
}
function getTag(content, tag) {
  const m = content.match(new RegExp('<' + tag + '>([^<\\n\\r]*)', 'i'));
  return m ? m[1].trim() : '';
}
function parseOFXHeader(content) {
  return {
    org: getTag(content, 'ORG'),
    fid: getTag(content, 'FID'),
    bankId: getTag(content, 'BANKID'),
    branchId: getTag(content, 'BRANCHID'),
    acctId: getTag(content, 'ACCTID'),
    acctType: getTag(content, 'ACCTTYPE') || 'CHECKING',
    curDef: getTag(content, 'CURDEF') || 'BRL',
    dtStart: formatDate(getTag(content, 'DTSTART')),
    dtEnd: formatDate(getTag(content, 'DTEND')),
    ledgerBal: getTag(content, 'BALAMT'),
    ledgerDate: formatDate(getTag(content, 'DTASOF')),
    language: getTag(content, 'LANGUAGE') || 'POR'
  };
}
function parseOFXTransactions(content) {
  let matches = content.match(/<STMTTRN>[\s\S]*?<\/STMTTRN>/gi);
  if (!matches || !matches.length) {
    matches = content.match(/<STMTTRN>[\s\S]*?(?=<STMTTRN>|<\/BANKTRANLIST>|<LEDGERBAL>|<AVAILBAL>|$)/gi) || [];
  }
  return matches.map(trn => ({
    date: formatDate(getTag(trn, 'DTPOSTED')),
    amount: getTag(trn, 'TRNAMT'),
    memo: getTag(trn, 'MEMO') || getTag(trn, 'NAME') || '',
    type: getTag(trn, 'TRNTYPE'),
    fitid: getTag(trn, 'FITID'),
    checknum: getTag(trn, 'CHECKNUM')
  })).filter(t => t.date || t.amount);
}
function formatMoney(v) {
  const n = parseFloat(String(v).replace(',', '.'));
  if (isNaN(n)) return v || '';
  return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function renderFields(container, items) {
  container.innerHTML = items.map(i =>
    '<div class="pf"><div class="lbl">' + i.lbl + '</div><div class="val">' + (i.val || '—') + '</div></div>'
  ).join('');
}
function showPreview(which) {
  document.getElementById('previewPlaceholder').classList.add('hidden');
  document.getElementById('previewOFX').classList.toggle('visible', which === 'ofx');
  document.getElementById('previewExcel').classList.toggle('visible', which === 'excel');
}

document.getElementById('ofxInput').addEventListener('change', () => {
  const files = document.getElementById('ofxInput').files;
  if (!files.length) {
    document.getElementById('previewOFX').classList.remove('visible');
    document.getElementById('previewPlaceholder').classList.remove('hidden');
    return;
  }
  const file = files[0];
  const reader = new FileReader();
  reader.onload = () => {
    try {
      let content = reader.result;
      if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
      const header = parseOFXHeader(content);
      const transactions = parseOFXTransactions(content);
      document.getElementById('previewOFXFile').textContent =
        files.length === 1 ? file.name : file.name + ' (+' + (files.length - 1) + ')';
      if (!header.org && header.bankId) header.org = bankNameFromCode(header.bankId);
      const periodOFX = periodFromTransactions(transactions);
      if (periodOFX) {
        header.dtStart = periodOFX.dtStart;
        header.dtEnd = periodOFX.dtEnd;
      }
      renderFields(document.getElementById('previewOFXMeta'), [
        { lbl: 'Instituição (ORG)', val: header.org },
        { lbl: 'Banco (BANKID)', val: header.bankId },
        { lbl: 'Agência (BRANCHID)', val: header.branchId },
        { lbl: 'Conta (ACCTID)', val: header.acctId },
        { lbl: 'Período início (DTSTART)', val: header.dtStart },
        { lbl: 'Período fim (DTEND)', val: header.dtEnd }
      ]);
      if (header.bankId) {
        document.getElementById('fldBankId').value = normalizeBankCode(header.bankId);
        const resolved = bankNameFromCode(header.bankId);
        if (resolved && !header.org) header.org = resolved;
        const hint = document.getElementById('bankHint');
        if (hint) hint.textContent = resolved ? ('→ ' + resolved) : '';
      }
      if (header.branchId) document.getElementById('fldBranchId').value = header.branchId;
      if (header.acctId) document.getElementById('fldAcctId').value = header.acctId;
      if (header.org) document.getElementById('fldOrg').value = header.org;
      else if (header.bankId && bankNameFromCode(header.bankId)) {
        document.getElementById('fldOrg').value = bankNameFromCode(header.bankId);
        header.org = bankNameFromCode(header.bankId);
      }
      if (header.dtStart) document.getElementById('fldDtStart').value = header.dtStart;
      if (header.dtEnd) document.getElementById('fldDtEnd').value = header.dtEnd;
      document.getElementById('previewOFXBody').innerHTML = transactions.map(t => {
        const n = parseFloat(String(t.amount).replace(',', '.'));
        const cls = n < 0 ? 'neg' : 'pos';
        return '<tr><td>' + (t.date || '') + '</td><td title="' +
          (t.memo || '').replace(/"/g, '&quot;') + '">' + (t.memo || '') +
          '</td><td class="' + cls + '">' + formatMoney(t.amount) + '</td></tr>';
      }).join('');
      document.getElementById('previewOFXCount').textContent =
        transactions.length + ' transação(ões)' +
        (transactions.length === 0 ? ' — verifique se o arquivo é OFX válido' : '') +
        (files.length > 1 ? ' · prévia do 1º arquivo' : '');
      showPreview('ofx');
    } catch (err) {
      alert('Erro ao ler OFX: ' + err.message);
      console.error(err);
    }
  };
  reader.onerror = () => alert('Não foi possível ler o arquivo.');
  reader.readAsText(file, 'ISO-8859-1');
});

function readHeaderFromSheet(workbook) {
  let rows = null;
  const named = workbook.SheetNames.find(n => /cabeçalho|cabecalho|header|info/i.test(n));
  if (named) {
    rows = XLSX.utils.sheet_to_json(workbook.Sheets[named], { header: 1 });
  } else {
    for (const sn of workbook.SheetNames) {
      const all = XLSX.utils.sheet_to_json(workbook.Sheets[sn], { header: 1 });
      if (all.length && /campo/i.test(String(all[0][0] || ''))) { rows = all; break; }
    }
  }
  if (!rows) return null;
  const map = {};
  for (const r of rows) {
    if (!r || !r[0]) continue;
    const key = String(r[0]).toLowerCase().trim();
    if (/^data$/i.test(key) || key === 'tipo' || key === 'memo' || key === 'nome') break;
    if (r[1] !== undefined && r[1] !== '') map[key] = String(r[1]).trim();
  }
  return {
    org: map['instituição (org)'] || map['org'] || map['instituição'] || '',
    fid: map['fid'] || '',
    bankId: map['banco (bankid)'] || map['bankid'] || map['banco'] || '',
    branchId: map['agência (branchid)'] || map['agencia (branchid)'] || map['branchid'] || map['agência'] || map['agencia'] || '',
    acctId: map['conta (acctid)'] || map['acctid'] || map['conta'] || '',
    acctType: map['tipo de conta (accttype)'] || map['accttype'] || map['tipo de conta'] || 'CHECKING',
    curDef: map['moeda (curdef)'] || map['curdef'] || map['moeda'] || 'BRL',
    dtStart: map['período início (dtstart)'] || map['dtstart'] || map['período início'] || '',
    dtEnd: map['período fim (dtend)'] || map['dtend'] || map['período fim'] || '',
    ledgerBal: map['saldo contábil (ledgerbal)'] || map['ledgerbal'] || map['saldo'] || '',
    ledgerDate: map['data do saldo (dtasof)'] || map['dtasof'] || map['data do saldo'] || '',
    language: 'POR'
  };
}

document.getElementById('csvInput').addEventListener('change', () => {
  const files = document.getElementById('csvInput').files;
  if (!files.length) {
    document.getElementById('previewExcel').classList.remove('visible');
    document.getElementById('previewPlaceholder').classList.remove('hidden');
    return;
  }
  const file = files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const header = readHeaderFromSheet(workbook);
      if (header) {
        if (!header.org && header.bankId) header.org = bankNameFromCode(header.bankId);
        renderFields(document.getElementById('previewExcelMeta'), [
          { lbl: 'Instituição (ORG)', val: header.org },
          { lbl: 'Banco (BANKID)', val: header.bankId },
          { lbl: 'Agência (BRANCHID)', val: header.branchId },
          { lbl: 'Conta (ACCTID)', val: header.acctId },
          { lbl: 'Período início (DTSTART)', val: header.dtStart },
          { lbl: 'Período fim (DTEND)', val: header.dtEnd }
        ]);
        if (header.bankId) document.getElementById('fldBankId').value = header.bankId;
        if (header.branchId) document.getElementById('fldBranchId').value = header.branchId;
        if (header.acctId) document.getElementById('fldAcctId').value = header.acctId;
        if (header.org) document.getElementById('fldOrg').value = header.org;
        if (header.dtStart) document.getElementById('fldDtStart').value = header.dtStart;
        if (header.dtEnd) document.getElementById('fldDtEnd').value = header.dtEnd;
      } else {
        document.getElementById('previewExcelMeta').innerHTML = '';
      }
      const sheetName = workbook.SheetNames.find(n => /extrato|transa|dados|sheet/i.test(n))
        || workbook.SheetNames.find(n => !/cabeçalho|cabecalho|header|info/i.test(n))
        || workbook.SheetNames[0];
      const allRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      let headerIdx = allRows.findIndex(r => r && /^data$/i.test(String(r[0] || '').trim()));
      if (headerIdx < 0) headerIdx = allRows.findIndex(r => r && r.some(c => /^data$/i.test(String(c || '').trim())));
      if (headerIdx < 0) headerIdx = 0;
      const rows = allRows.slice(headerIdx);
      document.getElementById('previewExcelFile').textContent =
        (files.length === 1 ? file.name : file.name + ' (+' + (files.length - 1) + ')') + ' · ' + sheetName;
      if (!rows.length) {
        document.getElementById('previewExcelHead').innerHTML = '';
        document.getElementById('previewExcelBody').innerHTML = '';
        document.getElementById('previewExcelCount').textContent = 'Nenhum dado';
        showPreview('excel');
        return;
      }
      const head = (rows[0] || []).map(c => String(c || ''));
      const hl = head.map(c => c.toLowerCase());
      let idxDate = hl.findIndex(c => /data|date/.test(c));
      let idxMemo = hl.findIndex(c => /memo|nome|descri|hist/.test(c));
      let idxAmount = hl.findIndex(c => /valor|amount/.test(c));
      if (idxDate < 0) idxDate = 0;
      if (idxMemo < 0) idxMemo = 1;
      if (idxAmount < 0) idxAmount = 2;
      const cols = [idxDate, idxMemo, idxAmount];
      document.getElementById('previewExcelHead').innerHTML = '<tr><th>Data</th><th>Memo</th><th>Valor</th></tr>';
      document.getElementById('previewExcelBody').innerHTML = rows.slice(1, 101).map(row => {
        const cells = cols.map((ci, i) => {
          let v = (row || [])[ci] ?? '';
          if (typeof v === 'number' && i !== 0) v = formatMoney(v);
          if (typeof v === 'number' && i === 0) {
            // serial Excel → DD/MM/YYYY para exibição
            const ofx = formatDateToOFX(v);
            v = ofx ? (ofx.substring(6,8) + '/' + ofx.substring(4,6) + '/' + ofx.substring(0,4)) : v;
          }
          let cls = '';
          if (i === 2) {
            const n = parseFloat(String(v).replace(/\./g, '').replace(',', '.'));
            if (!isNaN(n)) cls = n < 0 ? 'neg' : 'pos';
          }
          return '<td class="' + cls + '" title="' + String(v).replace(/"/g, '&quot;') + '">' + v + '</td>';
        });
        return '<tr>' + cells.join('') + '</tr>';
      }).join('');

      // Calcula período pelas datas das linhas e preenche formulário + meta
      const txForPeriod = [];
      for (let i = 1; i < rows.length; i++) {
        const raw = (rows[i] || [])[idxDate];
        if (raw === undefined || raw === '') continue;
        let dstr = '';
        if (typeof raw === 'number') {
          const ofx = formatDateToOFX(raw);
          if (ofx) dstr = ofx.substring(6,8) + '/' + ofx.substring(4,6) + '/' + ofx.substring(0,4);
        } else {
          dstr = String(raw).trim();
          // se já YYYYMMDD
          if (/^\d{8}$/.test(dstr)) dstr = dstr.substring(6,8) + '/' + dstr.substring(4,6) + '/' + dstr.substring(0,4);
        }
        if (dstr) txForPeriod.push({ date: dstr });
      }
      const periodX = periodFromTransactions(txForPeriod);
      let h = header || {};
      if (periodX) {
        h.dtStart = periodX.dtStart;
        h.dtEnd = periodX.dtEnd;
        document.getElementById('fldDtStart').value = periodX.dtStart;
        document.getElementById('fldDtEnd').value = periodX.dtEnd;
      }
      if (!h.org && h.bankId) h.org = bankNameFromCode(h.bankId);
      renderFields(document.getElementById('previewExcelMeta'), [
        { lbl: 'Instituição (ORG)', val: h.org },
        { lbl: 'Banco (BANKID)', val: h.bankId },
        { lbl: 'Agência (BRANCHID)', val: h.branchId },
        { lbl: 'Conta (ACCTID)', val: h.acctId },
        { lbl: 'Período início (DTSTART)', val: h.dtStart },
        { lbl: 'Período fim (DTEND)', val: h.dtEnd }
      ]);

      document.getElementById('previewExcelCount').textContent =
        (rows.length - 1) + ' linha(s)' + (rows.length - 1 > 100 ? ' (mostrando 100)' : '');
      showPreview('excel');
    } catch (err) {
      alert('Erro ao ler Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

function toExcel(transactions, header) {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['Campo', 'Valor'],
    ['Instituição (ORG)', header.org || ''],
    ['Banco (BANKID)', header.bankId || ''],
    ['Agência (BRANCHID)', header.branchId || ''],
    ['Conta (ACCTID)', header.acctId || ''],
    ['Período início (DTSTART)', header.dtStart || ''],
    ['Período fim (DTEND)', header.dtEnd || ''],
    [],
    ['Data', 'Memo', 'Valor'],
    ...transactions.map(t => [
      t.date,
      t.memo || '',
      parseFloat(String(t.amount).replace(',', '.')) || 0
    ])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 28 }, { wch: 50 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, ws, 'Extrato');
  return new Blob([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/octet-stream' });
}

document.getElementById('btnConvertOFX').addEventListener('click', () => {
  const files = document.getElementById('ofxInput').files;
  const progress = document.getElementById('progressOFX');
  const bar = document.getElementById('progressBarOFX');
  const links = document.getElementById('downloadExcelLinks');
  progress.textContent = ''; bar.style.width = '0%'; links.innerHTML = '';
  if (!files.length) { alert('Selecione um ou mais arquivos OFX.'); return; }
  let processed = 0;
  [...files].forEach(file => {
    const reader = new FileReader();
    reader.onload = () => {
      let content = reader.result;
      if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
      const header = parseOFXHeader(content);
      if (!header.org && header.bankId) header.org = bankNameFromCode(header.bankId);
      const transactions = parseOFXTransactions(content);
      const periodExp = periodFromTransactions(transactions);
      if (periodExp) {
        header.dtStart = periodExp.dtStart;
        header.dtEnd = periodExp.dtEnd;
      }
      const blob = toExcel(transactions, header);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = file.name.replace(/\.ofx$/i, '') + '.xlsx';
      a.textContent = '⬇ ' + a.download;
      links.appendChild(a);
      processed++;
      const pct = Math.floor(processed / files.length * 100);
      progress.textContent = pct + '%'; bar.style.width = pct + '%';
      if (processed === files.length) { progress.textContent = 'Concluído!'; bar.style.width = '100%'; }
    };
    reader.readAsText(file, 'ISO-8859-1');
  });
});

function formatDateToOFX(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (typeof raw === 'number' && !isNaN(raw)) {
    const base = Date.UTC(1899, 11, 30);
    const d = new Date(base + raw * 86400000);
    return d.getUTCFullYear() + String(d.getUTCMonth() + 1).padStart(2, '0') + String(d.getUTCDate()).padStart(2, '0');
  }
  const str = String(raw).trim();
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts;
    if (c.length === 4) return c + b.padStart(2, '0') + a.padStart(2, '0');
    if (a.length === 4) return a + b.padStart(2, '0') + c.padStart(2, '0');
  }
  if (/^\d{8}$/.test(str)) {
    if (parseInt(str.substring(0, 4)) > 1900) return str;
    return str.substring(4, 8) + str.substring(2, 4) + str.substring(0, 2);
  }
  return '';
}

function generateOFX(transactions, header) {
  const now = new Date();
  const dtserver = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
  let dtStart = header.dtStart ? formatDateToOFX(header.dtStart) : '';
  let dtEnd = header.dtEnd ? formatDateToOFX(header.dtEnd) : '';
  if (!dtStart && transactions.length) {
    const dates = transactions.map(t => t.date).filter(Boolean).sort();
    dtStart = dates[0] || dtserver;
    dtEnd = dates[dates.length - 1] || dtserver;
  }
  if (!dtStart) dtStart = dtserver;
  if (!dtEnd) dtEnd = dtserver;
  const bankId = header.bankId || '000';
  const branchId = header.branchId || '';
  const branchLine = branchId ? '\n    <BRANCHID>' + branchId : '';
  const acctId = header.acctId || '000000';
  const acctType = header.acctType || 'CHECKING';
  const curDef = header.curDef || 'BRL';
  const org = header.org || 'Banco';
  const fid = header.fid || bankId;
  const balAmt = header.ledgerBal || '0.00';
  const balDate = header.ledgerDate ? formatDateToOFX(header.ledgerDate) : dtserver;
  let body = '';
  transactions.forEach((t, i) => {
    const trnType = t.type || (parseFloat(t.amount) >= 0 ? 'DEP' : 'DEBIT');
    const fitid = t.fitid || (t.date + (i + 1));
    const checknum = t.checknum ? '\n            <CHECKNUM>' + t.checknum : '';
    body += '\n          <STMTTRN>\n            <TRNTYPE>' + trnType +
      '\n            <DTPOSTED>' + t.date +
      '\n            <TRNAMT>' + t.amount +
      '\n            <FITID>' + fitid + checknum +
      '\n            <MEMO>' + (t.memo || '') +
      '\n          </STMTTRN>';
  });
  return 'OFXHEADER:100\nDATA:OFXSGML\nVERSION:102\nSECURITY:NONE\nENCODING:USASCII\nCHARSET:1252\nCOMPRESSION:NONE\nOLDFILEUID:NONE\nNEWFILEUID:NONE\n' +
    '<OFX>\n<SIGNONMSGSRSV1>\n <SONRS>\n  <STATUS>\n   <CODE>0\n   <SEVERITY>INFO\n  </STATUS>\n  <DTSERVER>' + dtserver +
    '\n  <LANGUAGE>POR\n  <FI>\n   <ORG>' + org + '\n   <FID>' + fid +
    '\n  </FI>\n </SONRS>\n</SIGNONMSGSRSV1>\n<BANKMSGSRSV1>\n <STMTTRNRS>\n  <TRNUID>0\n  <STATUS>\n   <CODE>0\n   <SEVERITY>INFO\n  </STATUS>\n  <STMTRS>\n   <CURDEF>' + curDef +
    '\n   <BANKACCTFROM>\n    <BANKID>' + bankId + branchLine +
    '\n    <ACCTID>' + acctId +
    '\n    <ACCTTYPE>' + acctType +
    '\n   </BANKACCTFROM>\n   <BANKTRANLIST>\n    <DTSTART>' + dtStart +
    '\n    <DTEND>' + dtEnd + body +
    '\n   </BANKTRANLIST>\n   <LEDGERBAL>\n    <BALAMT>' + balAmt +
    '\n    <DTASOF>' + balDate +
    '\n   </LEDGERBAL>\n  </STMTRS>\n </STMTTRNRS>\n</BANKMSGSRSV1>\n</OFX>';
}

document.getElementById('btnConvertCSV').addEventListener('click', () => {
  const files = document.getElementById('csvInput').files;
  const progress = document.getElementById('progressCSV');
  const bar = document.getElementById('progressBarCSV');
  const links = document.getElementById('downloadOFXLinks');
  progress.textContent = ''; bar.style.width = '0%'; links.innerHTML = '';
  if (!files.length) { alert('Selecione uma ou mais planilhas Excel.'); return; }
  let processed = 0;
  [...files].forEach(file => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      let header = readHeaderFromSheet(workbook) || {};
      const formBank = document.getElementById('fldBankId').value.trim();
      const formBranch = document.getElementById('fldBranchId').value.trim();
      const formAcct = document.getElementById('fldAcctId').value.trim();
      const formOrg = document.getElementById('fldOrg').value.trim();
      const formDtStart = document.getElementById('fldDtStart').value.trim();
      const formDtEnd = document.getElementById('fldDtEnd').value.trim();
      const resolvedBankId = formBank || header.bankId || '';
      header = {
        org: formOrg || header.org || bankNameFromCode(resolvedBankId) || '',
        fid: header.fid || formBank || header.bankId || '001',
        bankId: resolvedBankId,
        branchId: formBranch || header.branchId || '',
        acctId: formAcct || header.acctId || '',
        acctType: header.acctType || 'CHECKING',
        curDef: header.curDef || 'BRL',
        dtStart: formDtStart || header.dtStart || '',
        dtEnd: formDtEnd || header.dtEnd || '',
        ledgerBal: header.ledgerBal || '0.00',
        ledgerDate: header.ledgerDate || '',
        language: 'POR'
      };
      const sheetName = workbook.SheetNames.find(n => /extrato|transa|dados|sheet/i.test(n))
        || workbook.SheetNames.find(n => !/cabeçalho|cabecalho|header|info/i.test(n))
        || workbook.SheetNames[0];
      const allRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      let headerIdx = allRows.findIndex(r => r && /^data$/i.test(String(r[0] || '').trim()));
      if (headerIdx < 0) headerIdx = allRows.findIndex(r => r && r.some(c => /^data$/i.test(String(c || '').trim())));
      if (headerIdx < 0) headerIdx = 0;
      const rows = allRows.slice(headerIdx);
      const headerRow = (rows[0] || []).map(c => String(c || '').toLowerCase().trim());
      const idxDate = headerRow.findIndex(c => /data|date/.test(c));
      const idxType = headerRow.findIndex(c => /tipo|type|trntype/.test(c));
      const idxMemo = headerRow.findIndex(c => /memo|nome|descri|hist|name/.test(c));
      const idxAmount = headerRow.findIndex(c => /valor|amount|trnamt/.test(c));
      const idxFitid = headerRow.findIndex(c => /fitid/.test(c));
      const idxCheck = headerRow.findIndex(c => /checknum|check|cheque/.test(c));
      const colDate = idxDate >= 0 ? idxDate : 0;
      const colMemo = idxMemo >= 0 ? idxMemo : 1;
      const colAmount = idxAmount >= 0 ? idxAmount : 2;
      const transactions = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row.length) continue;
        if (row[colDate] === undefined || row[colDate] === '' || row[colAmount] === undefined || row[colAmount] === '') continue;
        transactions.push({
          date: formatDateToOFX(row[colDate]),
          amount: String(row[colAmount]).replace(',', '.').replace(/\s/g, '').trim(),
          memo: String(row[colMemo] || '').trim(),
          type: idxType >= 0 ? String(row[idxType] || '').trim() : '',
          fitid: idxFitid >= 0 ? String(row[idxFitid] || '').trim() : '',
          checknum: idxCheck >= 0 ? String(row[idxCheck] || '').trim() : ''
        });
      }
      if (!header.dtStart || !header.dtEnd) {
        const period = periodFromTransactions(transactions);
        if (period) {
          if (!header.dtStart) header.dtStart = period.dtStart;
          if (!header.dtEnd) header.dtEnd = period.dtEnd;
        }
      }
      const ofx = generateOFX(transactions, header);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([ofx], { type: 'text/plain;charset=utf-8;' }));
      a.download = file.name.replace(/\.(xlsx|xls|csv)$/i, '') + '.ofx';
      a.textContent = '⬇ ' + a.download;
      links.appendChild(a);
      processed++;
      const pct = Math.floor(processed / files.length * 100);
      progress.textContent = pct + '%'; bar.style.width = pct + '%';
      if (processed === files.length) { progress.textContent = 'Concluído!'; bar.style.width = '100%'; }
    };
    reader.readAsArrayBuffer(file);
  });
});


// Máscara DD/MM/AAAA ao digitar
function maskDateInput(el) {
  el.addEventListener('input', () => {
    let v = el.value.replace(/\D/g, '').slice(0, 8);
    if (v.length > 4) v = v.slice(0, 2) + '/' + v.slice(2, 4) + '/' + v.slice(4);
    else if (v.length > 2) v = v.slice(0, 2) + '/' + v.slice(2);
    el.value = v;
  });
  el.addEventListener('blur', () => {
    const m = el.value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) return;
    const d = parseInt(m[1], 10), mo = parseInt(m[2], 10), y = parseInt(m[3], 10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) {
      el.style.borderColor = '#f87171';
    } else {
      el.style.borderColor = '';
    }
  });
}
const fldDtStart = document.getElementById('fldDtStart');
const fldDtEnd = document.getElementById('fldDtEnd');
if (fldDtStart) maskDateInput(fldDtStart);
if (fldDtEnd) maskDateInput(fldDtEnd);

// ID → nome do banco
document.getElementById('fldBankId').addEventListener('input', () => {

  applyBankCodeToForm(document.getElementById('fldBankId').value);
});

// Busca por nome
const searchInput = document.getElementById('fldBankSearch');
const searchBox = document.getElementById('bankSearchResults');
if (searchInput && searchBox) {
  searchInput.addEventListener('input', () => {
    const results = searchBanks(searchInput.value);
    if (!results.length) {
      searchBox.classList.remove('open');
      searchBox.innerHTML = '';
      return;
    }
    searchBox.innerHTML = results.map(r =>
      '<button type="button" data-code="' + r.code + '"><span class="code">' + r.code +
      '</span>' + r.name + '</button>'
    ).join('');
    searchBox.classList.add('open');
  });
  searchBox.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-code]');
    if (!btn) return;
    applyBankCodeToForm(btn.getAttribute('data-code'));
    searchInput.value = '';
    searchBox.classList.remove('open');
    searchBox.innerHTML = '';
  });
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.bank-search-wrap')) {
      searchBox.classList.remove('open');
    }
  });
}
