'use strict';

const ownerSelect = document.getElementById('owner-select');
const codeInput = document.getElementById('code-input');
const status = document.getElementById('status');
const tableWrap = document.getElementById('table-wrap');
const resultBody = document.getElementById('result-body');
const resultCount = document.getElementById('result-count');
const dataPeriod = document.getElementById('data-period');
const assignmentMonth = document.getElementById('assignment-month');

let roster = [];
let peelRows = [];

function normalizeCode(value) {
  const code = String(value ?? '').trim().toUpperCase();
  return /^\d$/.test(code) ? code.padStart(2, '0') : code;
}

function displayCode(value) {
  const code = String(value);
  return /^0[1-9]$/.test(code) ? code.slice(1) : code;
}

async function fetchText(path) {
  const response = await fetch(path, { cache: 'no-store' });
  if (!response.ok) throw new Error(`${path}：HTTP ${response.status}`);
  return response.text();
}

function setMessage(message) {
  status.textContent = message;
  status.hidden = false;
  tableWrap.hidden = true;
  resultCount.textContent = '';
  resultBody.replaceChildren();
}

function appendCell(row, label, value) {
  const cell = document.createElement('td');
  cell.dataset.label = label;
  cell.textContent = String(value);
  row.appendChild(cell);
}

function renderResults() {
  const code = normalizeCode(codeInput.value);
  const known = roster.some(person => normalizeCode(person['代號']) === code);
  ownerSelect.value = known ? code : '';

  if (!code) {
    setMessage('請輸入代號，或從上方選擇負責人。');
    return;
  }
  if (!known) {
    setMessage(`查無代號「${codeInput.value.trim()}」，請重新輸入。`);
    return;
  }

  const matched = peelRows.filter(item => normalizeCode(item['代號']) === code);
  if (!matched.length) {
    setMessage(`代號「${displayCode(code)}」目前沒有剝藥資料。`);
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const item of matched) {
    const row = document.createElement('tr');
    appendCell(row, '機號－藥盒', item['藥盒M']);
    appendCell(row, '藥名', item['藥名']);
    appendCell(row, '代號', item['代號']);
    appendCell(row, '週耗量', item['周耗量']);
    appendCell(row, '負責人', item['負責人']);
    fragment.appendChild(row);
  }
  resultBody.replaceChildren(fragment);
  resultCount.textContent = `共 ${matched.length} 筆`;
  status.hidden = true;
  tableWrap.hidden = false;
}

function populateSelect() {
  ownerSelect.replaceChildren();
  const placeholder = new Option('請選擇代號與負責人', '');
  ownerSelect.add(placeholder);
  for (const person of roster) {
    const code = normalizeCode(person['代號']);
    ownerSelect.add(new Option(`${displayCode(code)}－${person['負責人']}`, code));
  }
  ownerSelect.disabled = false;
}

function showPeriod(period) {
  const start = period['資料起日'];
  const end = period['資料迄日'];
  const month = period['負責區月份'];
  dataPeriod.textContent = start && end ? `${start} ～ ${end}` : '待確認';
  if (assignmentMonth) assignmentMonth.textContent = month || '待確認';
}

async function loadData() {
  try {
    const [rosterText, peelText, periodText] = await Promise.all([
      fetchText('./responsible.json'),
      fetchText('./peel.json'),
      fetchText('./period.json')
    ]);
    const parsedRoster = JSON.parse(rosterText);
    const parsedRows = peelText.split(/\r?\n/).filter(line => line.trim()).map(line => JSON.parse(line));
    const period = JSON.parse(periodText);
    if (!Array.isArray(parsedRoster) || !Array.isArray(parsedRows) || !period || typeof period !== 'object') {
      throw new Error('資料格式錯誤');
    }
    roster = parsedRoster;
    peelRows = parsedRows;
    populateSelect();
    showPeriod(period);
    renderResults();
  } catch (error) {
    console.error('資料載入失敗', error);
    setMessage('資料載入失敗，請確認網路連線或資料檔案後重新整理。');
  }
}

ownerSelect.addEventListener('change', () => {
  if (!ownerSelect.value) return;
  codeInput.value = displayCode(ownerSelect.value);
  renderResults();
});
codeInput.addEventListener('input', renderResults);
loadData();
