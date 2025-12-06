// ===== GAS API URL =====
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwcRDMfcPWb205wkYWv1f8lanZeKGW5dVW0_8k48EYWqu2dmuwFjOVU8UPue-wg-QTVSw/exec';

// ==== 簡易認証（ローカルストレージ） ====
const LOCAL_STORAGE_AUTH_KEY = 'sheetAuthKey';
let authKey = localStorage.getItem(LOCAL_STORAGE_AUTH_KEY) || null;


// ==== 日付管理 ====
let currentDate = new Date();  // ヘッダーの日付（表示・シート両方の基準）
let currentSheetId = "";       // "YYYYMMDD" 形式

function dateToSheetId(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}${m}${d}`; // シート名
}

function formatHeaderDate(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();
  return `${y}年${m}月${d}日`;
}


function loadForCurrentDate() {
  currentSheetId = dateToSheetId(currentDate); // "YYYYMMDD"
  renderCurrentDateHeader();
  fetchAndRender(currentSheetId);
}

// ==== 部屋移動モーダル用 ====
let roomMoveModalEl = null;
let roomMoveTableBodyEl = null;
let roomMoveExecButton = null;
let roomMoveStatusEl = null;
let roomMoveSelectedRows = []; // シートの行番号（1始まり）を入れる
document.getElementById('roomMoveButton').addEventListener('click' , () => {
  openRoomMoveModal();
});

function initRoomMoveModal() {
  if (roomMoveModalEl) return;

  // オーバーレイ
  roomMoveModalEl = document.createElement('div');
  roomMoveModalEl.id = 'roomMoveModal';
  roomMoveModalEl.style.position = 'fixed';
  roomMoveModalEl.style.inset = '0';
  roomMoveModalEl.style.background = 'rgba(0,0,0,0.4)';
  roomMoveModalEl.style.display = 'none';
  roomMoveModalEl.style.zIndex = '9999';
  roomMoveModalEl.style.alignItems = 'center';
  roomMoveModalEl.style.justifyContent = 'center';

  // コンテンツ
  const content = document.createElement('div');
  content.style.background = '#fff';
  content.style.borderRadius = '8px';
  content.style.padding = '30px';
  content.style.minWidth = '320px';
  content.style.maxHeight = '80vh';
  content.style.display = 'flex';
  content.style.flexDirection = 'column';
  content.style.boxShadow = '0 2px 10px rgba(0,0,0,0.3)';

  const title = document.createElement('div');
  title.textContent = '部屋移動';
  title.style.fontWeight = 'bold';
  title.style.fontSize = '1.4rem';
  title.style.marginBottom = '8px';
  content.appendChild(title);

  const desc = document.createElement('div');
  desc.textContent = '入れ替えたい部屋を2つ選択してください。';
  desc.style.fontSize = '0.85rem';
  desc.style.marginBottom = '8px';
  content.appendChild(desc);

  // 簡易テーブル（No & 氏名）
  const tableWrapper = document.createElement('div');
  tableWrapper.style.flex = '1';
  tableWrapper.style.overflowY = 'auto';
  tableWrapper.style.marginBottom = '8px';

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.minWidth = '400px';
  table.style.borderCollapse = 'collapse';
  table.innerHTML = `
    <thead>
      <tr>
        <th style="border-bottom:1px solid #ccc; padding:4px;">No</th>
        <th style="border-bottom:1px solid #ccc; padding:4px;">氏名</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;

  roomMoveTableBodyEl = table.querySelector('tbody');
  tableWrapper.appendChild(table);
  content.appendChild(tableWrapper);

  // ステータス表示
  roomMoveStatusEl = document.createElement('div');
  roomMoveStatusEl.style.minHeight = '1.2em';
  roomMoveStatusEl.style.fontSize = '0.85rem';
  roomMoveStatusEl.style.marginBottom = '8px';
  content.appendChild(roomMoveStatusEl);

  // ボタン行
  const buttonRow = document.createElement('div');
  buttonRow.style.display = 'flex';
  buttonRow.style.flexDirection = 'column';
  buttonRow.style.gap = '10px';

  
  roomMoveExecButton = document.createElement('button');
  roomMoveExecButton.textContent = '部屋移動を実行';
  roomMoveExecButton.disabled = true;
  roomMoveExecButton.className = 'movebtn';
  roomMoveExecButton.addEventListener('click', () => {
    executeRoomMove();
  });


  buttonRow.appendChild(roomMoveExecButton);
  content.appendChild(buttonRow);

  roomMoveModalEl.appendChild(content);
  document.body.appendChild(roomMoveModalEl);

  // オーバーレイクリックで閉じる（中身クリックは閉じない）
  roomMoveModalEl.addEventListener('click', (e) => {
    if (e.target === roomMoveModalEl) {
      closeRoomMoveModal();
    }
  });
}
function openRoomMoveModal() {
  initRoomMoveModal();

  // 今表示中のメインテーブルから、No & 氏名 & 行番号を取得
  const mainTable = tableContainer.querySelector('table');
  if (!mainTable) {
    alert('テーブルが読み込まれていません。');
    return;
  }
  const mainTbody = mainTable.querySelector('tbody');
  if (!mainTbody) {
    alert('テーブルが読み込まれていません。');
    return;
  }

  roomMoveTableBodyEl.innerHTML = '';
  roomMoveSelectedRows = [];
  roomMoveStatusEl.textContent = '';
  roomMoveExecButton.disabled = true;

  const trs = Array.from(mainTbody.querySelectorAll('tr'));

  trs.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length < 3) return;

    const noText = tds[1].textContent || '';
    const nameText = tds[2].textContent || '';

    // No or 氏名セルに埋め込んでいる data-row-index からシート行番号を取得
    let rowIndex = null;
    for (const td of tds) {
      const idxStr = td.dataset && td.dataset.rowIndex;
      if (idxStr) {
        rowIndex = parseInt(idxStr, 10);
        break;
      }
    }
    if (!rowIndex) return;

    const modalTr = document.createElement('tr');
    modalTr.dataset.rowIndex = String(rowIndex);

    const tdNo = document.createElement('td');
    tdNo.textContent = noText;
    tdNo.style.padding = '4px';
    tdNo.style.borderBottom = '1px solid #eee';

    const tdName = document.createElement('td');
    tdName.textContent = nameText;
    tdName.style.padding = '4px';
    tdName.style.borderBottom = '1px solid #eee';

    modalTr.appendChild(tdNo);
    modalTr.appendChild(tdName);

    modalTr.addEventListener('click', () => {
      toggleRoomMoveRowSelection(modalTr);
    });

    roomMoveTableBodyEl.appendChild(modalTr);
  });

  roomMoveModalEl.style.display = 'flex';
}

function toggleRoomMoveRowSelection(tr) {
  const rowIndex = parseInt(tr.dataset.rowIndex, 10);
  if (!rowIndex) return;

  const selectedIdx = roomMoveSelectedRows.indexOf(rowIndex);

  if (selectedIdx >= 0) {
    // 選択解除
    roomMoveSelectedRows.splice(selectedIdx, 1);
    tr.classList.remove('selected');
  } else {
    // 2行までしか選択させない
    if (roomMoveSelectedRows.length >= 2) {
      return; // 3行目以降の選択は無視
    }
    roomMoveSelectedRows.push(rowIndex);
    tr.classList.add('selected');
  }

  // ちょうど2行のときだけ実行ボタンを有効化
  roomMoveExecButton.disabled = roomMoveSelectedRows.length !== 2;
}

function closeRoomMoveModal() {
  if (!roomMoveModalEl) return;
  roomMoveModalEl.style.display = 'none';
}

async function executeRoomMove() {
  if (roomMoveSelectedRows.length !== 2) return;

  const [row1, row2] = roomMoveSelectedRows;
  roomMoveExecButton.disabled = true;
  roomMoveStatusEl.textContent = '部屋移動中...';

  try {
    const res = await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        action: 'swapRows',
        date: currentSheetId,
        authKey: authKey,
        row1: row1,
        row2: row2
      })
    });

    const json = await res.json();
    if (!json.success) {
      roomMoveStatusEl.textContent = json.message || '部屋移動に失敗しました';
      roomMoveExecButton.disabled = false;
      return;
    }

    // 成功
    closeRoomMoveModal();
    // 再読み込み
    tableContainer.innerHTML = '';
    statusSpan.textContent = '読み込み中…';
    statusSpan.style.display = 'inline';
    loadForCurrentDate();

  } catch (e) {
    console.error(e);
    roomMoveStatusEl.textContent = 'エラー: ' + e.message;
    roomMoveExecButton.disabled = false;
  }
}




// =======================
// カラーピッカー関連（No列＆氏名列）
// =======================

const COLOR_OPTIONS = [
  'transparent', // 無色
  '#d8a393ff',
  '#e4e3b3ff',
  '#8acad3ff',
  '#b3aae4ff',
  '#d892b9ff',
  '#7aff5fff',
  '#20def7ff',
  '#bdbdbdff'
];
const CI_ROW_BG = '#dceeddff'; // C/Iの薄いグリーン
let colorPickerEl = null;
let currentRowIndexForColor = null; // どの行の色か（シート行番号）

// 連絡事項の文字色マップ
const NOTE_COLOR_MAP = {
  black: '#000000',
  red:   '#ff0000ff',
  blue:  '#0026ffff'
};

// 現在の連絡事項の色
let currentNoteColor = 'black';
// ラジオボタンへの参照を保持
let noteColorRadioMap = {};

// 客室係の候補（staffシートA列）と現在値
let staffOptions = [];



function initColorPicker() {
  if (colorPickerEl) return;

  colorPickerEl = document.createElement('div');
  colorPickerEl.className = 'color-picker pickwindow';
  colorPickerEl.style.position = 'absolute';
  colorPickerEl.style.display = 'none';
  colorPickerEl.style.background = '#fff';
  colorPickerEl.style.padding = '6px';
  colorPickerEl.style.borderRadius = '6px';
  colorPickerEl.style.zIndex = '9999';
  colorPickerEl.style.display = 'flex';
  colorPickerEl.style.flexWrap = 'wrap';

  COLOR_OPTIONS.forEach(color => {
    const swatch = document.createElement('div');
    swatch.className = 'color-swatch';
    swatch.style.width = '20px';
    swatch.style.height = '20px';
    swatch.style.borderRadius = '4px';
    swatch.style.border = '1px solid #ccc';
    swatch.style.cursor = 'pointer';
    swatch.style.margin = '2px';

    if (color === 'transparent') {
      swatch.classList.add('transparent');
      swatch.style.background =
        'repeating-linear-gradient(45deg,#f5f5f5,#f5f5f5 4px,#ddd 4px,#ddd 8px)';
      swatch.title = '無色';
    } else {
      swatch.style.backgroundColor = color;
    }

    swatch.addEventListener('click', (e) => {
      e.stopPropagation();
      const appliedColor = (color === 'transparent') ? '' : color;

      if (currentRowIndexForColor != null) {
        // 同じ行の No＆氏名セルすべてに適用
        const selector = `td[data-row-index="${currentRowIndexForColor}"][data-color-target="1"]`;
        document.querySelectorAll(selector).forEach(td => {
          td.style.backgroundColor = appliedColor;
        });

        // U列に保存
        saveColorToSheet(currentRowIndexForColor, appliedColor);
      }

      hideColorPicker();
    });

    colorPickerEl.appendChild(swatch);
  });

  document.body.appendChild(colorPickerEl);

  // ピッカー外クリックで閉じる
  document.addEventListener('click', (e) => {
    if (!colorPickerEl) return;
    if (colorPickerEl.contains(e.target)) return;
    hideColorPicker();
  });
}



function showColorPicker(cell, rowIndex) {
  initColorPicker();
  closeAllPickers();
  currentRowIndexForColor = rowIndex;

  colorPickerEl.style.display = 'flex';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    colorPickerEl.style.left = window.scrollX + rect.left + 'px';
    colorPickerEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    colorPickerEl.style.left = window.scrollX + rect.left + 'px';
    colorPickerEl.style.top  = window.scrollY + rect.top - colorPickerEl.offsetHeight - 4 + 'px';
  }
}


function hideColorPicker() {
  if (!colorPickerEl) return;
  colorPickerEl.style.display = 'none';
  currentRowIndexForColor = null;
}

// ==== 色情報をシート（U列:21列目）に保存 ====
async function saveColorToSheet(rowIndex, color) {
  try {
    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row: rowIndex,
        color: color || ''
      })
    });
  } catch (err) {
    console.error('色の保存に失敗しました:', err);
  }
}



// =======================
// ステータスピッカー（1列目）
// =======================
const STATUS_OPTIONS = ["", "荷", "荷,P", "C/I"];

let statusPickerEl = null;
let currentStatusTarget = null;
let currentStatusRowIndex = null;

function initStatusPicker() {
  if (statusPickerEl) return;

  statusPickerEl = document.createElement('div');
  statusPickerEl.className = 'status-picker pickwindow';
  statusPickerEl.style.position = 'absolute';
  statusPickerEl.style.display = 'none';
  statusPickerEl.style.background = '#fff';
  statusPickerEl.style.padding = '6px';
  statusPickerEl.style.borderRadius = '6px';
  statusPickerEl.style.zIndex = '9999';
  statusPickerEl.style.fontSize = '1.3rem';

  STATUS_OPTIONS.forEach(st => {
    const btn = document.createElement('div');
    btn.textContent = st === "" ? "(空白)" : st;
    btn.style.padding = '8px 8px';
    btn.style.cursor = 'pointer';

    btn.addEventListener('click', () => {
      if (currentStatusTarget) {
        // 1列目セルにステータス文字を反映
        currentStatusTarget.textContent = st;

        // ★ 行全体の背景を更新（C/I なら薄グリーン、それ以外は解除）
        const tr = currentStatusTarget.parentElement; // このセルの行
        updateRowStatusBackground(tr, st);

        // ★ V列 に保存
        if (currentStatusRowIndex != null) {
          saveStatusToSheet(currentStatusRowIndex, st);
        }
      }
      hideStatusPicker();
    });

    statusPickerEl.appendChild(btn);
  });

  document.body.appendChild(statusPickerEl);

  document.addEventListener('click', (e) => {
    if (!statusPickerEl) return;
    if (statusPickerEl.contains(e.target)) return;
    hideStatusPicker();
  });
}


// C/I 行の背景色を反映するヘルパー
function updateRowStatusBackground(tr, status) {
  if (!tr) return;
  if (status === "C/I") {
    tr.style.backgroundColor = CI_ROW_BG; // 薄いグリーン
  } else {
    tr.style.backgroundColor = "";        // 元に戻す
  }
}


function showStatusPicker(cell, rowIndex) {
  initStatusPicker();
  closeAllPickers();
  currentStatusTarget = cell;
  currentStatusRowIndex = rowIndex;

  statusPickerEl.style.display = 'block';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    statusPickerEl.style.left = window.scrollX + rect.left + 'px';
    statusPickerEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    statusPickerEl.style.left = window.scrollX + rect.left + 'px';
    statusPickerEl.style.top  = window.scrollY + rect.top - statusPickerEl.offsetHeight - 4 + 'px';
  }
}


function hideStatusPicker() {
  if (!statusPickerEl) return;
  statusPickerEl.style.display = 'none';
  currentStatusTarget = null;
  currentStatusRowIndex = null;
}

// ==== ステータスを V列(22列目) に保存 ====
async function saveStatusToSheet(rowIndex, status) {
  try {
    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row: rowIndex,
        status: status
      })
    });
  } catch (err) {
    console.error("ステータス保存失敗:", err);
  }
}




// ==== シートの時間データを表示用に整形 ====
function normalizeTimeValue(v) {
  if (v === null || v === undefined || v === "") return "";

  // そのまま使いたいテキスト（例: 弁当）
  const s = String(v).trim();
  if (s === "弁当") return s;

  // すでに "H:MM" / "HH:MM" の形式ならそのまま
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;

  // Date型の場合
  if (v instanceof Date) {
    const h = v.getHours();
    const m = v.getMinutes();
    return `${h}:${String(m).padStart(2, "0")}`;
  }

  // ISO文字列などを Date として解釈して時間だけ取り出す
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const h = d.getHours();
    const m = d.getMinutes();
    return `${h}:${String(m).padStart(2, "0")}`;
  }

  // どうしても解釈できないものはそのまま
  return s;
}


// ==== 人数変更モーダル用 ====
let guestModalEl = null;
let guestNoEl = null;
let guestNameEl = null;
let guestCountEls = {}; // { male, female, child, infant } の表示用
let guestCounts = { male: 0, female: 0, child: 0, infant: 0 };
let currentGuestRowIndex = null;

function initGuestModal() {
  if (guestModalEl) return;

  // オーバーレイ
  guestModalEl = document.createElement('div');
  guestModalEl.id = 'guestModal';
  guestModalEl.className = 'guest-modal-overlay';  // ★ 見た目はCSSに任せる
  guestModalEl.style.display = 'none';             // 表示/非表示だけJSで制御

  const content = document.createElement('div');
  content.className = 'guest-modal';               // ★ 本体の見た目はCSS側

  const title = document.createElement('div');
  title.textContent = '人数変更';
  title.className = 'guest-modal-title';
  content.appendChild(title);

  const noRow = document.createElement('div');
  noRow.className = 'guest-modal-row';
  guestNoEl = document.createElement('span');
  noRow.appendChild(document.createTextNode(''));
  noRow.appendChild(guestNoEl);
  content.appendChild(noRow);

  const nameRow = document.createElement('div');
  nameRow.className = 'guest-modal-row';
  guestNameEl = document.createElement('span');
  nameRow.appendChild(document.createTextNode(''));
  nameRow.appendChild(guestNameEl);
  content.appendChild(nameRow);

  // 人数行
  const labels = [
    { key: 'male',   label: '男' },
    { key: 'female', label: '女' },
    { key: 'child',  label: '子供' },
    { key: 'infant', label: '幼児' }
  ];

  const table = document.createElement('table');
  table.className = 'guest-modal-table';

  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  labels.forEach(info => {
    const th = document.createElement('th');
    th.textContent = info.label;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const trUp = document.createElement('tr');
  const trVal = document.createElement('tr');
  const trDown = document.createElement('tr');

  labels.forEach(info => {
    // ▲
    const tdUp = document.createElement('td');
    tdUp.className = 'guest-modal-cell';
    const btnUp = document.createElement('button');
    btnUp.textContent = '▲';
    btnUp.className = 'guest-modal-arrow guest-modal-arrow-up';
    btnUp.addEventListener('click', () => {
      changeGuestCount(info.key, +1);
    });
    tdUp.appendChild(btnUp);
    trUp.appendChild(tdUp);

    // 値
    const tdVal = document.createElement('td');
    tdVal.className = 'guest-modal-cell';
    const span = document.createElement('span');
    span.textContent = '0';
    span.className = 'guest-modal-count';
    guestCountEls[info.key] = span;
    tdVal.appendChild(span);
    trVal.appendChild(tdVal);

    // ▼
    const tdDown = document.createElement('td');
    tdDown.className = 'guest-modal-cell';
    const btnDown = document.createElement('button');
    btnDown.textContent = '▼';
    btnDown.className = 'guest-modal-arrow guest-modal-arrow-down';
    btnDown.addEventListener('click', () => {
      changeGuestCount(info.key, -1);
    });
    tdDown.appendChild(btnDown);
    trDown.appendChild(tdDown);
  });

  tbody.appendChild(trUp);
  tbody.appendChild(trVal);
  tbody.appendChild(trDown);
  table.appendChild(tbody);

  content.appendChild(table);

  // ボタン行
  const btnRow = document.createElement('div');
  btnRow.className = 'guest-modal-buttons';

  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'キャンセル';
  cancelBtn.className = 'guest-modal-button guest-modal-button-secondary';
  cancelBtn.addEventListener('click', () => {
    hideGuestModal();
  });

  const applyBtn = document.createElement('button');
  applyBtn.textContent = '変更';
  applyBtn.className = 'guest-modal-button guest-modal-button-primary';
  applyBtn.addEventListener('click', () => {
    applyGuestCounts(applyBtn);
  });

  btnRow.appendChild(cancelBtn);
  btnRow.appendChild(applyBtn);
  content.appendChild(btnRow);

  guestModalEl.appendChild(content);
  document.body.appendChild(guestModalEl);

  // オーバーレイクリックで閉じる（中身クリックは閉じない）
  guestModalEl.addEventListener('click', (e) => {
    if (e.target === guestModalEl) {
      hideGuestModal();
    }
  });
}


function changeGuestCount(key, delta) {
  let v = guestCounts[key] || 0;
  v += delta;
  if (v < 0) v = 0;
  guestCounts[key] = v;
  if (guestCountEls[key]) {
    guestCountEls[key].textContent = String(v);
  }
}

function showGuestModal(rowIndex, noText, nameText, male, female, child, infant) {
  initGuestModal();
  closeAllPickers && closeAllPickers(); // 他のピッカーは閉じる（関数があれば）

  currentGuestRowIndex = rowIndex;

  guestNoEl.textContent = noText || '';
  guestNameEl.textContent = nameText || '';

  guestCounts.male   = male   || 0;
  guestCounts.female = female || 0;
  guestCounts.child  = child  || 0;
  guestCounts.infant = infant || 0;

  Object.keys(guestCountEls).forEach(key => {
    guestCountEls[key].textContent = String(guestCounts[key] || 0);
  });

  guestModalEl.style.display = 'flex';
}

function hideGuestModal() {
  if (!guestModalEl) return;
  guestModalEl.style.display = 'none';
  currentGuestRowIndex = null;
}

async function applyGuestCounts(applyBtn) {
  if (!currentGuestRowIndex) {
    hideGuestModal();
    return;
  }

  if (applyBtn) {
    applyBtn.disabled = true;
  }

  try {
    await saveGuestCountsToSheet(currentGuestRowIndex, guestCounts);
    hideGuestModal();
    // 再読み込み（現在日付のシート）
    loadForCurrentDate();
  } catch (e) {
    console.error('人数更新エラー:', e);
    alert('人数の更新に失敗しました: ' + e.message);
  } finally {
     if (applyBtn) {
      applyBtn.disabled = false;
    }
  }
}

async function saveGuestCountsToSheet(rowIndex, counts) {
  await fetch(GAS_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({
      action: 'updateGuests',
      date: currentSheetId,
      authKey: authKey,
      row: rowIndex,
      male: counts.male || 0,
      female: counts.female || 0,
      child: counts.child || 0,
      infant: counts.infant || 0
    })
  });
}





// ==== 夕食（W列）を保存 ====
async function saveDinnerToSheet(rowIndex, time) {
  try {
    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row: rowIndex,
        dinner: time || ''
      })
    });

    const idx = rowIndex - 2;
    if (idx >= 0) {
      currentDinnerValues[idx] = time || '';
    }
  } catch (err) {
    console.error("夕食保存失敗:", err);
  }
}



// ==== 夕食ピッカー ====
const DINNER_OPTIONS = ["17:30", "18:00", "18:30", "19:00", "19:30", "20:00"];
// 夕食の現在値（行ごとの値）を保持してカウントに使う
let currentDinnerValues = [];
// 時間ごとの件数表示用の span を保持
let dinnerCountSpans = {};


let dinnerPickerEl = null;
let currentDinnerCell = null;
let currentDinnerRowIndex = null;

function initDinnerPicker() {
  if (dinnerPickerEl) return;

  dinnerPickerEl = document.createElement('div');
  dinnerPickerEl.className = 'dinner-picker pickwindow';
  dinnerPickerEl.style.position = 'absolute';
  dinnerPickerEl.style.display = 'none';
  dinnerPickerEl.style.background = '#fff';
  dinnerPickerEl.style.padding = '10px';
  dinnerPickerEl.style.borderRadius = '6px';
  dinnerPickerEl.style.zIndex = '9999';

  dinnerCountSpans = {}; // 件数表示用

  DINNER_OPTIONS.forEach(t => {
    const line = document.createElement('div');
    line.style.display = 'flex';
    line.style.alignItems = 'center';
    line.style.gap = '10px';
    line.style.padding = '2px 0';

    const btn = document.createElement('div');
    btn.textContent = t;
    btn.style.fontSize = '1.4rem';
    btn.style.padding = '4px 8px';
    btn.style.cursor = 'pointer';
    btn.style.flex = '0 0 auto';

    const countSpan = document.createElement('span');
    countSpan.style.fontSize = '1rem';
    countSpan.style.color = '#555';
    countSpan.textContent = '';

    dinnerCountSpans[t] = countSpan;

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (currentDinnerCell && currentDinnerRowIndex != null) {
        currentDinnerCell.textContent = t;
        saveDinnerToSheet(currentDinnerRowIndex, t);
      }
      hideDinnerPicker(); // ★ 必ず閉じる
    });

    line.appendChild(btn);
    line.appendChild(countSpan);
    dinnerPickerEl.appendChild(line);
  });

  // テキスト入力
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'その他の時間...';
  input.style.marginTop = '10px';
  input.style.width = '100%';
  input.style.padding = '8px';

  const okBtn = document.createElement('button');
  okBtn.textContent = 'OK';
  okBtn.style.marginTop = '10px';
  okBtn.style.width = '100%';
  okBtn.style.padding = '8px';

  okBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const val = input.value.trim();
    if (currentDinnerCell && currentDinnerRowIndex != null) {
      currentDinnerCell.textContent = val;
      saveDinnerToSheet(currentDinnerRowIndex, val);
    }
    hideDinnerPicker(); // ★ 必ず閉じる
  });

  dinnerPickerEl.appendChild(input);
  dinnerPickerEl.appendChild(okBtn);

  document.body.appendChild(dinnerPickerEl);

  // ピッカー外クリックで閉じる
  document.addEventListener('click', (e) => {
    if (!dinnerPickerEl) return;
    if (dinnerPickerEl.contains(e.target)) return;
    hideDinnerPicker();
  });
}
// ==== 夕食ピッカーを閉じる ====
function hideDinnerPicker() {
  if (!dinnerPickerEl) return;
  dinnerPickerEl.style.display = 'none';
  currentDinnerCell = null;
  currentDinnerRowIndex = null;
}



// ==== 夕食の件数表示を更新 ====
function updateDinnerCounts() {
  if (!Array.isArray(currentDinnerValues)) return;

  // 各時間のカウントを初期化
  const counts = {};
  DINNER_OPTIONS.forEach(t => { counts[t] = 0; });

  // 現在の夕食値を集計
  currentDinnerValues.forEach(v => {
    if (!v) return;
    if (counts[v] !== undefined) {
      counts[v]++;
    }
  });

  // span に反映
  DINNER_OPTIONS.forEach(t => {
    const span = dinnerCountSpans[t];
    if (!span) return;
    const n = counts[t] || 0;
    span.textContent = n > 0 ? ` [${n}件]` : '';
  });
}



function showDinnerPicker(cell, rowIndex, currentValue) {
  initDinnerPicker();
  closeAllPickers();
  currentDinnerCell = cell;
  currentDinnerRowIndex = rowIndex;

  const input = dinnerPickerEl.querySelector('input');
  if (input) input.value = currentValue || '';

  updateDinnerCounts();

  dinnerPickerEl.style.display = 'block';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    dinnerPickerEl.style.left = window.scrollX + rect.left + 'px';
    dinnerPickerEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    dinnerPickerEl.style.left = window.scrollX + rect.left + 'px';
    dinnerPickerEl.style.top  = window.scrollY + rect.top - dinnerPickerEl.offsetHeight - 4 + 'px';
  }
}




// ==== 朝食（X列）を保存 ====
async function saveBreakfastToSheet(rowIndex, time) {
  try {
    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row: rowIndex,
        breakfast: time || ''
      })
    });
  } catch (err) {
    console.error("朝食保存失敗:", err);
  }
}


// ==== 朝食ピッカー ====
const BREAKFAST_OPTIONS = ["8:00", "8:30", "9:00", "弁当"];

let breakfastPickerEl = null;
let currentBreakfastCell = null;
let currentBreakfastRowIndex = null;

function initBreakfastPicker() {
  if (breakfastPickerEl) return;

  breakfastPickerEl = document.createElement('div');
  breakfastPickerEl.className = 'breakfast-picker pickwindow';
  breakfastPickerEl.style.position = 'absolute';
  breakfastPickerEl.style.display = 'none';
  breakfastPickerEl.style.background = '#fff';
  breakfastPickerEl.style.padding = '10px';
  breakfastPickerEl.style.borderRadius = '6px';
  breakfastPickerEl.style.zIndex = '9999';

  BREAKFAST_OPTIONS.forEach(t => {
    const btn = document.createElement('div');
    btn.textContent = t;
    btn.style.padding = '4px 8px';
    btn.style.fontSize = '1.4rem';
    btn.style.cursor = 'pointer';

    btn.addEventListener('click', () => {
      if (currentBreakfastCell && currentBreakfastRowIndex != null) {
        currentBreakfastCell.textContent = t;
        saveBreakfastToSheet(currentBreakfastRowIndex, t);
      }
      hideBreakfastPicker();
    });

    breakfastPickerEl.appendChild(btn);
  });

  // テキスト入力
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'その他...';
  input.style.marginTop = '10px';
  input.style.width = '100%';
  input.style.padding = '8px';

  const okBtn = document.createElement('button');
  okBtn.textContent = 'OK';
  okBtn.style.marginTop = '10px';
  okBtn.style.padding = '8px';
  okBtn.style.width = '100%';

  okBtn.addEventListener('click', () => {
    const val = input.value.trim();
    if (currentBreakfastCell && currentBreakfastRowIndex != null) {
      currentBreakfastCell.textContent = val;
      saveBreakfastToSheet(currentBreakfastRowIndex, val);
    }
    hideBreakfastPicker();
  });

  breakfastPickerEl.appendChild(input);
  breakfastPickerEl.appendChild(okBtn);

  document.body.appendChild(breakfastPickerEl);

  document.addEventListener('click', (e) => {
    if (!breakfastPickerEl) return;
    if (breakfastPickerEl.contains(e.target)) return;
    hideBreakfastPicker();
  });
}

function showBreakfastPicker(cell, rowIndex, currentValue) {
  initBreakfastPicker();
  closeAllPickers();
  currentBreakfastCell = cell;
  currentBreakfastRowIndex = rowIndex;

  const input = breakfastPickerEl.querySelector('input');
  if (input) input.value = currentValue || '';

  breakfastPickerEl.style.display = 'block';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    breakfastPickerEl.style.left = window.scrollX + rect.left + 'px';
    breakfastPickerEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    breakfastPickerEl.style.left = window.scrollX + rect.left + 'px';
    breakfastPickerEl.style.top  = window.scrollY + rect.top - breakfastPickerEl.offsetHeight - 4 + 'px';
  }
}


function hideBreakfastPicker() {
  if (!breakfastPickerEl) return;
  breakfastPickerEl.style.display = 'none';
  currentBreakfastCell = null;
  currentBreakfastRowIndex = null;
}


// ==== 連絡事項（Y列）を保存 ====
async function saveNoteToSheet(rowIndex, text, colorKey) {
  try {
    const payload = JSON.stringify({
      text:  text || '',
      color: colorKey || 'black'
    });

    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row:  rowIndex,
        note: payload
      })
    });
  } catch (err) {
    console.error("連絡事項保存失敗:", err);
  }
}


// ==== 連絡事項入力 ====
let noteEditorEl = null;
let noteTextarea = null;
let currentNoteCell = null;
let currentNoteRowIndex = null;

// ==== 連絡事項セル文字列 → { text, color } に分解 ====
function parseNoteCell(v) {
  if (v === null || v === undefined || v === '') {
    return { text: '', color: 'black' };
  }

  const s = String(v).trim();

  // JSON形式 {"text":"...","color":"red"} を優先
  if (s.startsWith('{') && s.endsWith('}')) {
    try {
      const obj = JSON.parse(s);
      const text = typeof obj.text === 'string' ? obj.text : s;
      const color = typeof obj.color === 'string' ? obj.color : 'black';
      return { text, color };
    } catch (e) {
      // 失敗したら普通のテキスト扱い
      return { text: s, color: 'black' };
    }
  }

  // 古いデータ（プレーンテキスト）は黒として扱う
  return { text: s, color: 'black' };
}


function initNoteEditor() {
  if (noteEditorEl) return;

  noteEditorEl = document.createElement('div');
  noteEditorEl.className = 'note-editor pickwindow';
  noteEditorEl.style.position = 'absolute';
  noteEditorEl.style.display = 'none';
  noteEditorEl.style.background = '#fff';
  noteEditorEl.style.padding = '10px';
  noteEditorEl.style.borderRadius = '6px';
  noteEditorEl.style.zIndex = '9999';
  noteEditorEl.style.width = '260px';
  noteEditorEl.id = 'Notetxt';

  // テキストエリア
  noteTextarea = document.createElement('textarea');
  noteTextarea.rows = 4;
  noteTextarea.style.width = '100%';

  // 色選択（黒・赤・青のラジオボタン）
  const colorWrapper = document.createElement('div');
  colorWrapper.style.display = 'flex';
  colorWrapper.style.gap = '8px';
  colorWrapper.style.marginTop = '4px';
  colorWrapper.style.alignItems = 'center';

  const labelTitle = document.createElement('span');
  labelTitle.textContent = '文字色:';
  labelTitle.style.fontSize = '0.85rem';
  colorWrapper.appendChild(labelTitle);

  noteColorRadioMap = {};
  ['black', 'red', 'blue'].forEach(colorKey => {
    const label = document.createElement('label');
    label.style.display = 'flex';
    label.style.alignItems = 'center';
    label.style.gap = '2px';
    label.style.fontSize = '0.85rem';

    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'noteColor';
    radio.value = colorKey;

    const swatch = document.createElement('span');
    swatch.textContent =
      colorKey === 'black' ? '黒' :
      colorKey === 'red'   ? '赤' : '青';
    swatch.style.color = NOTE_COLOR_MAP[colorKey];

    radio.addEventListener('change', () => {
      if (radio.checked) {
        currentNoteColor = colorKey;
      }
    });

    noteColorRadioMap[colorKey] = radio;

    label.appendChild(radio);
    label.appendChild(swatch);
    colorWrapper.appendChild(label);
  });

  // OK ボタン
  const okBtn = document.createElement('button');
  okBtn.textContent = 'OK';
  okBtn.style.marginTop = '10px';
  okBtn.style.padding = '10px';
  okBtn.style.width = '100%';

  okBtn.addEventListener('click', () => {
    const val = noteTextarea.value.trim();
    if (currentNoteCell && currentNoteRowIndex != null) {
      currentNoteCell.textContent = val;
      currentNoteCell.style.color =
        NOTE_COLOR_MAP[currentNoteColor] || NOTE_COLOR_MAP.black;

      // テキストと色を Y列に JSON 文字列として保存
      saveNoteToSheet(currentNoteRowIndex, val, currentNoteColor);
    }
    hideNoteEditor();
  });

  noteEditorEl.appendChild(noteTextarea);
  noteEditorEl.appendChild(colorWrapper);
  noteEditorEl.appendChild(okBtn);

  document.body.appendChild(noteEditorEl);

  document.addEventListener('click', (e) => {
    if (!noteEditorEl) return;
    if (noteEditorEl.contains(e.target)) return;
    hideNoteEditor();
  });
}


function showNoteEditor(cell, rowIndex, currentText, currentColorKey) {
  initNoteEditor();
  closeAllPickers();
  currentNoteCell = cell;
  currentNoteRowIndex = rowIndex;

  noteTextarea.value = currentText || '';

  currentNoteColor = currentColorKey || 'black';
  Object.entries(noteColorRadioMap).forEach(([key, radio]) => {
    radio.checked = (key === currentNoteColor);
  });

  noteEditorEl.style.display = 'block';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    noteEditorEl.style.left = window.scrollX + rect.left - 200 + 'px';
    noteEditorEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    noteEditorEl.style.left = window.scrollX + rect.left - 200 + 'px';
    noteEditorEl.style.top  = window.scrollY + rect.top - noteEditorEl.offsetHeight - 4 + 'px';
  }

  setTimeout(() => {
  document.getElementById('Notetxt').focus();
}, 50);

}

function hideNoteEditor() {
  if (!noteEditorEl) return;
  noteEditorEl.style.display = 'none';
  currentNoteCell = null;
  currentNoteRowIndex = null;
}


// ==== 客室係（Z列）を保存 ====
async function saveStaffToSheet(rowIndex, staffName) {
  try {
    await fetch(GAS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        date: currentSheetId,
        authKey: authKey,
        row: rowIndex,
        staff: staffName || ''
      })
    });
  } catch (err) {
    console.error("客室係の保存失敗:", err);
  }
}



// ==== 客室係ピッカー ====
let staffPickerEl = null;
let currentStaffCell = null;
let currentStaffRowIndex = null;

function initStaffPicker() {
  if (staffPickerEl) return;

  staffPickerEl = document.createElement('div');
  staffPickerEl.className = 'staff-picker pickwindow';
  staffPickerEl.style.position = 'absolute';
  staffPickerEl.style.display = 'none';
  staffPickerEl.style.background = '#fff';
  staffPickerEl.style.padding = '6px';
  staffPickerEl.style.borderRadius = '6px';
  staffPickerEl.style.zIndex = '9999';
  staffPickerEl.style.maxHeight = '240px';
  staffPickerEl.style.overflowY = 'auto';
  staffPickerEl.style.minWidth = '120px';

  document.body.appendChild(staffPickerEl);

  document.addEventListener('click', (e) => {
    if (!staffPickerEl) return;
    if (staffPickerEl.contains(e.target)) return;
    hideStaffPicker();
  });
}

function showStaffPicker(cell, rowIndex, currentValue) {
  initStaffPicker();
  closeAllPickers();
  currentStaffCell = cell;
  currentStaffRowIndex = rowIndex;

  // 一旦中身クリア
  staffPickerEl.innerHTML = "";

  if (!staffOptions || staffOptions.length === 0) {
    const msg = document.createElement('div');
    msg.textContent = "客室係リストがありません";
    msg.style.fontSize = '1rem';
    staffPickerEl.appendChild(msg);
  } else {
    staffOptions.forEach(name => {
      const btn = document.createElement('div');
      btn.textContent = name;
      btn.style.padding = '10px 8px';
      btn.style.cursor = 'pointer';
      if (name === currentValue) {
        btn.style.fontWeight = 'bold';
      }
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (currentStaffCell && currentStaffRowIndex != null) {
          currentStaffCell.textContent = name;
          saveStaffToSheet(currentStaffRowIndex, name);
        }
        hideStaffPicker();
      });
      staffPickerEl.appendChild(btn);
    });
  }

  // 表示位置（12行目より下は上側に表示）
  staffPickerEl.style.display = 'block';
  const rect = cell.getBoundingClientRect();
  const isBelow = rowIndex <= 12;

  if (isBelow) {
    staffPickerEl.style.left = window.scrollX + rect.left + 'px';
    staffPickerEl.style.top  = window.scrollY + rect.bottom + 4 + 'px';
  } else {
    staffPickerEl.style.left = window.scrollX + rect.left + 'px';
    staffPickerEl.style.top  = window.scrollY + rect.top - staffPickerEl.offsetHeight - 4 + 'px';
  }
}

function hideStaffPicker() {
  if (!staffPickerEl) return;
  staffPickerEl.style.display = 'none';
  currentStaffCell = null;
  currentStaffRowIndex = null;
}



// ==== すべてのピッカーを閉じる ====
function closeAllPickers() {
  if (colorPickerEl)      colorPickerEl.style.display = 'none';
  if (statusPickerEl)     statusPickerEl.style.display = 'none';
  if (staffPickerEl)      staffPickerEl.style.display = 'none';
  if (dinnerPickerEl)     dinnerPickerEl.style.display = 'none';
  if (breakfastPickerEl)  breakfastPickerEl.style.display = 'none';
  if (noteEditorEl)       noteEditorEl.style.display = 'none';
}



// =======================
// 今日の日付をヘッダーに表示
// =======================
function renderCurrentDateHeader() {
  const headerEl = document.getElementById("todayHeader");
  if (!headerEl) return;

  const dateInput = document.getElementById("dateInput");
  if (!dateInput) return;

  const y = currentDate.getFullYear();
  const m = String(currentDate.getMonth() + 1).padStart(2, '0');
  const d = String(currentDate.getDate()).padStart(2, '0');

  // ★ input の value を更新（表示はブラウザ標準の date UI）
  dateInput.value = `${y}-${m}-${d}`;
}


function initDateNavigation() {
  const headerEl = document.getElementById("todayHeader");
  if (!headerEl) return;

  // いったん中身を空にする
  headerEl.innerHTML = "";

  // 前日ボタン「<<」
  const prevBtn = document.createElement('button');
  prevBtn.textContent = '<<';
  prevBtn.id = 'prevDayButton';
  prevBtn.className = 'daybtn';
  prevBtn.addEventListener('click', () => {
    changeCurrentDateByDays(-1);
  });

  // 日付入力（type=date）
  const dateInput = document.createElement('input');
  dateInput.type = 'date';
  dateInput.id = 'dateInput';
  dateInput.style.margin = '0 4px';

  // 日付が変更されたら、その日付のシートを読み込む
  dateInput.addEventListener('change', () => {
    const val = dateInput.value; // "YYYY-MM-DD"
    if (!val) return;
    const [y, m, d] = val.split('-').map(Number);
    currentDate = new Date(y, m - 1, d);

    tableContainer.innerHTML = '';
    statusSpan.textContent = '読み込み中…';
    statusSpan.style.display = 'inline';
    loadForCurrentDate();
  });

  // 次日ボタン「>>」
  const nextBtn = document.createElement('button');
  nextBtn.textContent = '>>';
  nextBtn.id = 'nextDayButton';
  nextBtn.className = 'daybtn';
  nextBtn.addEventListener('click', () => {
    changeCurrentDateByDays(1);
  });

  // ヘッダーに配置：<< [date] >>
  headerEl.appendChild(prevBtn);
  headerEl.appendChild(dateInput);
  headerEl.appendChild(nextBtn);
}


function changeCurrentDateByDays(delta) {
  currentDate.setDate(currentDate.getDate() + delta);
  
  tableContainer.innerHTML = '';
  statusSpan.textContent = '読み込み中…';
  statusSpan.style.display = 'inline';
  loadForCurrentDate();
}

// 初期化：ページロード時に呼ぶ
window.addEventListener('DOMContentLoaded', () => {
  currentDate = new Date();   // 今日
  initDateNavigation();       // ヘッダーに type=date＋ボタンをセット
  
  statusSpan.textContent = '読み込み中…';
  statusSpan.style.display = 'inline';
  loadForCurrentDate();       //  自動で今日のシートを読み込む
});

// ==== 60秒ごとに自動リロード ====
setInterval(() => {
  // ★ ピッカーが開いているときは更新しない
  if (isAnyPickerOpen()) {
    return;
  }
  // ★ 読み込み
  loadForCurrentDate();
}, 60000);


function isAnyPickerOpen() {
  const pickers = document.querySelectorAll('.pickwindow');
  for (const p of pickers) {
    const style = window.getComputedStyle(p);
    if (style.display !== 'none') {
      return true;
    }
  }
  return false;
}





// =======================
// DOM 取得
// =======================
const loadButton = document.getElementById('loadButton');
const statusSpan = document.getElementById('status');
const tableContainer = document.getElementById('tableContainer');

// =======================
// ボタン押下イベント
// =======================
loadButton.addEventListener('click', () => {
  statusSpan.textContent = '読み込み中…';
  statusSpan.style.display = 'inline';
  loadForCurrentDate();
});


// =======================
// データ取得 + 加工 + 表示
// =======================
async function fetchAndRender(sheetId) {
  const targetSheetId = sheetId || currentSheetId || dateToSheetId(currentDate);

  try {
    // ★ すでに authKey を持っている端末 → どの日付でもキーワード認証は一切しない
    if (authKey) {
      await fetchWithKey(targetSheetId);
      return;
    }

    // ★ authKey が無い端末だけ、最初の1回だけキーワード認証
    await authWithKeywordAndFetch(targetSheetId);

  } catch (e) {
    console.error(e);
    statusSpan.textContent = 'エラー: ' + e.message;
    statusSpan.style.display = 'inline';
  }
}




function handleFetchedData(json) {
  const data          = json.data;
  const colors        = json.colors      || [];
  const statuses      = json.statuses    || [];
  const dinnersRaw    = json.dinners     || [];
  const breakfastsRaw = json.breakfasts  || [];
  const notesRaw      = json.notes       || [];
  const staffList     = json.staffList   || [];
  const staffValues   = json.staffValues || [];

  // データが無い場合（=シートが無い/空）はメッセージだけ表示して終了
  if (!data || data.length === 0) {
    statusSpan.textContent = json.message || 'シートがありません';
    statusSpan.style.display = 'inline';
    tableContainer.innerHTML = '';
    return false; // 失敗扱い
  }

  // 夕食・朝食
  const dinners    = dinnersRaw.map(normalizeTimeValue);
  const breakfasts = breakfastsRaw.map(normalizeTimeValue);
  currentDinnerValues = dinners.slice();

  // 連絡事項
  const noteTexts  = [];
  const noteColors = [];
  notesRaw.forEach(v => {
    const parsed = parseNoteCell(v);
    noteTexts.push(parsed.text);
    noteColors.push(parsed.color);
  });

  // 客室係リスト
  staffOptions = staffList.slice();

  const withNights  = convertNightsFormat(data);
  const withNo      = transformNoColumn(withNights);
  const finalMatrix = addExtraColumns(withNo);

  renderTable(finalMatrix, colors, statuses, dinners, breakfasts, noteTexts, noteColors, staffValues);
  return true; // 成功
}


async function authWithKeywordAndFetch(sheetId) {
  const keyword = window.prompt('キーワードを入力してください');
  if (keyword == null || keyword.trim() === '') {
    // キャンセル or 空入力
    statusSpan.textContent = '認証がキャンセルされました';
    statusSpan.style.display = 'inline';
    tableContainer.innerHTML = '';
    return false;
  }

  try {
    const url = `${GAS_API_URL}?date=${encodeURIComponent(sheetId)}&keyword=${encodeURIComponent(keyword.trim())}`;
    const res = await fetch(url);
    const json = await res.json();

    if (json.authorized === false) {
      // キーワード間違い
      statusSpan.textContent = json.message || '認証に失敗しました';
      statusSpan.style.display = 'inline';
      tableContainer.innerHTML = '';
      return false;
    }

    // ★ 正しいキーワード → GAS から固定キーを受け取り、localStorage に保存
    if (json.authKey) {
      authKey = json.authKey;
      localStorage.setItem(LOCAL_STORAGE_AUTH_KEY, authKey);
    }

    const ok = handleFetchedData(json);
    if (ok) {
      statusSpan.textContent = '';
      statusSpan.style.display = 'none';
    }
    // シートがない場合は handleFetchedData がメッセージを表示している
    return ok;

  } catch (e) {
    console.error(e);
    statusSpan.textContent = 'エラー: ' + e.message;
    statusSpan.style.display = 'inline';
    tableContainer.innerHTML = '';
    return false;
  }
}



async function fetchWithKey(sheetId) {
  try {
    const url = `${GAS_API_URL}?date=${encodeURIComponent(sheetId)}&key=${encodeURIComponent(authKey)}`;
    const res = await fetch(url);
    const json = await res.json();

    if (json.authorized === false) {
      // ★ キーが無効（GAS側の認証エラー）
      statusSpan.textContent = json.message || '認証に失敗しました（authKeyが無効です）';
      statusSpan.style.display = 'inline';
      tableContainer.innerHTML = '';
      return false;
    }

    const ok = handleFetchedData(json);
    if (ok) {
      // データがあるときだけステータスクリア
      statusSpan.textContent = '';
      statusSpan.style.display = 'none';
    }
    // シートがないときは handleFetchedData がメッセージを出している
    return ok;

  } catch (e) {
    console.error(e);
    statusSpan.textContent = 'エラー: ' + e.message;
    statusSpan.style.display = 'inline';
    return false;
  }
}



// =======================
// 泊数列 → 月/日 表示
// =======================
function convertNightsFormat(matrix) {
  const header = matrix[0];
  const nightCol = header.indexOf("泊数");
  if (nightCol === -1) return matrix;

  const newMatrix = matrix.map((row, i) => {
    if (i === 0) return row;

    let raw = row[nightCol];

    if (raw === 0 || raw === "0" || raw == null || raw === "") {
      row[nightCol] = "";
      return row;
    }

    if (raw instanceof Date) {
      const mm = raw.getMonth() + 1;
      const dd = raw.getDate();
      row[nightCol] = `${mm}/${dd}`;
      return row;
    }

    let s = String(raw).trim();
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      row[nightCol] = `${d.getMonth() + 1}/${d.getDate()}`;
      return row;
    }

    row[nightCol] = s;
    return row;
  });

  return newMatrix;
}

// =======================
// No列の短縮
// =======================
function transformNoColumn(matrix) {
  if (!matrix || matrix.length === 0) return matrix;

  for (let r = 1; r < matrix.length; r++) {
    let raw = matrix[r][0];
    if (!raw) continue;

    let s = String(raw).trim();
    if (s === "▲" || s === "▼") {
      matrix[r][0] = s;
      continue;
    }

    if (/^[A-Za-z]/.test(s)) {
      matrix[r][0] = s.slice(0, 2);
      continue;
    }

    const beforeParen = s.split("(")[0];
    let m = beforeParen.match(/\d{2,3}/);
    if (m) {
      matrix[r][0] = m[0];
      continue;
    }

    m = s.match(/^\d{1,3}/);
    if (m) {
      matrix[r][0] = m[0];
      continue;
    }

    matrix[r][0] = s.slice(0, 2);
  }
  return matrix;
}

// =======================
// 列追加
// =======================
function addExtraColumns(matrix) {
  const header = matrix[0];

  const idxProduct = header.indexOf("商品名");
  const idxMemo    = header.indexOf("MEMO");

  if (idxProduct === -1 || idxMemo === -1) {
    alert("商品名または MEMO の列が見つかりません。ヘッダー名を確認してください。");
    return matrix;
  }

  return matrix.map((row, r) => {
    // 左端に空白列追加（ステータス列）
    row = ["", ...row];

    const newIdxMemo    = idxMemo + 1;
    const idxDinner = 9; //夕食列10列目

    // 夕食の右に 3 列追加
    row.splice(idxDinner + 1, 0, "", "", "");
    if (r === 0) {
      row[idxDinner + 1] = "客室係";
      row[idxDinner + 2] = "夕食時間";
      row[idxDinner + 3] = "朝食時間";
    }

    // MEMO の右に 1 列追加（商品名右に3列足したぶん +3）
    row.splice(newIdxMemo + 1 + 3, 0, "");
    if (r === 0) {
      row[newIdxMemo + 1 + 3] = "連絡事項";
    }

    return row;
  });
}

// =======================
// テーブル描画
// =======================
function renderTable(
  matrix,
  rowColors = [],       // U列の色（行ごと）
  rowStatuses = [],     // V列のステータス
  rowDinners = [],      // W列の夕食
  rowBreakfasts = [],   // X列の朝食
  rowNotes = [],        // Y列の連絡事項テキスト
  rowNoteColors = [],   // Y列の連絡事項文字色キー（black/red/blue）
  rowStaffs = []
) {
  const table = document.createElement("table");

  // header
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  matrix[0].forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  // 各列インデックス
  const header = matrix[0];
  let noColIndex       = header.indexOf("No");
  let nameColIndex     = header.indexOf("氏名");
  const maleColIndex      = header.indexOf("男");
  const femaleColIndex    = header.indexOf("女");
  const childColIndex     = header.indexOf("子供");
  const infantColIndex    = header.indexOf("幼児");
  const nightsColIndex     = header.indexOf("泊数");
  const stayplanColIndex = header.indexOf("商品名");
  const staffColIndex     = header.indexOf("客室係");
  const dinnerColIndex    = header.indexOf("夕食時間");
  const breakfastColIndex = header.indexOf("朝食時間");
  const noteColIndex      = header.indexOf("連絡事項");

  if (noColIndex === -1)   noColIndex = 1; // 0:ステータス,1:No
  if (nameColIndex === -1) nameColIndex = 2; // 2:氏名

  const tbody = document.createElement("tbody");

  for (let r = 1; r < matrix.length; r++) {
    const tr = document.createElement("tr");

    const sheetRowIndex = r + 1; // シート上の行番号（ヘッダーが1行目）

    const rowColor      = rowColors[r - 1]      || '';
    const rowStatus     = rowStatuses[r - 1]    || '';
    const rowDinner     = rowDinners[r - 1]     || '';
    const rowBreakfast  = rowBreakfasts[r - 1]  || '';
    const rowNoteText   = rowNotes[r - 1]       || '';
    const rowNoteColor  = rowNoteColors[r - 1]  || 'black';
    const rowStaff      = rowStaffs[r - 1]      || '';

    // この行の No / 氏名 / 各人数（元データから）
    const rowData   = matrix[r];
    const rowNoText   = String(rowData[noColIndex]   ?? '');
    const rowNameText = String(rowData[nameColIndex] ?? '');
    const rowMale   = (maleColIndex   !== -1 && rowData[maleColIndex])   ? Number(rowData[maleColIndex])   || 0 : 0;
    const rowFemale = (femaleColIndex !== -1 && rowData[femaleColIndex]) ? Number(rowData[femaleColIndex]) || 0 : 0;
    const rowChild  = (childColIndex  !== -1 && rowData[childColIndex])  ? Number(rowData[childColIndex])  || 0 : 0;
    const rowInfant = (infantColIndex !== -1 && rowData[infantColIndex]) ? Number(rowData[infantColIndex]) || 0 : 0;

    // C/I 行の背景色を反映（No/氏名のセル色はこの上に重ねる）
    updateRowStatusBackground(tr, rowStatus);

    rowData.forEach((cell, cIndex) => {
      const td = document.createElement("td");
      td.innerHTML = escapeHtml(String(cell ?? '')).replace(/\n/g, "<br>");

      // 1列目（ステータス）
      if (cIndex === 0) {
        td.textContent = rowStatus || "";
        td.style.cursor = 'pointer';
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showStatusPicker(td, sheetRowIndex);
        });
      }

      // ★ 氏名列の語尾に「様」を付ける
      if (cIndex === nameColIndex) {
        const rawName = String(cell ?? '').trim();   // 元データの生文字
        if (rawName !== '') {                        // 空白チェック
          td.innerHTML += '<span style="font-size:0.6em;font-weight:400;"> 様</span>';
        }
      }


      // No & 氏名：色適用 + カラーピッカー
      if (cIndex === noColIndex || cIndex === nameColIndex) {
        if (rowColor) {
          td.style.backgroundColor = rowColor; // 行背景よりセル色を優先
        }
        td.dataset.rowIndex = String(sheetRowIndex);
        td.dataset.colorTarget = "1";
        td.classList.add('no-cell');
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showColorPicker(td, sheetRowIndex);
        });
      }

      // 男・女・子供・幼児 → 人数変更モーダル
      const isGuestCol =
        (maleColIndex   !== -1 && cIndex === maleColIndex) ||
        (femaleColIndex !== -1 && cIndex === femaleColIndex) ||
        (childColIndex  !== -1 && cIndex === childColIndex) ||
        (infantColIndex !== -1 && cIndex === infantColIndex);

      if (isGuestCol) {
        td.style.cursor = 'pointer';
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showGuestModal(
            sheetRowIndex,
            rowNoText,
            rowNameText,
            rowMale,
            rowFemale,
            rowChild,
            rowInfant
          );
        });
      }

      // ★ 泊数列："1/1" 以外なら赤字
      if (nightsColIndex !== -1 && cIndex === nightsColIndex) {
        const text = td.textContent.trim();
        if (text !== "" && text !== "1/1") {
          td.className = "highlight-yellow";
        }
      }

      // 商品名が「部屋食」なら中央帯ハイライト
      if (stayplanColIndex !== -1 && cIndex === stayplanColIndex) {
        const text = td.textContent.trim();
        if (text === "部屋食") {
          td.classList.add('highlight-green');  
        }
      }

      // 夕食
      if (cIndex === dinnerColIndex) {
        td.textContent = rowDinner || "";
        td.style.cursor = 'pointer';
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showDinnerPicker(td, sheetRowIndex, rowDinner);
        });
      }

      // 朝食
      if (cIndex === breakfastColIndex) {
        td.textContent = rowBreakfast || "";
        td.style.cursor = 'pointer';
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showBreakfastPicker(td, sheetRowIndex, rowBreakfast);
        });
      }

      // 連絡事項
      if (cIndex === noteColIndex) {
        td.textContent = rowNoteText || "";
        td.style.cursor = 'pointer';
        td.style.color =
          NOTE_COLOR_MAP[rowNoteColor] || NOTE_COLOR_MAP.black;

        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showNoteEditor(td, sheetRowIndex, rowNoteText, rowNoteColor);
        });
      }

      // 客室係
      if (cIndex === staffColIndex) {
        td.textContent = rowStaff || "";
        td.style.cursor = 'pointer';
        td.addEventListener('click', (event) => {
          event.stopPropagation();
          showStaffPicker(td, sheetRowIndex, rowStaff);
        });
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableContainer.innerHTML = "";
  tableContainer.appendChild(table);
}




// =======================
// HTML エスケープ
// =======================
function escapeHtml(str) {
  return str.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
}
