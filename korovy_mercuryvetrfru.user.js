// ==UserScript==
// @name         korovy_mercuryvetrfru
// @namespace    mercuryvetrfru
// @version      0.1
// @description  Автоматический поиск и добавление животных из файла Excel в ИС Хорриот
// @author       Vit O
// @license      MIT
// @match        https://mercury.vetrf.ru/gve/operatorui*
// @run-at       document-end
// @grant        none
// @unwrap
// ==/UserScript==
 
(function () {
  'use strict';
  var CFG = { pauseMs: 2000, searchTimeoutMs: 60000, addTimeoutMs: 90000, maxPages: 100, pagePauseMs: 800 };
  var state = { animals: [], log: [], running: false, stopRequested: false, index: 0, ok: 0, fail: 0, skip: 0, xlsxLoaded: false };
 
  // Автоматический перебор CDN для 100% надежной загрузки XLSX
  function loadXLSX() {
    return new Promise(function (resolve, reject) {
      if (typeof XLSX !== 'undefined') {
        state.xlsxLoaded = true;
        return resolve();
      }
      var sources = [
        'https://cdn.jsdelivr.net/npm/xlsx@0.20.3/dist/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.20.3/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.20.3/dist/xlsx.full.min.js'
      ];
      var index = 0;
      function tryNext() {
        if (index >= sources.length) {
          return reject(new Error('Не удалось загрузить библиотеку XLSX ни с одного CDN. Проверьте интернет.'));
        }
        var url = sources[index++];
        var script = document.createElement('script');
        script.src = url;
        script.async = true;
        script.onload = function () {
          if (typeof XLSX !== 'undefined') {
            state.xlsxLoaded = true;
            resolve();
          } else {
            tryNext();
          }
        };
        script.onerror = function () {
          tryNext();
        };
        document.head.appendChild(script);
      }
      tryNext();
    });
  }
 
  function humanSleep(min, max) {
    if (max === undefined) { max = min * 1.25; min = min * 0.75; }
    var ms = Math.floor(Math.random() * (max - min + 1)) + min;
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }
 
  async function humanClick(el) {
    if (!el) return;
    if (el && el.jquery) el = el[0];
    await humanSleep(150, 450);
    el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    await humanSleep(60, 180);
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  }
 
  async function humanType(input, text) {
    if (!input) return;
    input.focus();
    await humanClick(input);
    input.value = '';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await humanSleep(200, 500);
    for (var i = 0; i < text.length; i++) {
      var char = text[i];
      input.dispatchEvent(new KeyboardEvent('keydown', { key: char, bubbles: true }));
      input.value += char;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keyup', { key: char, bubbles: true }));
      await humanSleep(50, 160);
    }
    await humanSleep(100, 300);
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.blur();
  }
 
  function norm(s) {
    return String(s == null ? '' : s).replace(/\u00a0/g, ' ').trim().toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ');
  }
 
  function cellText(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') {
      if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
      return String(v);
    }
    return String(v).trim();
  }
 
  function updateLiveLog(number, status) {
    var container = document.getElementById('haa-live-log');
    if (!container) return;
    var timeStr = new Date().toLocaleTimeString('ru-RU');
    var color = '#333';
    if (status === 'Добавлено') color = '#2e7d32';
    else if (status.indexOf('Уже') !== -1 || status.indexOf('стоит') !== -1) color = '#e65100';
    else if (status.indexOf('Ошибка') !== -1 || status.indexOf('Не найдено') !== -1 || status.indexOf('Кличка не совпала') !== -1) color = '#c62828';
 
    var div = document.createElement('div');
    div.style.color = color;
    div.style.marginBottom = '2px';
    div.textContent = '[' + timeStr + '] #' + number + ': ' + status;
 
    container.appendChild(div);
    while (container.children.length > 15) {
      container.removeChild(container.firstChild);
    }
    container.scrollTop = container.scrollHeight;
  }
 
  function logStatus(item, status, details) {
    state.log.push({ 'Внутрихозяйственный номер': item.number, 'Кличка (из Excel)': item.name, Статус: status, Детали: details || '', Время: new Date().toLocaleString('ru-RU') });
    updateLiveLog(item.number, status);
  }
 
  function setStatus(text) {
    var el = document.getElementById('haa-status');
    if (el) el.textContent = text;
  }
 
  function setProgress() {
    var el = document.getElementById('haa-progress');
    if (!el) return;
    el.textContent = 'Обработано: ' + state.index + ' / ' + state.animals.length + ' | ОК: ' + state.ok + ' | пропуск: ' + state.skip + ' | ошибки: ' + state.fail;
  }
 
  async function parseExcel(file) {
    await loadXLSX();
    var buf = await file.arrayBuffer();
    if (typeof XLSX === 'undefined') throw new Error('Библиотека XLSX не загружена.');
    var wb = XLSX.read(buf, { type: 'array' });
    var sheetName = wb.SheetNames[0];
    var sheet = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    if (!rows.length) throw new Error('Первый лист пустой: ' + sheetName);
 
    var animals = [];
    rows.forEach(function (row, i) {
      var number = cellText(row['Внутрихозяйственный номер']) || cellText(row['№']) || cellText(row['Номер']);
      var name = cellText(row['Кличка']);
      if (!number) return;
      animals.push({ row: i + 2, number: number, name: name, sheet: sheetName });
    });
 
    if (!animals.length) {
      throw new Error('На первом листе («' + sheetName + '») не найдено ни одной строки с внутрихозяйственным номером.');
    }
    return { sheetName: sheetName, animals: animals };
  }
 
  function downloadLog() {
    if (typeof XLSX === 'undefined') { alert('Библиотека XLSX не доступна.'); return; }
    var rows = state.log.length ? state.log : [{ Статус: 'Лог пуст', Время: new Date().toLocaleString('ru-RU') }];
    var ws = XLSX.utils.json_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'лог');
    var name = 'лог_хорриот_' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.xlsx';
    XLSX.writeFile(wb, name);
  }
 
  async function ensureDialogOpen() {
    if (typeof $ === 'undefined') throw new Error('jQuery не загружен на странице.');
    var dialog = document.getElementById('uan-search-dialog');
    var isOpen = dialog && $(dialog).hasClass('ui-dialog-content') && $(dialog).dialog('isOpen');
    if (!isOpen) {
      var link = document.querySelector('a.herriot-add.animal-list-action-link');
      if (!link) throw new Error('Не найдена кнопка «Добавить из ИС "Хорриот"».');
      await humanClick(link);
    }
    return waitUntil(function () {
      var d = document.getElementById('uan-search-dialog');
      return d && $(d).dialog('isOpen');
    }, 15000, 'Не открылся диалог Хорриот');
  }
 
  function waitUntil(fn, timeoutMs, errMsg) {
    var start = Date.now();
    return new Promise(function (resolve, reject) {
      (function tick() {
        try { if (fn()) return resolve(); } catch (e) {}
        if (Date.now() - start > timeoutMs) return reject(new Error(errMsg || 'Таймаут ожидания'));
        setTimeout(tick, 200);
      })();
    });
  }
 
  function waitAjaxIdle(timeoutMs) {
    return waitUntil(function () {
      var isUpdating = (typeof IS_UPDATING !== 'undefined') ? IS_UPDATING : false;
      var ajaxActive = window.ajaxVetcontrol && window.ajaxVetcontrol.active > 0;
      return !ajaxActive && !isUpdating;
    }, timeoutMs || CFG.searchTimeoutMs, 'Таймаут AJAX');
  }
 
  async function clearFilters() {
    var ids = ['herriot-marking-number', 'herriot-number', 'herriot-name'];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (el && el.value) {
        el.value = '';
        if (typeof $ !== 'undefined') $(el).trigger('input').trigger('change');
      }
    }
    if (typeof $ !== 'undefined') $('#herriot-type-filter input[type=radio]').prop('checked', false);
    await humanSleep(100, 300);
  }
 
  async function searchByMarkingNumber(number) {
    await clearFilters();
    var input = document.getElementById('herriot-marking-number');
    if (!input) throw new Error('Не найдено поле «Номера средств маркирования».');
    await humanType(input, number);
    var searchBtn = document.getElementById('herriot-search-button') || document.getElementById('searchInHorriot') || document.getElementById('doSearch');
    if (!searchBtn) {
      var dialog = document.getElementById('uan-search-dialog');
      if (dialog) {
        var buttons = dialog.querySelectorAll('button');
        for (var i = 0; i < buttons.length; i++) {
          if (buttons[i].textContent.trim().toLowerCase() === 'поиск' || buttons[i].textContent.trim().toLowerCase() === 'найти') {
            searchBtn = buttons[i];
            break;
          }
        }
      }
    }
    if (!searchBtn) searchBtn = document.querySelector('.search-btn') || document.querySelector('.herriot-search-btn') || document.querySelector('button[type="submit"]');
    if (!searchBtn) throw new Error('Не найдена кнопка поиска в диалоге Хорриот.');
    await humanClick(searchBtn);
    await waitAjaxIdle(CFG.searchTimeoutMs);
    await humanSleep(300, 700);
  }
 
  function getResultRows() {
    return Array.prototype.slice.call(document.querySelectorAll('#animalsFromHerriotAjaxList tr.herriot-result-row'));
  }
 
  function rowNickname(tr) {
    var tds = tr.querySelectorAll('td');
    if (tds.length < 6) return '';
    return (tds[5].textContent || '').replace(/\s+/g, ' ').trim();
  }
 
  function rowMarkingsText(tr) {
    var tds = tr.querySelectorAll('td');
    if (tds.length < 4) return '';
    return (tds[3].textContent || '').replace(/\s+/g, ' ').trim();
  }
 
  function pickRow(rows, expectedName) {
    if (!rows.length) return { kind: 'not_found' };
    var want = norm(expectedName);
    var matched = rows.filter(function (tr) { return want && norm(rowNickname(tr)) === want; });
    if (matched.length) return { kind: 'ok', row: matched[0], foundName: rowNickname(matched[0]), multi: matched.length > 1 };
    if (!want) return { kind: 'ok', row: rows[0], foundName: rowNickname(rows[0]), multi: rows.length > 1, warn: 'Кличка в Excel пустая — выбрана первая строка' };
    var foundNames = rows.map(rowNickname).filter(Boolean).filter(function (v, i, a) { return a.indexOf(v) === i; });
    return { kind: 'nickname_mismatch', foundNames: foundNames, sampleMarkings: rows.slice(0, 3).map(rowMarkingsText) };
  }
 
  function getCurrentPageNumber() {
    var input = document.querySelector('#pageNavBlock input[name="pageList"]');
    if (!input) return 1;
    var n = parseInt(input.value, 10);
    return isNaN(n) || n < 1 ? 1 : n;
  }
 
  function getNextPageButton() {
    return document.querySelector('#pageNavBlock button[title="перейти на следующую страницу"]');
  }
 
  function canGoNextPage() {
    var btn = getNextPageButton();
    if (!btn || btn.disabled || btn.classList.contains('ui-state-disabled') || btn.classList.contains('ui-button-disabled')) return false;
    return true;
  }
 
  async function goToNextResultsPage() {
    if (!canGoNextPage()) return false;
    var beforePage = getCurrentPageNumber();
    var beforeFirstId = '';
    var firstRow = document.querySelector('#animalsFromHerriotAjaxList tr.herriot-result-row');
    if (firstRow) beforeFirstId = firstRow.id || '';
    var btn = getNextPageButton();
    await humanClick(btn);
    await waitAjaxIdle(CFG.searchTimeoutMs);
    await humanSleep(CFG.pagePauseMs, CFG.pagePauseMs * 1.5);
    var afterPage = getCurrentPageNumber();
    var afterFirstId = '';
    var afterRow = document.querySelector('#animalsFromHerriotAjaxList tr.herriot-result-row');
    if (afterRow) afterFirstId = afterRow.id || '';
    if (afterPage === beforePage && afterFirstId === beforeFirstId) return false;
    return true;
  }
 
  function uniqPush(arr, values) {
    (values || []).forEach(function (v) { if (v && arr.indexOf(v) === -1) arr.push(v); });
  }
 
  async function findAcrossAllPages(expectedName, onPage) {
    var allFoundNames = [], sampleMarkings = [], totalRows = 0, pagesChecked = 0;
    for (var p = 0; p < CFG.maxPages; p++) {
      if (state.stopRequested) return { kind: 'stopped', pagesChecked: pagesChecked };
      pagesChecked++;
      var pageNo = getCurrentPageNumber();
      if (typeof onPage === 'function') onPage(pageNo, pagesChecked);
      var rows = getResultRows();
      totalRows += rows.length;
      var emptyMsg = document.querySelector('#animalsFromHerriotAjaxList .errorMessage');
      if (emptyMsg && !rows.length && pagesChecked === 1) return { kind: 'not_found', details: emptyMsg.textContent.trim() || 'Нет данных', pagesChecked: pagesChecked };
      var pick = pickRow(rows, expectedName);
      if (pick.kind === 'ok') {
        pick.pagesChecked = pagesChecked;
        pick.pageNo = pageNo;
        if (pagesChecked > 1) pick.warn = (pick.warn ? pick.warn + '; ' : '') + 'найдено на странице ' + pageNo;
        return pick;
      }
      if (pick.kind === 'nickname_mismatch') {
        uniqPush(allFoundNames, pick.foundNames);
        if (sampleMarkings.length < 5) uniqPush(sampleMarkings, pick.sampleMarkings);
      }
      var moved = await goToNextResultsPage();
      if (!moved) break;
    }
    if (!totalRows) return { kind: 'not_found', details: 'Поиск ничего не вернул', pagesChecked: pagesChecked };
    return { kind: 'nickname_mismatch', foundNames: allFoundNames, sampleMarkings: sampleMarkings.slice(0, 5), pagesChecked: pagesChecked };
  }
 
  function isAlreadyInInventory(guid) {
    if (!guid) return false;
    if (typeof HERRIOT_ADDED !== 'undefined' && HERRIOT_ADDED[guid] === true) return true;
    return !!document.querySelector('#tbodyAnimalList tr[data-guid="' + guid + '"]');
  }
 
  async function clickAddAndWait() {
    var btn = document.getElementById('herriot-search-btn');
    if (!btn) throw new Error('Не найдена кнопка «Добавить».');
    if (btn.disabled) throw new Error('Кнопка «Добавить» неактивна.');
    var beforeCount = document.querySelectorAll('#tbodyAnimalList tr.animal-list-row').length;
    await humanClick(btn);
    return waitUntil(function () {
      if (typeof IS_UPDATING !== 'undefined' && IS_UPDATING) return false;
      var dialog = document.getElementById('uan-search-dialog');
      var closed = dialog && typeof $ !== 'undefined' && $(dialog).hasClass('ui-dialog-content') && !$(dialog).dialog('isOpen');
      var msg = document.getElementById('herriot-add-message');
      var afterCount = document.querySelectorAll('#tbodyAnimalList tr.animal-list-row').length;
      return closed || !!msg || afterCount > beforeCount;
    }, CFG.addTimeoutMs, 'Таймаут добавления животного').then(function () { return humanSleep(400, 800); });
  }
 
  async function processOne(item) {
    await ensureDialogOpen();
    var keeping = document.getElementById('herriot-keeping-place-select');
    if (keeping && (keeping.value === 'null' || !keeping.value)) throw new Error('Не выбрано место содержания в диалоге Хорриот.');
    await searchByMarkingNumber(item.number);
    var pick = await findAcrossAllPages(item.name, function (pageNo, pagesChecked) {
      setStatus('Ищу: ' + item.number + ' / ' + (item.name || '—') + ' (стр. ' + pageNo + ', просмотрено: ' + pagesChecked + ')');
    });
    if (pick.kind === 'stopped') { state.skip++; logStatus(item, 'Остановлено', 'Остановка во время просмотра страниц'); return; }
    if (pick.kind === 'not_found') { state.fail++; logStatus(item, 'Не найдено', (pick.details || 'Поиск ничего не вернул') + (pick.pagesChecked ? '; страниц просмотрено: ' + pick.pagesChecked : '')); return; }
    if (pick.kind === 'nickname_mismatch') {
      state.fail++;
      logStatus(item, 'Кличка не совпала', 'Найдены клички: ' + (pick.foundNames.length ? pick.foundNames.join(', ') : '(пусто)') + (pick.sampleMarkings && pick.sampleMarkings.length ? '; маркировки: ' + pick.sampleMarkings.join(' | ') : '') + (pick.pagesChecked ? '; страниц просмотрено: ' + pick.pagesChecked : ''));
      return;
    }
    var tr = pick.row;
    var checkbox = tr.querySelector('input.herriot-checkbox');
    if (!checkbox) { state.fail++; logStatus(item, 'Ошибка', 'В строке результата нет чекбокса'); return; }
    var guid = checkbox.value;
    var detailsExtra = [];
    if (pick.multi) detailsExtra.push('найдено несколько строк с совпадающей кличкой — взята первая');
    if (pick.warn) detailsExtra.push(pick.warn);
    if (pick.foundName) detailsExtra.push('кличка в Хорриот: ' + pick.foundName);
    if (isAlreadyInInventory(guid)) { state.skip++; logStatus(item, 'Уже в описи', detailsExtra.concat(['guid: ' + guid]).join('; ')); return; }
    if (checkbox.checked) { state.skip++; logStatus(item, 'Галочка уже стоит', detailsExtra.concat(['guid: ' + guid, 'галочку не снимали']).join('; ')); return; }
    await humanClick(checkbox);
    checkbox.checked = true;
    if (typeof $ !== 'undefined') $(checkbox).trigger('change');
    await clickAddAndWait();
    state.ok++;
    logStatus(item, 'Добавлено', detailsExtra.join('; '));
  }
 
  async function run() {
    if (state.running) return;
    if (!state.animals.length) { alert('Сначала загрузите Excel.'); return; }
    state.running = true; state.stopRequested = false; state.index = 0; state.ok = 0; state.fail = 0; state.skip = 0; state.log = [];
    var liveLogEl = document.getElementById('haa-live-log');
    if (liveLogEl) liveLogEl.innerHTML = '';
    setProgress(); setStatus('Работаю…');
    var pauseInput = document.getElementById('haa-pause');
    CFG.pauseMs = Math.max(500, Number(pauseInput && pauseInput.value) || 2000);
    try {
      for (var i = 0; i < state.animals.length; i++) {
        if (state.stopRequested) { setStatus('Остановлено пользователем.'); break; }
        state.index = i; setProgress();
        var item = state.animals[i];
        setStatus('Ищу: ' + item.number + ' / ' + (item.name || '—'));
        try { await processOne(item); } catch (e) { state.fail++; logStatus(item, 'Ошибка', (e && e.message) || String(e)); console.error(e); }
        state.index = i + 1; setProgress();
        if (i < state.animals.length - 1 && !state.stopRequested) {
          var currentPause = Math.floor(CFG.pauseMs * (0.8 + Math.random() * 0.4));
          setStatus('Пауза ' + currentPause + ' мс…');
          await humanSleep(currentPause, currentPause);
        }
      }
      if (!state.stopRequested) setStatus('Готово. Скачиваю лог…');
      downloadLog();
      setStatus((state.stopRequested ? 'Остановлено. ' : 'Готово. ') + 'Лог скачан. ОК: ' + state.ok + ', пропуск: ' + state.skip + ', ошибки: ' + state.fail);
    } finally {
      state.running = false;
      updateButtons();
    }
  }
 
  function updateButtons() {
    var start = document.getElementById('haa-start');
    var stop = document.getElementById('haa-stop');
    if (start) start.disabled = state.running || !state.animals.length;
    if (stop) stop.disabled = !state.running;
  }
 
  function createUI() {
    if (document.getElementById('horriot-auto-add-panel')) return;
    var panel = document.createElement('div');
    panel.id = 'horriot-auto-add-panel';
 
    var html = '';
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
    html += '<div style="font-weight:700;font-size:12px;">Хорриот ← Excel</div>';
    html += '<a href="https://vk.ru/osievich" target="_blank" title="Связаться с автором" style="display:inline-flex;justify-content:center;align-items:center;width:16px;height:16px;background:#8a6d3b;color:#fff;border-radius:50%;font-size:10px;font-weight:bold;text-decoration:none;">?</a>';
    html += '</div>';
    html += '<div style="margin-bottom:6px;font-size:11px;opacity:.85;">1-й лист: номер + кличка</div>';
    html += '<input id="haa-file" type="file" accept=".xlsx,.xls" style="width:100%;margin-bottom:6px;font-size:11px;">';
    html += '<label style="display:block;margin-bottom:6px;font-size:11px;">Пауза между поисками (мс)';
    html += '<input id="haa-pause" type="number" min="500" step="100" value="2000" style="width:100%;margin-top:2px;box-sizing:border-box;font-size:11px;padding:2px;"></label>';
    html += '<div id="haa-file-info" style="font-size:11px;margin-bottom:6px;color:#1b5e20;"></div>';
    html += '<div style="display:flex;gap:4px;margin-bottom:6px;">';
    html += '<button id="haa-start" type="button" style="flex:1;padding:4px;font-size:11px;">Старт</button>';
    html += '<button id="haa-stop" type="button" style="flex:1;padding:4px;font-size:11px;" disabled>Стоп</button>';
    html += '</div>';
    html += '<button id="haa-log" type="button" style="width:100%;padding:4px;margin-bottom:6px;font-size:11px;">Скачать лог</button>';
    html += '<div id="haa-progress" style="font-size:11px;margin-bottom:4px;"></div>';
    html += '<div id="haa-status" style="font-size:11px;line-height:1.35;margin-bottom:6px;"></div>';
    html += '<div style="font-weight:600;font-size:10px;margin-bottom:2px;color:#555;">Живой лог (последние записи):</div>';
    html += '<div id="haa-live-log" style="height:110px;overflow-y:auto;background:#fff;border:1px solid #dcd6c8;padding:4px;font-size:10px;font-family:monospace;line-height:1.2;text-align:left;"></div>';
 
    panel.innerHTML = html;
 
    Object.assign(panel.style, {
      position: 'fixed', top: '80px', right: '16px', width: '240px', zIndex: '999999',
      background: '#fffef8', border: '1px solid #8a6d3b', borderRadius: '6px',
      boxShadow: '0 4px 16px rgba(0,0,0,.15)', padding: '8px',
      fontFamily: 'Segoe UI, Tahoma, sans-serif', fontSize: '11px', color: '#222'
    });
 
    document.body.appendChild(panel);
 
    document.getElementById('haa-file').addEventListener('change', function (e) {
      var file = e.target.files && e.target.files[0];
      if (!file) return;
      setStatus('Загрузка библиотек и чтение Excel…');
      parseExcel(file).then(function (res) {
        state.animals = res.animals;
        document.getElementById('haa-file-info').textContent = 'Лист «' + res.sheetName + '»: ' + res.animals.length + ' жив.';
        setStatus('Файл загружен. Старт.');
        setProgress();
        updateButtons();
      }).catch(function (err) {
        state.animals = [];
        document.getElementById('haa-file-info').textContent = '';
        setStatus((err && err.message) || String(err));
        updateButtons();
      });
    });
 
    document.getElementById('haa-start').addEventListener('click', function () { run(); updateButtons(); });
    document.getElementById('haa-stop').addEventListener('click', function () { state.stopRequested = true; setStatus('Останавливаю…'); });
    document.getElementById('haa-log').addEventListener('click', downloadLog);
 
    updateButtons();
    setStatus('Загрузите Excel и нажмите Старт.');
  }
 
  createUI();
  // Предзагрузка библиотеки в фоне при старте страницы
  loadXLSX().catch(function() {});
 
  setInterval(function() {
    if (!document.getElementById('horriot-auto-add-panel') && document.body) {
      createUI();
    }
  }, 1000);
 
  window.__horriotAutoAdd = {
    destroy: function () {
      state.stopRequested = true;
      var panel = document.getElementById('horriot-auto-add-panel');
      if (panel) panel.remove();
    },
    getState: function () { return state; }
  };
 
  console.log('%cХорриот←Excel: панель справа. Загрузите xlsx и нажмите Старт.', 'color:#1b5e20;font-weight:bold;');
})();
