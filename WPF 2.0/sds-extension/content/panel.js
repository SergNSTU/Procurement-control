/*
 * SDS Помощник — panel.js (content script, isolated world).
 * Панель в shadow DOM: «Помощник распознавания» на странице расчёта
 * (/user/carts/{id}) + «Лог RPC» (диагностика). Мутации — только по кнопке
 * «Применить…» с явным предпросмотром.
 */
(function () {
  'use strict';
  var S = window.SDSX;

  var CSS = [
    ':host{all:initial}',
    '*{box-sizing:border-box;font-family:"Segoe UI",Arial,sans-serif}',
    '.fab{position:fixed;right:0;top:42%;z-index:2147483600;background:#1a56db;color:#fff;',
    'border:none;border-radius:6px 0 0 6px;padding:10px 6px;cursor:pointer;font-size:12px;font-weight:600;',
    'writing-mode:vertical-rl;letter-spacing:1px;box-shadow:0 2px 8px rgba(0,0,0,.35)}',
    '.fab:hover{background:#1244b8}',
    '.drawer{position:fixed;top:0;right:-460px;width:440px;height:100vh;z-index:2147483601;background:#f6f7f9;',
    'box-shadow:-4px 0 18px rgba(0,0,0,.25);transition:right .18s ease;display:flex;flex-direction:column}',
    '.drawer.open{right:0}',
    '.head{background:#1a56db;color:#fff;padding:10px 12px;display:flex;align-items:center;gap:8px}',
    '.head .t{font-size:14px;font-weight:600;flex:1}',
    '.head button{background:rgba(255,255,255,.15);border:none;color:#fff;border-radius:4px;padding:4px 8px;cursor:pointer;font-size:12px}',
    '.head button:hover{background:rgba(255,255,255,.3)}',
    '.tabs{display:flex;background:#e8ebf1;border-bottom:1px solid #d5dae3}',
    '.tabs button{flex:1;border:none;background:none;padding:8px 4px;font-size:12px;cursor:pointer;color:#444}',
    '.tabs button.act{background:#fff;font-weight:600;color:#1a56db;box-shadow:inset 0 -2px 0 #1a56db}',
    '.body{flex:1;overflow:auto;padding:10px 12px;font-size:12.5px;color:#222}',
    '.row{background:#fff;border:1px solid #dde2ea;border-radius:6px;padding:8px;margin-bottom:8px}',
    '.row .nm{font-weight:600;word-break:break-all}',
    '.row .meta{color:#667;font-size:11.5px;margin-top:2px}',
    '.chip{display:inline-block;background:#e5edff;color:#1a56db;border-radius:10px;padding:1px 8px;font-size:11px;margin:2px 4px 0 0}',
    '.chip.warn{background:#fff3cd;color:#8a6d00}',
    '.chip.bad{background:#fde8e8;color:#b91c1c}',
    '.chip.ok{background:#def7ec;color:#046c4e}',
    'button.b{background:#1a56db;color:#fff;border:none;border-radius:4px;padding:6px 10px;font-size:12px;cursor:pointer}',
    'button.b:hover{background:#1244b8}',
    'button.b:disabled{background:#9db4e8;cursor:default}',
    'button.b.ghost{background:#fff;color:#1a56db;border:1px solid #b9c8ee}',
    'button.b.green{background:#0e7a4a}button.b.green:hover{background:#0a5f3a}',
    'button.b.red{background:#b91c1c}',
    '.cands{margin-top:6px;border-top:1px dashed #d5dae3;padding-top:6px}',
    '.cand{display:flex;gap:6px;align-items:flex-start;padding:4px;border-radius:4px;cursor:pointer;font-size:11.5px}',
    '.cand:hover{background:#eef3ff}',
    '.cand.sel{background:#e5edff;outline:1px solid #b9c8ee}',
    '.cand input{margin-top:2px}',
    '.cand .pn{font-weight:600}',
    '.cand .dim{color:#667}',
    '.toolbar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}',
    '.sum{font-size:12px;color:#333;background:#fff;border:1px solid #dde2ea;border-radius:6px;padding:6px 8px;margin-bottom:8px}',
    '.log{font-family:Consolas,monospace;font-size:11px;background:#111;color:#cde3ff;border-radius:6px;padding:8px;margin-top:8px;white-space:pre-wrap;word-break:break-all;max-height:180px;overflow:auto}',
    '.log .err{color:#ff9c9c}',
    '.log .ok{color:#8ff0b6}',
    '.tabpane{display:none}.tabpane.act{display:block}',
    'table.preview{width:100%;border-collapse:collapse;font-size:11.5px;background:#fff}',
    'table.preview th,table.preview td{border:1px solid #dde2ea;padding:4px 6px;text-align:left;word-break:break-all}',
    'table.preview th{background:#eef1f6}',
    '.hint{font-size:11.5px;color:#556;background:#fffceb;border:1px solid #eadfa8;border-radius:6px;padding:6px 8px;margin-bottom:8px}',
    'a.lnk{color:#1a56db;cursor:pointer;text-decoration:underline}'
  ].join('');

  var state = {
    fid: null,
    tab: 'helper',
    analyzing: false,
    rows: [],
    summary: '',
    panelLog: [],       // журнал операций панели
    open: false,
    applying: false,
    previewApply: false,
    structureNoteShown: false
  };

  var host, root, fab, drawer, elTabs = {}, elPanes = {};

  function h(tag, attrs) {
    var node = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (k === 'class') node.className = v;
      else if (k === 'text') node.textContent = v;
      else if (k === 'style') node.setAttribute('style', v);
      else if (k === 'on' && v && typeof v === 'object') {
        Object.keys(v).forEach(function (evt) { node.addEventListener(evt, v[evt]); });
      }
      else if (v == null) { /* пропуск (например, checked: null) */ }
      else node.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) {
      var c = arguments[i];
      if (c == null) continue;
      if (Array.isArray(c)) c.forEach(function (x) {
        if (x == null) return;
        if (typeof x === 'object') node.appendChild(x);
        else node.appendChild(document.createTextNode(String(x)));
      });
      else if (typeof c === 'object') node.appendChild(c);
      else node.appendChild(document.createTextNode(String(c)));
    }
    return node;
  }

  function plog(msg, cls) {
    state.panelLog.push({ t: new Date().toISOString().slice(11, 19), msg: msg, cls: cls || '' });
    if (state.panelLog.length > 200) state.panelLog.shift();
    renderLog();
  }

  function renderLog() {
    if (!elPanes.helper) return;
    var box = elPanes.helper.querySelector('.log');
    if (!box) return;
    box.textContent = '';
    state.panelLog.slice(-60).forEach(function (l) {
      var line = h('div', { class: l.cls || '' }, '[' + l.t + '] ' + l.msg);
      box.appendChild(line);
    });
    box.scrollTop = box.scrollHeight;
  }

  // ---------- Скелет панели ----------
  function buildUI() {
    host = document.createElement('div');
    host.id = 'sdsx-host';
    root = host.attachShadow({ mode: 'open' });
    var style = document.createElement('style');
    style.textContent = CSS;
    root.appendChild(style);

    fab = h('button', { class: 'fab', on: { click: function () { state.open = !state.open; syncOpen(); } } }, 'SDS+');
    drawer = h('div', { class: 'drawer' });

    var head = h('div', { class: 'head' },
      h('div', { class: 't' }, 'SDS Помощник'),
      h('span', { class: 'fid', style: 'opacity:.85;font-size:12px' }, ''),
      h('button', { on: { click: analyze }, title: 'Перечитать расчёт' }, '⟳'),
      h('button', { on: { click: function () { state.open = false; syncOpen(); } } }, '✕'));
    drawer.appendChild(head);
    elTabs.fid = head.querySelector('.fid');

    var tabs = h('div', { class: 'tabs' },
      h('button', { class: 'act', 'data-tab': 'helper', on: { click: function () { setTab('helper'); } } }, 'Помощник'),
      h('button', { 'data-tab': 'rpclog', on: { click: function () { setTab('rpclog'); } } }, 'Лог RPC'),
      h('button', { 'data-tab': 'about', on: { click: function () { setTab('about'); } } }, 'О'));
    drawer.appendChild(tabs);

    var body = h('div', { class: 'body' });
    drawer.appendChild(body);

    // --- Помощник ---
    elPanes.helper = h('div', { class: 'tabpane act', 'data-pane': 'helper' },
      h('div', { class: 'hint' },
        'Панель работает с открытой страницей расчёта. Чтение — сразу, изменения — только по кнопке «Применить…» с предпросмотром. ',
        'Первый шаг: ', h('a', { class: 'lnk', on: { click: analyze } }, '«Анализировать расчёт»'), '.'),
      h('div', { class: 'toolbar' },
        h('button', { class: 'b', on: { click: analyze } }, 'Анализировать расчёт'),
        h('button', { class: 'b ghost', on: { click: searchAll } }, 'Найти кандидатов для всех'),
        h('button', { class: 'b green', on: { click: startApply } }, 'Применить выбранные…'),
        h('button', { class: 'b ghost', on: { click: testParser } }, 'Тест парсера'),
        h('button', { class: 'b ghost', on: { click: copyDataJson } }, 'Копировать JSON расчёта')),
      h('div', { class: 'sum' }, 'Откройте страницу расчёта /user/carts/{id}'),
      h('div', { class: 'rows' }),
      h('div', { class: 'log' }));
    body.appendChild(elPanes.helper);

    // --- Лог RPC ---
    elPanes.rpclog = h('div', { class: 'tabpane', 'data-pane': 'rpclog' },
      h('div', { class: 'hint' },
        'Запись запросов сайта к /api/ включена всегда. ',
        'Чтобы снять контракт ручного «Распознать»: откройте расчёт, нажмите «Распознать» на одной строке пассива, доведите до результата, затем — ', h('b', {}, '«Скопировать UserFolders»'), ' и пришлите лог.'),
      h('div', { class: 'toolbar' },
        h('button', { class: 'b', on: { click: copyUserFolders } }, 'Скопировать UserFolders'),
        h('button', { class: 'b ghost', on: { click: copyAllLog } }, 'Скопировать всё'),
        h('button', { class: 'b ghost', on: { click: function () { S.log.length = 0; renderRpcLog(); plog('лог RPC очищен'); } } }, 'Очистить'),
        h('button', { class: 'b ghost', on: { click: renderRpcLog } }, 'Обновить')),
      h('div', { class: 'rpclist', style: 'font-family:Consolas,monospace;font-size:11px' }, 'пусто'));
    body.appendChild(elPanes.rpclog);

    // --- О ---
    elPanes.about = h('div', { class: 'tabpane', 'data-pane': 'about' },
      h('div', { class: 'row' },
        h('div', { class: 'nm' }, 'SDS Помощник v0.4.0'),
        h('div', { class: 'meta' }, 'Расширение для sds.compel.ru. Не связано с другими приложениями; работает только внутри открытой вами сессии сайта.'),
        h('div', { class: 'meta' }, 'Читает: UserFolders/GetData, Search2/Load2, Search2/GetClasses. Меняет (по кнопке): UserFolders/SaveData (выбор кандидата в строке).'),
        h('div', { class: 'meta' }, 'API-версия берётся из data-ver страницы: '),
        h('div', { class: 'meta ver' }, '…')),
      h('div', { class: 'row' },
        h('div', { class: 'nm' }, 'Приватность'),
        h('div', { class: 'meta' }, 'Все данные остаются в браузере. Запросы идут только к sds.compel.ru.')));
    body.appendChild(elPanes.about);
    root.appendChild(fab);
    root.appendChild(drawer);
    (document.body || document.documentElement).appendChild(host);
  }

  function syncOpen() {
    if (!drawer) return;
    drawer.classList.toggle('open', state.open);
    fab.style.display = state.open ? 'none' : '';
    if (state.open && state.tab === 'rpclog') renderRpcLog();
  }

  function setTab(name) {
    state.tab = name;
    root.querySelectorAll('.tabs button').forEach(function (b) {
      b.classList.toggle('act', b.getAttribute('data-tab') === name);
    });
    Object.keys(elPanes).forEach(function (k) {
      elPanes[k].classList.toggle('act', k === name);
    });
    if (name === 'rpclog') renderRpcLog();
    if (name === 'about') {
      var v = elPanes.about.querySelector('.ver');
      if (v) v.textContent = 'data-ver = ' + (S.apiVer ? S.apiVer() : '?');
    }
  }

  function copyText(text, what) {
    var done = function () { plog('скопировано в буфер: ' + what, 'ok'); };
    var fallback = function () {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed'; ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); done(); } catch (e) { plog('не удалось скопировать: ' + e, 'err'); }
      ta.remove();
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else fallback();
  }

  function copyAllLog() { copyText(JSON.stringify(S.log, null, 1), 'весь лог (' + S.log.length + ')'); }

  function copyUserFolders() {
    var sel = S.log.filter(function (l) { return /UserFolders|UserCarts|Search2/i.test(l.url); });
    if (!sel.length) { plog('в логе нет UserFolders/UserCarts/Search2 запросов', 'err'); return; }
    copyText(JSON.stringify(sel, null, 1), 'UserFolders-лог (' + sel.length + ')');
  }

  function renderRpcLog() {
    if (!elPanes.rpclog) return;
    var list = elPanes.rpclog.querySelector('.rpclist');
    list.textContent = '';
    if (!S.log.length) { list.textContent = 'пусто'; return; }
    S.log.slice(-80).reverse().forEach(function (l) {
      var short = l.url.replace(/^https?:\/\/[^/]+/, '');
      if (short.length > 70) short = short.slice(0, 70) + '…';
      var col = l.st >= 200 && l.st < 300 ? '#8ff0b6' : (l.st ? '#ff9c9c' : '#ccc');
      list.appendChild(h('div', { style: 'padding:2px 0;border-bottom:1px solid #222' },
        h('span', { style: 'color:' + col }, String(l.st || '—')),
        ' ' + l.m + ' ' + short + '  ' + (l.t || '')));
    });
  }

  // ---------- Анализ расчёта ----------
  function analyze() {
    var fid = S.currentFid();
    if (!fid) { plog('страница расчёта не открыта (/user/carts/{id})', 'err'); return; }
    state.fid = fid;
    elTabs.fid.textContent = 'расчёт ' + fid;
    if (state.analyzing) return;
    state.analyzing = true;
    plog('анализ расчёта ' + fid + '…');
    setSummary('Читаю расчёт…');
    S.api.getData(fid, 0).then(function (data) {
      state.analyzing = false;
      var ex = S.extractRows(data);
      state.rawData = data;
      if (!ex.groups.length) {
        setSummary('Структура GetData не распознана — нажмите «Копировать JSON расчёта» и пришлите файл, добавлю поддержку.');
        plog('extractRows: групп не найдено', 'err');
        renderRows();
        return;
      }
      state.rows = [];
      var seenRow = {};
      var emptyCount = 0;
      ex.groups.forEach(function (r) {
        if (!r.name) return;
        // требующие внимания: без позиций ИЛИ «проверить соответствие»
        var needsCheck = !r.hasItems || r.status === 'checkConformity';
        if (!needsCheck) return;
        var key = String(r.groupId) + '/' + String(r.rowId);
        if (seenRow[key]) return;
        seenRow[key] = 1;
        var row = {
          groupId: r.groupId, rowId: r.rowId, name: r.name, qty: r.qty,
          sdsStatus: r.status, reason: r.reason,
          parsed: S.parsePassive(r.name),
          candidates: null, selected: null, include: false,
          status: 'pending', message: ''
        };
        if (r.items && r.items.length) {
          // кандидаты уже найдены SDS — показываем их сразу
          row.candidates = r.items.map(function (it) {
            return {
              pnid: it.pnid || it.pn || '', hash: it.hash, brand: it.brand || '',
              pn: it.pn || it.pnid || '', sim: it.sim,
              package: it.package || null, scode: it.scode != null ? it.scode : '',
              sType: it.sType || 'PN', sOrigin: it.sOrigin || '', stdn: it.stdn || null, open: 0
            };
          });
          row.selected = 0;
          for (var ci = 0; ci < row.candidates.length; ci++) {
            if (r.selected && row.candidates[ci].hash === r.selected) { row.selected = ci; break; }
          }
          row.status = 'found';
          row.include = r.status === 'checkConformity';
        } else {
          emptyCount++;
        }
        state.rows.push(row);
      });
      var total = ex.groups.length;
      var attention = state.rows.length;
      var parsedCount = state.rows.filter(function (r) { return r.parsed.type; }).length;
      setSummary('Строк в расчёте: ' + total + ', требуют внимания: ' + attention +
        ' (без позиций: ' + emptyCount + ', пассив распознан парсером: ' + parsedCount + ')' +
        (attention ? '' : ' — всё распознано, хорошо'));
      plog('анализ готов: строк ' + total + ', требуют внимания ' + attention);
      state.rows.forEach(function (row) {
        if (row.status === 'found') return; // кандидаты уже из SDS
        if (row.parsed.type) { row.status = 'parsed'; row.include = row.parsed.confidence >= 0.65; }
        else row.status = 'unparsed';
      });
      renderRows();
    }).catch(function (e) {
      state.analyzing = false;
      setSummary('Ошибка чтения расчёта: ' + e.message);
      plog('getData: ' + e.message + (e.appError ? ' [' + e.appError + ']' : ''), 'err');
    });
  }

  function setSummary(t) {
    var s = elPanes.helper.querySelector('.sum');
    if (s) s.textContent = t;
  }

  function candLabel(c) {
    var bits = [];
    if (c.brand) bits.push(c.brand);
    if (c.pn || c.pnid) bits.push(c.pn || c.pnid);
    if (c.sim != null) bits.push('совпадение ' + c.sim + '%');
    var price = c.centrePrice != null ? c.centrePrice : c.dmsPrice;
    var qnt = c.centreQnt != null ? c.centreQnt : c.dmsQnt;
    var dlv = c.dmsDlv != null ? c.dmsDlv : '';
    if (qnt !== undefined && qnt !== null && qnt !== '') bits.push('склад ' + qnt);
    if (price !== undefined && price !== null && price !== '') bits.push('$' + price);
    if (dlv !== '') bits.push('срок ' + dlv);
    if (c.mpq != null) bits.push('MPQ ' + c.mpq);
    return { pn: (c.brand ? c.brand + ' ' : '') + (c.pn || c.pnid || '?'), rest: bits.join(' · ') };
  }

  function sdsStatusChip(row) {
    if (row.status === 'applying') return h('span', { class: 'chip' }, 'применяю…');
    if (row.status === 'applied') return h('span', { class: 'chip ok' }, 'распознано ✓');
    if (row.status === 'searching') return h('span', { class: 'chip' }, 'ищу…');
    if (row.status === 'notfound') return h('span', { class: 'chip warn' }, 'ничего не найдено');
    if (row.status === 'error') return h('span', { class: 'chip bad' }, 'ошибка: ' + row.message);
    if (row.sdsStatus === 'checkConformity') return h('span', { class: 'chip warn' }, 'проверить соответствие');
    if (row.sdsStatus === 'emptyStocks') return h('span', { class: 'chip warn' }, 'нет остатков');
    if (row.sdsStatus === 'offerSelected') return h('span', { class: 'chip ok' }, 'выбрано');
    return null;
  }

  function renderRows() {
    var cont = elPanes.helper.querySelector('.rows');
    cont.textContent = '';
    if (!state.rows.length) return;
    state.rows.forEach(function (row, i) {
      var p = row.parsed;
      var chips = [];
      if (p.type === 'cap') chips.push(h('span', { class: 'chip' }, 'конденсатор'));
      else if (p.type === 'res') chips.push(h('span', { class: 'chip' }, 'резистор'));
      else chips.push(h('span', { class: 'chip warn' }, 'не похоже на пассив'));
      if (p.package) chips.push(h('span', { class: 'chip' }, p.package));
      if (p.value) chips.push(h('span', { class: 'chip' }, p.value));
      if (p.voltage) chips.push(h('span', { class: 'chip' }, p.voltage));
      if (p.dielectric) chips.push(h('span', { class: 'chip' }, p.dielectric));
      if (p.tolerance) chips.push(h('span', { class: 'chip' }, '±' + p.tolerance));
      if (p.power) chips.push(h('span', { class: 'chip' }, p.power));

      var statusChip = sdsStatusChip(row);
      if (row.reason) {
        statusChip = statusChip ? [statusChip, h('span', { class: 'chip bad' }, 'SDS: ' + row.reason)]
                                : h('span', { class: 'chip bad' }, 'SDS: ' + row.reason);
      }

      var chipList = chips.slice();
      if (statusChip) {
        if (Array.isArray(statusChip)) chipList = chipList.concat(statusChip);
        else chipList.push(statusChip);
      }

      var header = h('div', null,
        h('label', { style: 'display:flex;gap:6px;align-items:flex-start;cursor:pointer' },
          h('input', { type: 'checkbox', checked: row.include ? 'checked' : null, on: { change: function (ev) { row.include = ev.target.checked; } } }),
          h('div', null,
            h('div', { class: 'nm' }, row.name),
            h('div', { class: 'meta' }, 'группа ' + row.groupId + ' · строка ' + row.rowId + ' · кол-во: ' + (row.qty == null ? '?' : row.qty)))),
        h('div', { style: 'margin-top:4px' }, chipList));

      var box = h('div', { class: 'row', 'data-i': String(i) }, header);

      // Кнопки поиска по строке
      var actions = h('div', { style: 'margin-top:6px;display:flex;gap:6px' },
        h('button', { class: 'b ghost', on: { click: function () { searchRow(i, false); } } }, p.type ? 'Найти по параметрам' : 'Найти как есть'));
      if (p.type && row.name && p.query && p.query !== row.name) {
        actions.appendChild(h('button', { class: 'b ghost', on: { click: function () { searchRow(i, true); } } }, 'Искать исходную строку'));
      }
      box.appendChild(actions);

      if (row.candidates) {
        var cd = h('div', { class: 'cands' });
        row.candidates.forEach(function (c, ci) {
          var lb = candLabel(c);
          var input = h('input', { type: 'radio', name: 'cand' + i });
          if (row.selected === ci) input.checked = true;
          input.addEventListener('change', function () { row.selected = ci; row.include = true; renderRows(); });
          var info = h('div', null,
            h('div', { class: 'pn' }, lb.pn),
            h('div', { class: 'dim' }, lb.rest || '—'),
            h('div', { class: 'dim', style: 'font-size:10.5px;color:#99a' }, 'pnid: ' + (c.pnid || '?')));
          info.appendChild(h('div', { class: 'dim' }, 'hash: ' + (c.hash ? String(c.hash).slice(0, 12) + '…' : '?')));
          var cand = h('div', { class: 'cand' + (row.selected === ci ? ' sel' : '') }, input, info);
          cd.appendChild(cand);
        });
        if (!row.candidates.length) cd.appendChild(h('div', { class: 'dim' }, '— пусто —'));
        box.appendChild(cd);
      }
      cont.appendChild(box);
    });
  }

  function partsToCandidates(parts) {
    return (parts || []).slice(0, 6).map(function (p) {
      return {
        pnid: p.pnid, hash: p.hash, brand: p.brand, pn: p.pn || p.pnid,
        korpus: p.korpus || p.package,
        centrePrice: p.centrePrice, centreQnt: p.centreQnt,
        dmsPrice: p.dmsPrice, dmsQnt: p.dmsQnt, dmsDlv: p.dmsDlv, mpq: p.mpq,
        package: p.package || p.korpus || null, scode: p.scode != null ? p.scode : '',
        sim: p.sim, sType: p.sType || 'PN', sOrigin: p.sOrigin || '', stdn: p.stdn || null, open: 0
      };
    });
  }

  function searchRow(i, useRaw) {
    var row = state.rows[i];
    if (!row) return;
    var q = useRaw ? row.name : (row.parsed.query || row.name);
    row.status = 'searching';
    renderRows();
    S.api.search(q).then(function (res) {
      row.candidates = partsToCandidates(res && res.parts);
      row.selected = row.candidates.length ? 0 : null;
      row.include = !!row.candidates.length;
      row.status = row.candidates.length ? 'found' : 'notfound';
      plog('поиск «' + q + '»: ' + row.candidates.length + ' кандидатов');
      renderRows();
    }).catch(function (e) {
      row.status = 'error';
      row.message = e.message;
      plog('поиск «' + q + '»: ' + e.message, 'err');
      renderRows();
    });
  }

  function searchAll() {
    if (!state.rows.length) { plog('сначала «Анализировать расчёт»', 'err'); return; }
    var queue = [];
    state.rows.forEach(function (row, i) { if (row.parsed.type && row.include) queue.push(i); });
    if (!queue.length) { plog('нет распознанного пассива для поиска', 'err'); return; }
    plog('массовый поиск: ' + queue.length + ' строк…');
    (function next() {
      if (!queue.length) { plog('массовый поиск завершён', 'ok'); return; }
      searchRow(queue.shift(), false);
      setTimeout(next, 400);
    })();
  }

  // ---------- Применение ----------
  function selectedOps() {
    var ops = [];
    state.rows.forEach(function (row) {
      if (row.include && row.selected != null && row.candidates && row.candidates[row.selected]) {
        var c = row.candidates[row.selected];
        var qty = parseFloat(row.qty);
        ops.push({ row: row, cand: c, qty: isFinite(qty) && qty > 0 ? qty : 1 });
      }
    });
    return ops;
  }

  function startApply() {
    var ops = selectedOps();
    if (!ops.length) { plog('нет выбранных кандидатов для применения', 'err'); return; }
    showPreview(ops);
  }

  function showPreview(ops) {
    var cont = elPanes.helper.querySelector('.rows');
    cont.textContent = '';
    var tbl = h('table', { class: 'preview' },
      h('tr', null, h('th', null, 'Строка расчёта'), h('th', null, 'Будет выбран кандидат'), h('th', null, 'Кол-во')));
    ops.forEach(function (op) {
      var lb = candLabel(op.cand);
      tbl.appendChild(h('tr', null,
        h('td', null, op.row.name),
        h('td', null, lb.pn),
        h('td', null, String(op.qty))));
    });
    cont.appendChild(h('div', { class: 'hint' },
      'Будет выполнен один запрос UserFolders/SaveData к расчёту ' + state.fid +
      ': в ' + ops.length + ' строк(ах) будет закреплён выбранный кандидат (статус «предложение выбрано»). Отменить после применения нельзя.'));
    cont.appendChild(tbl);
    cont.appendChild(h('div', { class: 'toolbar', style: 'margin-top:8px' },
      h('button', { class: 'b red', on: { click: function () { doApply(ops); } } }, 'Подтверждаю, применить'),
      h('button', { class: 'b ghost', on: { click: renderRows } }, 'Отмена')));
  }

  function doApply(ops) {
    state.applying = true;
    // Реальный контракт сайта: SaveData {fid: число, data: группы из GetData},
    // в строках меняем selected + status='offerSelected'; если кандидата нет
    // в items — добавляем item в формате сайта. calc/need не трогаем.
    S.api.getData(state.fid, 0).then(function (res) {
      var data = res && res.data;
      if (!Array.isArray(data)) throw new Error('GetData: массив data не найден');
      var changed = 0, failed = 0;
      ops.forEach(function (op) {
        var g = null, row = null;
        for (var i = 0; i < data.length; i++) {
          if (String(data[i] && data[i].gid) === String(op.row.groupId)) { g = data[i]; break; }
        }
        if (!g || !Array.isArray(g.rows)) { failed++; return; }
        g.rows.forEach(function (r) { if (String(r.id) === String(op.row.rowId)) row = r; });
        if (!row) { failed++; return; }
        if (!Array.isArray(row.items)) row.items = [];
        var exists = null;
        row.items.forEach(function (it) { if (it && it.hash === op.cand.hash) exists = it; });
        if (!exists) {
          var maxId = 0;
          row.items.forEach(function (it) { if (it && typeof it.id === 'number' && it.id > maxId) maxId = it.id; });
          row.items.push({
            hash: op.cand.hash,
            package: op.cand.package || null,
            pnid: op.cand.pnid,
            pn: op.cand.pn || op.cand.pnid,
            brand: op.cand.brand || '',
            scode: op.cand.scode != null ? op.cand.scode : '',
            sim: op.cand.sim != null ? op.cand.sim : '',
            open: 0,
            sType: op.cand.sType || 'PN',
            sOrigin: op.cand.sOrigin || op.row.name,
            stdn: op.cand.stdn || null,
            id: maxId + 1
          });
        }
        row.selected = op.cand.hash;
        row.status = 'offerSelected';
        row.empty_items_reason = '';
        changed++;
      });
      if (!changed) throw new Error('строки не найдены в свежих данных расчёта');
      plog('SaveData: меняю ' + changed + ' строк(и)…');
      return S.api.saveData(state.fid, data).then(function (resp) {
        state.applying = false;
        plog('SaveData OK: saveVersion=' + (resp && resp.saveVersion != null ? resp.saveVersion : '?') +
          ', изменено ' + changed + (failed ? ', не найдено ' + failed : ''), 'ok');
        setSummary('Готово: применено ' + changed + ' из ' + (changed + failed) + '. Обновляю расчёт…');
        setTimeout(analyze, 800);
      });
    }).catch(function (e) {
      state.applying = false;
      plog('Применение не удалось: ' + e.message + (e.appError ? ' [' + e.appError + ']' : ''), 'err');
      setSummary('Применение не удалось: ' + e.message);
    });
  }

  function testParser() {
    var res = S.parsePassiveSelfTest();
    if (!res.fails.length) {
      plog('Тест парсера: ' + res.total + '/' + res.total + ' OK', 'ok');
      setSummary('Тест парсера пройден: ' + res.total + ' кейсов.');
    } else {
      plog('Тест парсера: провалено ' + res.fails.length + ' из ' + res.total, 'err');
      res.fails.forEach(function (f) { plog('  #' + f.n + ' «' + f.input + '»: ожидалось [' + f.expect + '], получено [' + f.got + ']', 'err'); });
      setSummary('Тест парсера: провалено ' + res.fails.length + ' из ' + res.total + ' — детали в логе ниже.');
    }
    setTab('helper');
  }

  function copyDataJson() {
    if (!state.rawData) { plog('сначала «Анализировать расчёт»', 'err'); return; }
    copyText(JSON.stringify(state.rawData, null, 1), 'JSON GetData расчёта ' + state.fid);
  }

  // ---------- Навигация ----------
  function onRoute() {
    var fid = S.currentFid();
    if (fid !== state.fid) {
      state.fid = fid;
      state.rows = [];
      state.rawData = null;
      if (elTabs.fid) elTabs.fid.textContent = fid ? 'расчёт ' + fid : '';
      setSummary(fid ? 'Расчёт ' + fid + ' открыт. Нажмите «Анализировать расчёт».' : 'Откройте страницу расчёта /user/carts/{id}');
      if (elPanes.helper) elPanes.helper.querySelector('.rows').textContent = '';
    }
  }

  S.panel = {
    init: function () {
      if (host) return;
      buildUI();
      S.onNav(function () { onRoute(); });
      S.onLogEntry = function () { if (state.tab === 'rpclog') renderRpcLog(); };
      onRoute();
      plog('панель готова');
    }
  };
})();
