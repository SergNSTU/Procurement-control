/*
 * SDS Помощник — sds-rpc.js (content script, isolated world).
 * Обёртка над RPC sds.compel.ru: POST /api/{ver}/{Controller}/{Method},
 * X-Requested-With обязателен, X-CCID берём из sessionStorage (через мост).
 * Запросы идут с same-origin сессией сайта — отдельный логин не нужен.
 */
(function () {
  'use strict';
  var S = (window.SDSX = window.SDSX || {});

  var ccid = null;
  function askCcid() {
    try { window.postMessage({ source: 'sdsx-iso', type: 'need-ccid' }, '*'); } catch (e) {}
  }
  window.addEventListener('message', function (e) {
    var d = e.data;
    if (d && d.source === 'sdsx-main' && d.type === 'ccid') ccid = d.ccid;
  });
  askCcid();

  function apiVer() {
    try {
      var v = document.body && document.body.getAttribute('data-ver');
      if (v) return v;
    } catch (e) {}
    return '20.1.198';
  }

  var lastCallAt = 0;
  function throttle(ms) {
    var wait = lastCallAt + ms - Date.now();
    if (wait < 0) wait = 0;
    lastCallAt = Date.now() + wait;
    return new Promise(function (resolve) { setTimeout(resolve, wait); });
  }

  // собственные RPC-вызовы расширения пишем в общий лог (видны в «Лог RPC»
  // и попадают в «Скопировать UserFolders» — для отладки контракта)
  function selfLog(method, reqBody, status, resText) {
    try {
      var S = window.SDSX;
      if (!S || !S.log) return;
      S.log.push({
        kind: 'self', m: 'POST',
        url: '/api/' + apiVer() + '/' + method,
        req: reqBody.length > 200000 ? reqBody.slice(0, 200000) + '…' : reqBody,
        st: status,
        res: resText.length > 200000 ? resText.slice(0, 200000) + '…' : resText,
        t: new Date().toISOString().slice(11, 23)
      });
      var max = S.LOG_MAX || 400;
      if (S.log.length > max) S.log.splice(0, S.log.length - max);
    } catch (e) {}
  }

  async function rpc(method, body, opts) {
    opts = opts || {};
    if (!opts.noThrottle) await throttle(opts.minGapMs || 200);
    askCcid();
    var headers = {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'X-Requested-With': 'XMLHttpRequest'
    };
    if (ccid) headers['X-CCID'] = ccid;
    var reqStr = JSON.stringify(body || {});
    var res = await fetch('/api/' + apiVer() + '/' + method, {
      method: 'POST',
      headers: headers,
      credentials: 'include',
      body: reqStr
    });
    var text = await res.text();
    selfLog(method, reqStr, res.status, text);
    if (!res.ok) {
      var err = new Error('RPC ' + method + ' → HTTP ' + res.status);
      err.status = res.status;
      err.appError = res.headers.get('X-APP-ERROR') || null;
      throw err;
    }
    try { return JSON.parse(text); } catch (e) { return text; }
  }

  S.rpc = rpc;
  S.apiVer = apiVer;

  // Готовые вызовы известных методов
  S.api = {
    getData: function (fid, ro) { return rpc('UserFolders/GetData', { fid: String(fid), ro: ro || 0 }); },
    getCarts: function (fid) { return rpc('UserCarts/GetCarts', { fid: String(fid) }); },
    getClasses: function () { return rpc('Search2/GetClasses', {}); },
    search: function (q, extra) {
      var body = Object.assign({ q: String(q), full: true, entities: [] }, extra || {});
      return rpc('Search2/Load2', body);
    },
    addItemsIntoRow: function (fid, groupId, rowId, items) {
      return rpc('UserFolders/AddItemsIntoRow', { fid: String(fid), groupId: groupId, rowId: rowId, items: items }, { minGapMs: 350 });
    },
    // Реальный контракт сайта (снят с ручного «Распознать»): тело = {fid: число,
    // data: массив групп из GetData с внесённым selected/status}; calc не трогаем.
    saveData: function (fid, data) {
      return rpc('UserFolders/SaveData', { fid: Number(fid), data: data }, { minGapMs: 500 });
    },
    castResolverList: function (items) {
      return rpc('UserFolders/CastResolverList', { items: items }, { minGapMs: 350 });
    }
  };

  // ---- Извлечение строк расчёта из GetData ----
  // Точная структура ответа SDS не известна, поэтому — глубокий рекурсивный
  // поиск «строк» по признакам: есть id, есть текстовое имя-описание, нет
  // дочерних массивов с объектами (иначе это контейнер/группа).
  S.extractRows = function (data) {
    var out = { groups: [], raw: data };
    if (!data || typeof data !== 'object') return out;

    var NAME_FIELDS = ['name', 'caption', 'title', 'description', 'desc', 'source', 'text', 'value', 'query', 'label', 'search'];
    var ITEM_FIELDS = ['items', 'parts', 'positions', 'offers', 'suggestions', 'candidates'];
    var ID_FIELDS = ['id', 'rowId', 'rid', 'row_id'];
    var QTY_FIELDS = ['need', 'qty', 'qnt', 'quantity', 'amount', 'count'];
    var seen = [];
    var MAX_DEPTH = 14;

    function idOf(obj) {
      for (var i = 0; i < ID_FIELDS.length; i++) {
        var v = obj[ID_FIELDS[i]];
        if (v !== undefined && v !== null && v !== '') return v;
      }
      return null;
    }
    function nameOf(obj) {
      for (var i = 0; i < NAME_FIELDS.length; i++) {
        var v = obj[NAME_FIELDS[i]];
        if (typeof v === 'string' && v.trim()) return v.trim();
      }
      if (Array.isArray(obj.names)) {
        for (var j = 0; j < obj.names.length; j++) {
          var n = obj.names[j];
          if (typeof n === 'string' && n.trim()) return n.trim();
          if (n && typeof n === 'object') {
            for (var k = 0; k < NAME_FIELDS.length; k++) {
              if (typeof n[NAME_FIELDS[k]] === 'string' && n[NAME_FIELDS[k]].trim()) return n[NAME_FIELDS[k]].trim();
            }
          }
        }
      }
      return null;
    }
    function itemsCountOf(obj) {
      for (var i = 0; i < ITEM_FIELDS.length; i++) {
        var v = obj[ITEM_FIELDS[i]];
        if (Array.isArray(v)) return v.length;
        if (v !== undefined && v !== null && v !== '' && typeof v !== 'object') return 1;
      }
      return 0;
    }
    function itemsListOf(obj) {
      for (var i = 0; i < ITEM_FIELDS.length; i++) {
        var v = obj[ITEM_FIELDS[i]];
        if (Array.isArray(v)) return v;
      }
      return null;
    }
    function qtyOf(obj) {
      for (var i = 0; i < QTY_FIELDS.length; i++) {
        var v = obj[QTY_FIELDS[i]];
        if (typeof v === 'number') return v;
        if (typeof v === 'string' && v.trim() !== '' && !isNaN(parseFloat(v))) return v;
      }
      return null;
    }
    // контейнер = есть массив-поле (НЕ поле позиций items!), в котором лежат
    // объекты с id (строки/группы). Поле items у реальной строки SDS содержит
    // объекты с id — его контейнером считать нельзя.
    function containerChildArrays(obj) {
      var res = [];
      var keys = Object.keys(obj);
      for (var i = 0; i < keys.length; i++) {
        if (ITEM_FIELDS.indexOf(keys[i]) !== -1) continue;
        var v = obj[keys[i]];
        if (Array.isArray(v) && v.length && typeof v[0] === 'object' && v[0] && idOf(v[0]) != null) {
          res.push(v);
        }
      }
      return res;
    }
    function visited(node) {
      for (var i = 0; i < seen.length; i++) if (seen[i] === node) return true;
      seen.push(node);
      return false;
    }

    function walk(node, parentGroupId, depth) {
      if (!node || typeof node !== 'object' || depth > MAX_DEPTH || visited(node)) return;
      if (Array.isArray(node)) {
        for (var i = 0; i < node.length; i++) walk(node[i], parentGroupId, depth + 1);
        return;
      }
      var id = idOf(node);
      var name = nameOf(node);
      var childArrays = containerChildArrays(node);
      var isContainer = childArrays.length > 0;
      // Группа: контейнер строк. Строка: id + имя и никаких дочерних контейнеров.
      if (isContainer) {
        // у группы SDS id лежит в gid (idOf её не видит — там нет поля id)
        var gid = (node.gid !== undefined && node.gid !== null && node.gid !== '') ? node.gid
          : (id != null ? id : parentGroupId);
        childArrays.forEach(function (arr) {
          for (var i = 0; i < arr.length; i++) walk(arr[i], gid, depth + 1);
        });
        // в прочие поля-объекты тоже заглянем (на всякий случай)
        Object.keys(node).forEach(function (k) {
          var v = node[k];
          if (v && typeof v === 'object' && !Array.isArray(v) && containerChildArrays(v).length === 0 && nameOf(v) == null) {
            walk(v, gid, depth + 1);
          }
        });
        return;
      }
      if (id != null && name != null) {
        var ic = itemsCountOf(node);
        out.groups.push({
          groupId: parentGroupId,
          rowId: id,
          name: name,
          hasItems: ic > 0,
          itemsCount: ic,
          qty: qtyOf(node),
          status: node.status != null ? String(node.status) : null,
          reason: node.empty_items_reason != null ? String(node.empty_items_reason) : null,
          selected: typeof node.selected === 'string' ? node.selected : null,
          items: itemsListOf(node),
          raw: node
        });
        return;
      }
      // не строка и не контейнер — идём глубже по полям-объектам и массивам объектов
      Object.keys(node).forEach(function (k) {
        var v = node[k];
        if (v && typeof v === 'object') walk(v, parentGroupId, depth + 1);
      });
    }

    walk(data, null, 0);
    return out;
  };
})();
