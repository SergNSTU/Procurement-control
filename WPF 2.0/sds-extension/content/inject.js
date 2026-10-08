/*
 * SDS Помощник — inject.js (выполняется в MAIN world, до кода сайта).
 * Тонкие обёртки XHR/fetch для записи RPC-трафика sds.compel.ru в лог
 * (диагностика контракта «Распознать») + мост для sessionStorage.ccid
 * и уведомлений о SPA-навигации. Ничего не блокирует и не меняет запросы:
 * любая ошибка обёртки глушится, оригинальное поведение сохраняется.
 */
(function () {
  'use strict';
  if (window.__sdsxInjected) return;
  window.__sdsxInjected = true;

  var MAX_BODY = 200000;   // символов на тело запроса/ответа в логе
  var MAX_LOG = 400;       // размер кольцевого буфера
  var logSeq = 0;

  function clip(s) {
    if (s == null) return null;
    s = String(s);
    return s.length > MAX_BODY ? s.slice(0, MAX_BODY) + '…[' + s.length + ' chars]' : s;
  }

  function isApi(url) {
    return typeof url === 'string' && url.indexOf('/api/') !== -1;
  }

  function push(entry) {
    entry.seq = ++logSeq;
    entry.t = new Date().toISOString().slice(11, 23);
    try { window.postMessage({ source: 'sdsx-main', type: 'log', entry: entry }, '*'); } catch (e) {}
  }

  // ---- XMLHttpRequest ----
  var XHR = window.XMLHttpRequest;
  var xhrOpen = XHR.prototype.open;
  var xhrSend = XHR.prototype.send;
  XHR.prototype.open = function (method, url) {
    try {
      this.__sdsx = isApi(url) ? { method: method, url: url, req: null } : null;
    } catch (e) { this.__sdsx = null; }
    return xhrOpen.apply(this, arguments);
  };
  XHR.prototype.send = function (body) {
    var self = this;
    if (self.__sdsx) {
      self.__sdsx.req = clip(typeof body === 'string' ? body : null);
      self.addEventListener('load', function () {
        try {
          var rt = self.responseType;
          var text = (rt === '' || rt === 'text') ? self.responseText : '(responseType=' + rt + ')';
          push({ kind: 'xhr', m: self.__sdsx.method, url: String(self.__sdsx.url), req: self.__sdsx.req, st: self.status, res: clip(text) });
        } catch (e) {}
      });
      self.addEventListener('error', function () {
        try { push({ kind: 'xhr', m: self.__sdsx.method, url: String(self.__sdsx.url), req: self.__sdsx.req, st: 0, res: '(network error)' }); } catch (e) {}
      });
    }
    return xhrSend.apply(this, arguments);
  };

  // ---- fetch ----
  var origFetch = window.fetch;
  if (typeof origFetch === 'function') {
    window.fetch = function (input) {
      var url, body, method = 'GET';
      try {
        if (typeof input === 'string' || input instanceof URL) {
          url = String(input);
          if (arguments.length > 1 && arguments[1]) {
            body = arguments[1].body;
            method = arguments[1].method || 'GET';
          }
        } else if (input && typeof input === 'object' && input.url) {
          url = input.url; body = input.body; method = input.method || 'GET';
        }
      } catch (e) {}
      var p = origFetch.apply(this, arguments);
      if (isApi(url)) {
        var reqStr = clip(typeof body === 'string' ? body : null);
        p.then(function (res) {
          try {
            var clone = res.clone();
            clone.text().then(function (txt) {
              push({ kind: 'fetch', m: method, url: url, req: reqStr, st: res.status, res: clip(txt) });
            }).catch(function () {});
          } catch (e) {}
        }).catch(function () {
          try { push({ kind: 'fetch', m: method, url: url, req: reqStr, st: 0, res: '(network error)' }); } catch (e) {}
        });
      }
      return p;
    };
  }

  // ---- мост: ccid из sessionStorage ----
  window.addEventListener('message', function (e) {
    var d = e.data;
    if (d && d.source === 'sdsx-iso') {
      if (d.type === 'need-ccid') {
        var ccid = null;
        try { ccid = window.sessionStorage.getItem('ccid'); } catch (err) {}
        try { window.postMessage({ source: 'sdsx-main', type: 'ccid', ccid: ccid }, '*'); } catch (err) {}
      } else if (d.type === 'get-log') {
        try { window.postMessage({ source: 'sdsx-main', type: 'full-log' }, '*'); } catch (err) {}
      }
    }
  });

  // ---- SPA-навигация ----
  function notifyNav() {
    try { window.postMessage({ source: 'sdsx-main', type: 'nav', url: location.href }, '*'); } catch (e) {}
  }
  ['pushState', 'replaceState'].forEach(function (fn) {
    var orig = history[fn];
    if (typeof orig !== 'function') return;
    history[fn] = function () {
      var r = orig.apply(this, arguments);
      try { notifyNav(); } catch (e) {}
      return r;
    };
  });
  window.addEventListener('popstate', notifyNav);
})();
