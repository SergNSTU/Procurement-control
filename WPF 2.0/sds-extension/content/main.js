/*
 * SDS Помощник — main.js (content script, isolated world, document_start).
 * 1) сразу инжектит inject.js в MAIN world (логгер RPC-трафика),
 * 2) собирает лог и события навигации через postMessage,
 * 3) после DOM ready поднимает панель.
 */
(function () {
  'use strict';
  var S = (window.SDSX = window.SDSX || {});

  // ---- 1. Инжекция MAIN-world логгера как можно раньше ----
  try {
    var sc = document.createElement('script');
    sc.src = (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getURL)
      ? chrome.runtime.getURL('content/inject.js') : 'inject.js';
    sc.async = false;
    (document.head || document.documentElement).appendChild(sc);
    sc.onload = function () { sc.remove(); };
  } catch (e) { console.warn('[SDS+] inject failed', e); }

  // ---- 2. Приём событий из MAIN world ----
  S.log = [];              // RPC-лог: {t, m, url, req, st, res}
  S.LOG_MAX = 400;
  S.navListeners = [];

  window.addEventListener('message', function (e) {
    var d = e.data;
    if (!d || d.source !== 'sdsx-main') return;
    if (d.type === 'log' && d.entry) {
      S.log.push(d.entry);
      if (S.log.length > S.LOG_MAX) S.log.splice(0, S.log.length - S.LOG_MAX);
      if (S.onLogEntry) S.onLogEntry(d.entry);
    } else if (d.type === 'nav' && d.url) {
      S.navListeners.forEach(function (cb) { try { cb(d.url); } catch (err) {} });
    }
  });

  // ---- 3. Роутинг по URL (страховка к history-хуку: интервал) ----
  var lastPath = location.pathname;
  function checkRoute() {
    if (location.pathname !== lastPath) {
      lastPath = location.pathname;
      S.navListeners.forEach(function (cb) { try { cb(location.href); } catch (err) {} });
    }
  }
  setInterval(checkRoute, 800);

  S.onNav = function (cb) { S.navListeners.push(cb); cb(location.href); };

  // fid текущего расчёта или null
  S.currentFid = function () {
    var m = location.pathname.match(/^\/user\/carts\/(\d+)/);
    return m ? m[1] : null;
  };

  // ---- 4. Панель после DOM ----
  function boot() {
    try {
      if (S.panel && S.panel.init) S.panel.init();
    } catch (e) { console.error('[SDS+] panel init failed', e); }
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
