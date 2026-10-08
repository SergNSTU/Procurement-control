/*
 * SDS Помощник — passive-parser.js (content script, isolated world).
 * Разбор описания пассивного компонента из строки BOM:
 *   «0402, 100 nF, 25V, X7R, 10%»  → керамический конденсатор
 *   «0402, 10k, 5%»                → постоянный резистор
 * Результат: { type: 'cap'|'res'|null, package, value, dielectric, voltage,
 *               tolerance, power, query, confidence, parts:{} }
 * query — нормализованная поисковая строка для Search2/Load2.
 */
(function () {
  'use strict';
  var S = (window.SDSX = window.SDSX || {});

  // Imperial ↔ метрические корпуса SMD
  var PACK_IMPERIAL = ['01005', '0201', '0402', '0603', '0805', '1206', '1210', '1806', '1812', '2010', '2512', '2920'];
  var PACK_METRIC = { '1005': '0402', '1608': '0603', '2012': '0805', '3216': '1206', '3225': '1210', '4516': '1806', '4532': '1812', '5025': '2010', '5750': '2220', '6332': '2512' };

  var DIELECTRICS = ['c0g', 'cog', 'np0', 'npo', 'x5r', 'x6r', 'x6s', 'x7r', 'x7s', 'x7t', 'x8r', 'x7l', 'y5v', 'z5u', 'y5u'];

  function num(ru) {
    // «4,7» → 4.7
    return parseFloat(String(ru).replace(',', '.'));
  }

  function fmtNum(x) {
    var s = String(x);
    return s;
  }

  // Приведение ёмкости к канонической записи: ≥1uF → uF, ≥1nF → nF, иначе pF
  function capCanonical(value, unit) {
    var f = value; // в исходной единице
    var farads;
    switch (unit) {
      case 'pf': farads = f * 1e-12; break;
      case 'nf': farads = f * 1e-9; break;
      case 'uf': farads = f * 1e-6; break;
      case 'mf': farads = f * 1e-3; break;
      case 'f': farads = f; break;
      default: return null;
    }
    if (farads >= 1e-6) return { text: trimZeros(farads / 1e-6) + 'uF', farads: farads };
    if (farads >= 1e-9) return { text: trimZeros(farads / 1e-9) + 'nF', farads: farads };
    return { text: trimZeros(farads * 1e12) + 'pF', farads: farads };
  }

  function trimZeros(x) {
    var s = (Math.round(x * 1e6) / 1e6).toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }

  // Номинал резистора → каноническая запись (10k, 4.7k, 100R, 1M, 0.01R)
  function resCanonical(value, mult) {
    var ohms;
    switch (mult) {
      case 'mo': case 'm': ohms = value * 1e6; break;   // M / МОм
      case 'ko': case 'k': ohms = value * 1e3; break;   // K / кОм
      case 'o': case 'r': ohms = value; break;          // R / Ом / Ohm / Ω
      case 'mo_milli': ohms = value / 1e3; break;       // mΩ (миллиом)
      default: return null;
    }
    if (ohms >= 1e6) return { text: trimZeros(ohms / 1e6) + 'M', ohms: ohms };
    if (ohms >= 1e3) return { text: trimZeros(ohms / 1e3) + 'k', ohms: ohms };
    return { text: trimZeros(ohms) + 'R', ohms: ohms };
  }

  S.parsePassive = function (rawText) {
    var r = { type: null, package: null, value: null, dielectric: null, voltage: null, tolerance: null, power: null, query: null, confidence: 0, parts: {}, src: String(rawText == null ? '' : rawText) };
    var text = r.src.toLowerCase()
      .replace(/µ/g, 'u').replace(/μ/g, 'u')
      .replace(/мкф/g, 'uf').replace(/пф/g, 'pf').replace(/нф/g, 'nf')
      .replace(/мком/g, 'm').replace(/ком/g, 'k').replace(/ом/g, 'o')
      .replace(/ω/g, 'o').replace(/Ω/g, 'o');
    // Словесные подсказки типа
    var hintCap = /конденсатор|керами|capacitor|ceramic|тантал|tantal/.test(text);
    var hintRes = /резистор|resistor|термистор|thermistor/.test(text);

    // Корпус
    var m = text.match(/\b(01005|0201|0204|0402|0603|0805|1206|1210|1806|1812|2010|2512|2920)\b/);
    if (m) { r.package = m[1]; r.parts.package = m[1]; }
    else {
      m = text.match(/\b(1005|1608|2012|3216|3225|4516|4532|5025|5750|6332)\b/);
      if (m) { r.package = PACK_METRIC[m[1]]; r.parts.packageMetric = m[1]; r.parts.package = r.package; }
    }

    // Ёмкость: 100nF, 10 uF, 4,7nF, 1n
    m = text.match(/(\d+(?:[.,]\d+)?)\s*(pf|nf|uf|mf|f)\b/);
    var capMatch = null;
    if (m) {
      capMatch = capCanonical(num(m[1]), m[2]);
      if (capMatch) { r.value = capMatch.text; r.parts.capacitance = capMatch.text; r.type = 'cap'; }
    }

    // Диэлектрик
    m = text.match(/\b(c0g|cog|np0|npo|x5r|x6r|x6s|x7r|x7s|x7t|x8r|x7l|y5v|z5u|y5u)\b/);
    if (m) { r.dielectric = m[1].toUpperCase().replace('COG', 'C0G').replace('NPO', 'NP0'); r.parts.dielectric = r.dielectric; if (!r.type) r.type = 'cap'; }

    // Напряжение: 25V, 6,3V, 50 vdc — признак конденсатора только вместе с корпусом/ёмкостью/диэлектриком
    m = text.match(/(\d+(?:[.,]\d+)?)\s*v(?:dc|ac)?\b/);
    if (m) {
      r.voltage = fmtNum(num(m[1])) + 'V'; r.parts.voltage = r.voltage;
      if (!r.type && (r.package || r.value || r.dielectric)) r.type = 'cap';
    }

    // Допуск: 10%, ±5%, 1 %
    m = text.match(/[±+~]?\s*(\d+(?:[.,]\d+)?)\s*%/);
    if (m) { r.tolerance = fmtNum(num(m[1])) + '%'; r.parts.tolerance = r.tolerance; }

    // Сопротивление (только если ёмкости нет): 10k, 4,7k, 100r, 1m, 10o («Ом»→o),
    // «10 мком»→10m, «10 ком»→10k (замены выше), EIA-нотация «1k5»/«4r7», «r047»
    if (!r.type) {
      m = text.match(/\b(\d+(?:[.,]\d+)?)\s*(mo|ko|m|k|o|r)(?![a-z0-9])/);
      var resMatch = null;
      if (m) {
        resMatch = resCanonical(num(m[1]), m[2]);
      } else {
        // EIA-нотация: 1k5, 4r7, 1m5
        var m2 = text.match(/\b(\d)(k|m|r)(\d)\b/);
        if (m2) {
          var multMap = { k: 'ko', m: 'mo', r: 'o' };
          resMatch = resCanonical(num(m2[1] + '.' + m2[3]), multMap[m2[2]]);
        } else {
          // миллиомы: 10mo — уже поймано выше; «0,01 o» поймано; «r047»:
          var m3 = text.match(/\br(\d{2,3})\b/);
          if (m3) resMatch = resCanonical(num(m3[1].replace(/^0+/, '') || '0') / 1000, 'o');
        }
      }
      if (resMatch) { r.value = resMatch.text; r.parts.resistance = resMatch.text; r.type = 'res'; }
    }

    // Мощность: 0,25W, 1/4W, 0,5 Вт
    m = text.match(/(\d+\/\d+)\s*w\b/);
    if (m) { r.power = m[1] + 'W'; r.parts.power = r.power; }
    else {
      m = text.match(/(\d+(?:[.,]\d+)?)\s*w(?:tt)?\b/);
      if (m) { r.power = fmtNum(num(m[1])) + 'W'; r.parts.power = r.power; }
      m = text.match(/(\d+(?:[.,]\d+)?)\s*вт\b/);
      if (!r.power && m) { r.power = fmtNum(num(m[1])) + 'W'; r.parts.power = r.power; }
    }

    // Сверка типа со словесными подсказками
    if (r.type === 'cap' && hintRes && !hintCap) r.type = 'cap'; // ёмкость сильнее слова
    if (!r.type && hintCap) r.type = 'cap';
    if (!r.type && hintRes) r.type = 'res';
    if (r.type === 'res' && hintCap && !hintRes) r.type = null; // «конденсатор» без ёмкости, но с k-числом? — не рискуем

    // Уверенность
    if (r.type === 'cap') {
      r.confidence = 1;
      if (!r.package) r.confidence -= 0.35;
      if (!r.value) r.confidence -= 0.5;
      if (!r.dielectric) r.confidence -= 0.05;
    } else if (r.type === 'res') {
      r.confidence = 1;
      if (!r.package) r.confidence -= 0.35;
      if (!r.value) r.confidence -= 0.5;
    }

    // Поисковый запрос
    if (r.type) {
      var q = [];
      if (r.package) q.push(r.package);
      if (r.value) q.push(r.value);
      if (r.type === 'cap' && r.voltage) q.push(r.voltage);
      if (r.type === 'cap' && r.dielectric) q.push(r.dielectric);
      if (r.tolerance) q.push(r.tolerance);
      if (r.power && r.type === 'res') q.push(r.power);
      r.query = q.join(' ');
    }
    return r;
  };

  // ---- Самопроверка (вызывается из панели, кнопка «Тест парсера») ----
  S.parsePassiveSelfTest = function () {
    var cases = [
      ['0402, 100 nF, 25V, X7R, 10%', 'cap', '0402', '100nF', '25V', 'X7R', '10%'],
      ['0402, 10k, 5%', 'res', '0402', '10k', null, null, '5%'],
      ['0805 10uF 16V X5R 20%', 'cap', '0805', '10uF', '16V', 'X5R', '20%'],
      ['0603 4,7k 1%', 'res', '0603', '4.7k', null, null, '1%'],
      ['Конденсатор 0603 1uF 25V X7R', 'cap', '0603', '1uF', '25V', 'X7R', null],
      ['Резистор 1206 10 Ом 1%', 'res', '1206', '10R', null, null, '1%'],
      ['0402 100nF', 'cap', '0402', '100nF', null, null, null],
      ['1k 0402 5%', 'res', '0402', '1k', null, null, '5%'],
      ['GRM188R71H104KA93D', null, null, null, null, null, null],
      ['LM317 (JSCJ)', null, null, null, null, null, null],
      ['conn DB9 female', null, null, null, null, null, null],
      ['0603 4r7 5%', 'res', '0603', '4.7R', null, null, '5%'],
      ['0805 2200nF 50V X7R 10%', 'cap', '0805', '2.2uF', '50V', 'X7R', '10%'],
      ['0402 0 Ом джампер', 'res', '0402', '0R', null, null, null]
    ];
    var fails = [];
    cases.forEach(function (c, i) {
      var r = S.parsePassive(c[0]);
      var ok = r.type === c[1] && (r.package || null) === c[2] && (r.value || null) === c[3] &&
        (r.voltage || null) === c[4] && (r.dielectric || null) === c[5] && (r.tolerance || null) === c[6];
      if (!ok) fails.push({ n: i + 1, input: c[0], expect: c.slice(1).join(' | '), got: [r.type, r.package, r.value, r.voltage, r.dielectric, r.tolerance].join(' | ') });
    });
    return { total: cases.length, fails: fails };
  };
})();
