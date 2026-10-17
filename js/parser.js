/**
 * UECN Passport DOCX parser — template-driven
 */
(function (global) {
  'use strict';

  function clean(s) {
    if (s == null) return '-';
    s = String(s).replace(/\u2000/g, ' ').replace(/\s+/g, ' ').trim();
    // убрать все звёздочки * вокруг/внутри значения
    s = s.replace(/\*/g, '').trim();
    if (!s || s === '-' || s === '–' || s === '…' ||
        s === '…..' || s === '….' || s === '—') return '-';
    // strip placeholder leftovers like [31]
    if (/^\[\d+\]$/.test(s)) return '-';
    return s;
  }

  function isEmpty(v) {
    v = clean(v);
    return !v || v === '-';
  }

  function getText(node) {
    var t = '';
    function walk(n) {
      if (n.nodeType === 3) { t += n.nodeValue; return; }
      if (n.nodeType !== 1) return;
      var name = n.localName;
      if (name === 'tab') { t += '\t'; return; }
      if (name === 'br' || name === 'cr') { t += '\n'; return; }
      for (var i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
    }
    walk(node);
    return t;
  }

  function uniqueCells(arr) {
    var out = [], prev = null;
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] !== prev) { out.push(arr[i]); prev = arr[i]; }
    }
    return out;
  }

  function stripLead(u) {
    var i = 0;
    while (i < u.length && u[i] === '') i++;
    return u.slice(i);
  }

  async function parseDocx(arrayBuffer) {
    var zip = await JSZip.loadAsync(arrayBuffer);
    var xmlStr = await zip.file('word/document.xml').async('string');
    var doc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    var NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    var paraTexts = [];
    var paras = doc.getElementsByTagNameNS(NS, 'p');
    for (var pi = 0; pi < paras.length; pi++) paraTexts.push(getText(paras[pi]));

    var tables = doc.getElementsByTagNameNS(NS, 'tbl');
    var allTables = [];
    for (var ti = 0; ti < tables.length; ti++) {
      var trs = tables[ti].getElementsByTagNameNS(NS, 'tr');
      var tableData = [], rawRows = [];
      for (var ri = 0; ri < trs.length; ri++) {
        var tcs = trs[ri].getElementsByTagNameNS(NS, 'tc');
        var cells = [];
        for (var ci = 0; ci < tcs.length; ci++) {
          cells.push(getText(tcs[ci]).replace(/\r/g, '').trim());
        }
        rawRows.push(cells);
        tableData.push(uniqueCells(cells));
      }
      allTables.push({ tableData: tableData, rawRows: rawRows });
    }
    return { paraTexts: paraTexts, tables: allTables };
  }

  function pickTable(parsed, detect) {
    var best = null, bestScore = -1;
    for (var t = 0; t < parsed.tables.length; t++) {
      var td = parsed.tables[t].tableData;
      var sample = '';
      for (var r = 0; r < Math.min(6, td.length); r++) sample += td[r].join(' ');
      var score = td.length;
      for (var d = 0; d < detect.length; d++) {
        if (sample.indexOf(detect[d]) >= 0) score += 50;
      }
      if (score > bestScore) { bestScore = score; best = parsed.tables[t]; }
    }
    return best;
  }

  function extract(parsed, template) {
    var v = {};
    var fieldIds = Object.keys(template.fields || {});
    for (var i = 0; i < fieldIds.length; i++) v[fieldIds[i]] = '-';

    var komplekt = '-';
    var kp = new RegExp(template.komplektPattern || 'Комплект\\s*№\\s*([^\\s\\u2000\\[]+)', 'i');
    for (var i = 0; i < parsed.paraTexts.length; i++) {
      var m = parsed.paraTexts[i].match(kp);
      if (m) {
        komplekt = m[1].trim();
        if (/^\[?\d+\]?$/.test(komplekt) && komplekt.length <= 4) komplekt = '-';
        break;
      }
    }

    var tbl = pickTable(parsed, template.tableDetect || ['1Я', 'ЭЦН']);
    if (!tbl) {
      if (template.fields['29']) v['29'] = komplekt;
      return v;
    }
    var T = tbl.tableData;
    var rawRows = tbl.rawRows;

    function rowHas(i, re) {
      if (i < 0 || i >= T.length) return false;
      return re.test(T[i].join(' '));
    }
    function findRow(re, from) {
      from = from || 0;
      for (var i = from; i < T.length; i++) if (rowHas(i, re)) return i;
      return -1;
    }
    function S(i) { return stripLead(T[i] || []); }
    function ph(x) { return clean(String(x || '').replace(/^\[\d+\]$/, '')); }

    /* ── ЭЦН ──
       full unique: label, marka, gr, Q, H, st, L, num, mufta, diam, max1, max2, priz, factory
       empty:       label, marka, num, mufta, diam, priz
    */
    var mlgStop = T.length;
    for (var i = 0; i < T.length; i++) {
      var j = T[i].join(' ');
      if (/Модульная\s*ловильная|ловильная\s*головка|Расчет\s*мощности|Тип,\/мощ/i.test(j)) {
        mlgStop = i; break;
      }
    }
    var ecnRows = [];
    for (var i = 0; i < mlgStop; i++) {
      var s = S(i);
      var label = s[0] || '';
      if (/^\d*Ян\.?с?$/i.test(label) || /^\d+Я$/i.test(label)) {
        if (s.indexOf('Марка (Серия)') >= 0 || s.indexOf('№ секции') >= 0) continue;
        ecnRows.push(i);
      }
    }
    var ecnData = [];
    for (var k = 0; k < 7; k++) {
      var ri = ecnRows[k];
      if (ri === undefined) {
        ecnData.push({ num: '-', marka: '-', priz: '-', mufta: '-', diam: '-' });
        continue;
      }
      var s = S(ri);
      if (s.length >= 13) {
        // full: [label,marka,gr,Q,H,st,L,num,mufta,diam,m1,m2,priz,factory]
        ecnData.push({
          marka: ph(s[1]), num: ph(s[7]), mufta: ph(s[8]), diam: ph(s[9]), priz: ph(s[12])
        });
      } else if (s.length >= 10) {
        // older layout without mufta/diam columns: marka,num,priz at 1,7,12-ish
        ecnData.push({
          marka: ph(s[1]), num: ph(s[7]), mufta: '-', diam: '-', priz: ph(s[s.length - 2] || s[s.length - 1])
        });
      } else {
        // sparse empty row: label, marka, num, mufta, diam, priz
        ecnData.push({
          marka: ph(s[1]), num: ph(s[2]), mufta: ph(s[3]), diam: ph(s[4]), priz: ph(s[5] || s[s.length - 1])
        });
      }
    }

    /* ── МЛГ ── */
    var mlg = { num: '-', type: '-', priz: '-' };
    var mlgRow = findRow(/Модульная\s*ловильная|ловильная\s*головка/i);
    if (mlgRow >= 0) {
      var s = S(mlgRow);
      var typ = (s[0] || '').replace(/Модульная\s*ловильная\s*головка\s*:\s*да\s*/i, '').trim();
      mlg.type = ph(typ);
      mlg.num = ph((s[1] || '').replace(/^№\s*/, ''));
      mlg.priz = ph((s[2] || '').replace(/Приз\.?\s*:\s*/i, ''));
    }

    /* ── ЭД ── */
    var ed = { type: '-', slots: [
      { num: '-', priz: '-', mufta: '-', diam: '-' },
      { num: '-', priz: '-', mufta: '-', diam: '-' }
    ]};
    var ed1Row = -1, ed2Row = -1;
    var edHdr = findRow(/Тип,\/мощ/);
    var searchFrom = edHdr >= 0 ? edHdr : 10;
    for (var i = searchFrom; i < T.length; i++) {
      var s = S(i);
      var label = s[0] || '';
      if ((label === '1Ян.с' || label === '1Я') && ed1Row < 0) {
        var joined = s.join(' ');
        if (/ЭДБ|ПЭД|-\d+В|_М\d|ЭД\d|Д\d/i.test(joined) || (s[1] && /ЭД|В-\d|Д\d|\[47\]/.test(s[1]))) {
          ed1Row = i;
        }
      }
      if (label === '2Я' && ed1Row >= 0 && ed2Row < 0 && i > ed1Row) ed2Row = i;
    }
    if (ed1Row < 0) {
      for (var i = 10; i < T.length; i++) {
        var s = S(i);
        if ((s[0] === '1Ян.с' || s[0] === '1Я') && s[1] && (/ЭД|В|Д\d|\[47\]/.test(s[1]))) {
          ed1Row = i; break;
        }
      }
    }
    if (ed2Row < 0 && ed1Row >= 0) {
      for (var i = ed1Row + 1; i < ed1Row + 4 && i < T.length; i++) {
        if (S(i)[0] === '2Я') { ed2Row = i; break; }
      }
    }
    if (ed1Row >= 0) {
      var s = S(ed1Row);
      ed.type = ph(s[1]);
      // [30] is usually last non-empty
      var n = (s.length >= 9) ? (s[s.length - 1] || '-') : '-';
      ed.slots[0].num = ph(n);
    }
    if (ed2Row >= 0) {
      var s = S(ed2Row).filter(function (t) { return t !== ''; });
      // possible: 2Я, L, num, diam, mufta, priz, factory
      // or: 2Я, num, diam, mufta, priz, factory
      var num = '-', diam = '-', mufta = '-', priz = '-';
      for (var i = 0; i < s.length; i++) {
        var t = s[i];
        if (/^[A-Za-zА-Яа-яЁёXxХх]\d{5,}/.test(t) || /^\d{8,}$/.test(t) || /^\[9\]$/.test(t)) {
          if (num === '-') num = t;
        }
      }
      // positional fallback for new template: label, ?, num, diam, mufta, priz
      if (s.length >= 6) {
        // find index of num-like
        for (var i = 1; i < s.length; i++) {
          var t = s[i];
          if (/^\d{5,}$/.test(t) || /^[A-Za-zА-Яа-я]\d{4,}/.test(t) || /^\[9\]$/.test(t)) {
            num = t;
            if (s[i + 1]) diam = s[i + 1];
            if (s[i + 2]) mufta = s[i + 2];
            if (s[i + 3] && /^(Р|Н|К\/Р|К)$/i.test(s[i + 3])) priz = s[i + 3];
            else if (s[i + 3]) {
              // maybe order is num, diam, mufta, priz
              if (/^(Р|Н|К\/Р|К)$/i.test(s[i + 3])) priz = s[i + 3];
            }
            break;
          }
        }
      }
      for (var i = 0; i < s.length; i++) {
        if (/^(Р|Н|К\/Р|К)$/i.test(s[i])) priz = s[i];
      }
      ed.slots[1].num = ph(num);
      ed.slots[1].diam = ph(diam);
      ed.slots[1].mufta = ph(mufta);
      ed.slots[1].priz = ph(priz);
      if (isEmpty(ed.type)) {
        for (var i = 0; i < s.length; i++) {
          if (/ЭДБ|ПЭД|-\d+В|Д\d/i.test(s[i])) { ed.type = ph(s[i]); break; }
        }
      }
    }

    /* ── Датчик ── */
    var sensor = { num: '-', type: '-', priz: '-' };
    var sensorRow = -1;
    for (var i = 0; i < T.length; i++) {
      var s = S(i);
      if (s[0] && (/^\d{2}-\d{2}/.test(s[0]) || /^\[49\]$/.test(s[0]))) { sensorRow = i; break; }
    }
    if (sensorRow < 0) {
      var hdr = findRow(/Погружной/i);
      if (hdr >= 0 && hdr + 1 < T.length) sensorRow = hdr + 1;
    }
    if (sensorRow >= 0) {
      var s = S(sensorRow);
      sensor.type = ph(s[0]);
      sensor.num = ph(s[3]);
      sensor.priz = ph(s[4]);
    }

    /* ── ППО ── */
    var ppo = { num: '-', type: '-' };
    var ppoRow = findRow(/ППО/);
    if (ppoRow >= 0) {
      var s = S(ppoRow);
      var raw = (rawRows && rawRows[ppoRow]) ? rawRows[ppoRow] : s;
      var pool = s.concat(raw);
      // type is often cell after ДА/НЕТ
      for (var i = 0; i < s.length; i++) {
        if (/^(ДА|НЕТ)$/i.test(s[i]) && s[i + 1]) {
          var tp = s[i + 1];
          if (!/^№/.test(tp) && tp.indexOf('\u2116') < 0) {
            ppo.type = ph(tp);
          }
        }
      }
      var ppoNum = '-';
      for (var i = 0; i < pool.length; i++) {
        var cell = pool[i] || '';
        if (cell.indexOf('№') < 0 && cell.indexOf('\u2116') < 0) continue;
        var idx = Math.max(cell.lastIndexOf('№'), cell.lastIndexOf('\u2116'));
        var rest = cell.slice(idx).replace(/^[№\u2116]\s*/, '').trim();
        rest = rest.replace(/^\[11\]$/, '').trim();
        if (rest && rest !== '-' && rest !== '—' && !/^\[\d+\]$/.test(rest)) {
          ppoNum = rest; break;
        }
      }
      if (ppoNum === '-') {
        for (var i = pool.length - 1; i >= 0; i--) {
          var t = (pool[i] || '').trim();
          if (!t || t === '-' || t === '—') continue;
          if (/^(ДА|НЕТ|да|нет|ППО)$/i.test(t)) continue;
          if (/^Ø/.test(t) || /ГЛУХ|ППО\s*к/i.test(t)) continue;
          if (/^\[?\d{1,3}\]?$/.test(t)) continue;
          if (/\d{5,}/.test(t) || /[A-Za-zА-Яа-яЁё]\d{4,}/.test(t)) {
            if (t.length > 25) continue;
            ppoNum = t.replace(/^[№\u2116]\s*/, '').trim();
            break;
          }
        }
      }
      ppo.num = ph(ppoNum);
    }

    /* ── Г/С, ГЗ, МС ──
       unique: label, type, [oil?], diam, L, num, mufta, priz, max1, max2, factory
    */
    function equip(re) {
      var row = findRow(re);
      var o = { type: '-', num: '-', priz: '-', mufta: '-', diam: '-' };
      if (row < 0) return o;
      var s = S(row);
      // Find positions by content
      o.type = ph(s[1]);
      // scan for num (long digits / serial), mufta, diam, priz
      for (var i = 2; i < s.length; i++) {
        var t = s[i];
        if (/^(Р|Н|К\/Р|К)$/i.test(t)) { o.priz = ph(t); continue; }
        if (/^\[\d+\]$/.test(t)) continue;
        if (/^\d+[.,‚]\d+$/.test(t) && parseFloat(t.replace('‚', '.')) < 50) {
          // L length small float — skip if already have diam candidate
          continue;
        }
        if (/^\d+[.,‚]\d+$/.test(t) || /^\d{2,3}([ЭэXxХх]\d{2,3})?$/.test(t) || /^\[(98|99|150)\]$/.test(t)) {
          if (o.diam === '-') o.diam = ph(t);
          continue;
        }
        if (/^\d{5,}$/.test(t) || /^[A-Za-zА-Яа-яЁё]\d{4,}/.test(t) || /^\[(12|13|14)\]$/.test(t)) {
          if (o.num === '-') { o.num = ph(t); continue; }
          if (o.mufta === '-') { o.mufta = ph(t); continue; }
        }
        if (/^\d{6,}$/.test(t) || /^\*\d+\*$/.test(t)) {
          if (o.mufta === '-') o.mufta = ph(t.replace(/\*/g, ''));
        }
      }
      // positional fallback for known structure
      if (s.length >= 8) {
        // label, type, ?, diam, L, num, mufta, priz
        if (o.num === '-' || o.num === undefined) o.num = ph(s[5]);
        if (o.mufta === '-') o.mufta = ph(s[6]);
        if (o.priz === '-') o.priz = ph(s[7]);
        if (o.diam === '-') o.diam = ph(s[3]);
      }
      return o;
    }
    var gs = equip(/Г\/С|Диспергатор|Газосеп/i);
    var gz = equip(/Гидрозащита/i);
    var ms = equip(/Модуль\s*смещ/i);

    /* ── ОК + ШУ ── */
    var okSlots = [
      { type: '-', num: '-', priz: '-', shuType: '-', shuNum: '-', shuPriz: '-' },
      { type: '-', num: '-', priz: '-', shuType: '-', shuNum: '-', shuPriz: '-' }
    ];
    var okRows = [];
    for (var i = 0; i < T.length; i++) {
      var s = S(i);
      if (s[0] === 'ОК' || s[0] === 'Ок') okRows.push(i);
    }
    for (var oi = 0; oi < Math.min(2, okRows.length); oi++) {
      var s = S(okRows[oi]);
      okSlots[oi] = {
        type: ph(s[1]), num: ph(s[2]), priz: ph(s[3]),
        shuType: ph(s[6]), shuNum: ph(s[7]), shuPriz: ph(s[8])
      };
    }

    /* ── КС ── */
    var ks = { type: '-', num: '-', priz: '-' };
    var ksRow = -1;
    for (var i = 0; i < T.length; i++) {
      if (S(i)[0] === 'КС') { ksRow = i; break; }
    }
    if (ksRow >= 0) {
      var s = S(ksRow);
      ks.type = ph(s[1]);
      ks.num = ph(s[2]);
      ks.priz = ph(s[3]);
    }

    /* ── Фильтр ── */
    var filter = { num1: '-', num2: '-', type1: '-', type2: '-' };
    var filtRow = findRow(/Сепаратор|Разобщитель/i);
    if (filtRow >= 0) {
      var s = S(filtRow);
      for (var i = 1; i < s.length; i++) {
        var t = s[i];
        if (!t || t === '-') continue;
        if (/Сепаратор|Разобщитель/.test(t)) continue;
        if ((/[A-Za-zА-Яа-яЁё]/.test(t) || /\[\d+\]\[\d+\]/.test(t)) && !/мм|Кол|Ø|Контейнер/i.test(t)) {
          var tp = t.replace(/\[76\]/g, '').replace(/\[77\]/g, '').trim();
          var tparts = tp.split(/[\n\r]+/).map(function (x) { return x.trim(); }).filter(Boolean);
          if (tparts.length < 2) {
            var mm = tp.match(/(\S+)\s+(\S+)/);
            if (mm) tparts = [mm[1], mm[2]];
          }
          if (tparts[0]) filter.type1 = ph(tparts[0]);
          if (tparts[1]) filter.type2 = ph(tparts[1]);
          break;
        }
      }
      var numCell = null;
      for (var i = 1; i < s.length; i++) {
        var t = s[i];
        if (!t || t === '-') continue;
        if (/мм|Сепаратор|Разобщитель|Контейнер|Кол\.|Ø/i.test(t)) continue;
        if (/[A-Za-zА-Яа-яЁё]{2,}/.test(t) && /\d/.test(t) && t.indexOf('-') >= 0 && !/\[\d+\]/.test(t)) continue;
        if (/^\d+[.,‚]\d+$/.test(t)) continue;
        if (/\[\d+\]\[\d+\]/.test(t) && !/\[19\]|\[20\]/.test(t)) continue;
        numCell = t; break;
      }
      if (numCell && numCell !== '-') {
        var parts = numCell.split(/[\n\r]+/).map(function (x) { return x.trim(); }).filter(Boolean);
        if (parts.length < 2) {
          var m = numCell.match(/\[?(\d+|-)\]?\s*\[?(\d+|-)\]?/);
          if (m) parts = [m[1], m[2]];
          else {
            var sp = numCell.trim().split(/\s+/).filter(Boolean);
            if (sp.length >= 2) parts = sp;
          }
        }
        parts = parts.map(function (p) { return p.replace(/^\[|\]$/g, '').trim(); });
        if (parts[0] && parts[0] !== '19') filter.num1 = ph(parts[0]);
        if (parts[1] && parts[1] !== '20') filter.num2 = ph(parts[1]);
      }
    }

    /* ── Протектолайзеры: 4 ряда × 2 группы ── */
    var protSlots = [];
    for (var pi = 0; pi < 8; pi++) protSlots.push({ num: '-', marka: '-', priz: '-' });
    function isProtLabel(t) {
      return t === 'Г/З' || t === 'МЛГ' || t === 'МС' || t === 'ГЗ' || /^\d+-я$/.test(t || '');
    }
    function isProtNum(t) {
      if (!t || t === '-' || t === '—') return false;
      if (isProtLabel(t)) return false;
      if (/^ПУ/i.test(t) || /^ПК/i.test(t)) return false;
      if (/^(Р|Н|К\/Р|К)$/i.test(t)) return false;
      if (/^\[\d+\]$/.test(t)) return false;
      return /\d{5,}/.test(t) || /[A-Za-zА-Яа-яЁё]\d{4,}/.test(t);
    }
    var protStart = -1;
    for (var i = 0; i < T.length; i++) {
      if (/Протектолайзер/i.test(T[i].join(' '))) { protStart = i; break; }
    }
    if (protStart >= 0) {
      var left = [], mid = [];
      for (var i = protStart + 1; i < Math.min(protStart + 8, T.length); i++) {
        var raw = (rawRows && rawRows[i]) ? rawRows[i].slice() : (T[i] || []).slice();
        while (raw.length && raw[0] === '') raw.shift();
        var groups = [];
        for (var g = 0; g + 3 < raw.length; g += 4) {
          var lab = (raw[g] || '').trim();
          var marka = (raw[g + 1] || '').trim();
          var num = (raw[g + 2] || '').trim();
          var priz = (raw[g + 3] || '').trim();
          if (!isProtLabel(lab)) continue;
          if (!isProtNum(num)) num = '-';
          else num = clean(num);
          if (/^ПУ/i.test(marka) || /^ПК/i.test(marka) || (/[A-Za-zА-Яа-я]/.test(marka) && marka.length > 1)) {
            marka = clean(marka);
          } else {
            marka = '-';
          }
          if (!/^(Р|Н|К\/Р|К)$/i.test(priz)) priz = '-';
          groups.push({ marka: marka, num: num, priz: clean(priz) });
        }
        if (groups[0]) left.push(groups[0]);
        if (groups[1]) mid.push(groups[1]);
      }
      var ordered = left.concat(mid);
      for (var pi = 0; pi < 8 && pi < ordered.length; pi++) {
        protSlots[pi] = ordered[pi];
      }
    }

    /* ── map to fields ── */
    var sources = {
      komplekt: komplekt,
      ecn: ecnData,
      mlg: mlg,
      ed: ed,
      sensor: sensor,
      ppo: ppo,
      gs: gs, gz: gz, ms: ms,
      ok: okSlots,
      ks: ks,
      filter: filter,
      prot: protSlots
    };

    for (var fid in template.fields) {
      var f = template.fields[fid];
      var src = f.source;
      if (src === 'komplekt') { v[fid] = clean(komplekt); continue; }
      if (src === 'ecn') {
        var slot = ecnData[f.slot] || {};
        v[fid] = clean(slot[f.part] || '-');
        continue;
      }
      if (src === 'ed') {
        if (f.part === 'type') { v[fid] = clean(ed.type); continue; }
        var es = ed.slots[f.slot] || {};
        v[fid] = clean(es[f.part] || '-');
        continue;
      }
      if (src === 'ok') {
        var os = okSlots[f.slot] || {};
        v[fid] = clean(os[f.part] || '-');
        continue;
      }
      if (src === 'prot') {
        var pk = protSlots[f.slot] || {};
        v[fid] = clean(pk[f.part] || '-');
        continue;
      }
      if (src === 'mlg' || src === 'sensor' || src === 'ppo' || src === 'gs' ||
          src === 'gz' || src === 'ms' || src === 'ks' || src === 'filter') {
        var obj = sources[src] || {};
        v[fid] = clean(obj[f.part] || '-');
        continue;
      }
    }

    for (var fid in v) {
      if (/^\[\d+\]$/.test(v[fid])) v[fid] = '-';
    }
    return v;
  }

  global.UecnParser = {
    parseDocx: parseDocx,
    extract: extract,
    clean: clean,
    isEmpty: isEmpty
  };
})(typeof window !== 'undefined' ? window : this);
