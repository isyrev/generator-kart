/**
 * UECN Extractor UI — 3 modes: ЭПОС / ММ / ЦДС
 */
(function () {
  'use strict';

  var TEMPLATE_URL = 'templates/uecn-v10.json';
  var currentTemplate = null;
  var currentMode = 'epos';
  var lastValues = null;
  var fileLoaded = false;

  var dropZone = document.getElementById('dropZone');
  var fileInput = document.getElementById('fileInput');
  var statusEl = document.getElementById('status');
  var output = document.getElementById('output');
  var overlay = document.getElementById('dragOverlay');
  var modeSelect = document.getElementById('modeSelect');
  var toolbar = document.getElementById('toolbar');
  var fileNameEl = document.getElementById('fileName');
  var dragCounter = 0;

  function setStatus(msg, cls) {
    statusEl.textContent = msg || '';
    statusEl.className = cls || '';
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function isEmpty(v) {
    return UecnParser.isEmpty(v);
  }

  window.copyVal = function (text, btn) {
    navigator.clipboard.writeText(text).then(function () {
      var prev = btn.textContent;
      btn.textContent = '✓';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = prev;
        btn.classList.remove('copied');
      }, 1100);
    });
  };

  window.copyBulk = function (btn) {
    var mode = getMode();
    if (!mode || !lastValues) return;
    var ids = mode.copyIds || [];
    var lines = [];
    for (var i = 0; i < ids.length; i++) {
      var v = UecnParser.clean(lastValues[ids[i]]);
      if (isEmpty(v)) continue;
      lines.push(v);
    }
    navigator.clipboard.writeText(lines.join('\n')).then(function () {
      btn.textContent = '✓ ' + lines.length + ' шт.';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = mode.copyMode === 'mufty' ? 'Копировать муфты' : 'Копировать';
        btn.classList.remove('copied');
      }, 1500);
    });
  };

  function cell(v, asNode) {
    v = UecnParser.clean(v);
    if (isEmpty(v)) return '<span class="empty">—</span>';
    if (asNode) return '<span class="num">' + esc(v) + '</span>';
    return '<span class="txt">' + esc(v) + '</span>';
  }

  function getMode() {
    if (!currentTemplate || !currentTemplate.modes) return null;
    return currentTemplate.modes[currentMode] || currentTemplate.modes[currentTemplate.defaultMode];
  }

  function render(values) {
    lastValues = values;
    var mode = getMode();
    if (!mode) return;

    var headers = mode.headers || [];
    var rows = mode.rows || [];
    var colCount = headers.length;
    var html = '';

    // header
    html += '<div class="tbl-head" style="grid-template-columns: repeat(' + colCount + ', minmax(0,1fr))">';
    for (var h = 0; h < headers.length; h++) {
      html += '<div class="th">' + esc(headers[h]) + '</div>';
    }
    html += '</div>';

    html += '<div class="tbl-body">';
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var ids = row.ids || [];

      // PPO: hide if empty
      if (row.ppo) {
        var ppoVal = UecnParser.clean(values[ids[0]]);
        if (isEmpty(ppoVal)) continue;
      }

      // Комплект №
      if (row.label) {
        html += '<div class="tbl-row label-row" style="grid-template-columns: auto 1fr">';
        html += '<span class="row-label">' + esc(row.label) + '</span>';
        html += cell(values[ids[0]], true);
        html += '</div>';
        continue;
      }

      // mid row (47/48 between 30 and 9)
      if (row.mid) {
        html += '<div class="tbl-row mid-row" style="grid-template-columns: repeat(' + colCount + ', minmax(0,1fr))">';
        for (var j = 0; j < colCount; j++) {
          if (j < ids.length && ids[j] != null) {
            html += '<div class="td">' + cell(values[ids[j]], false) + '</div>';
          } else {
            html += '<div class="td"></div>';
          }
        }
        html += '</div>';
        continue;
      }

      // span: only first col (30 or 9 alone)
      if (row.span) {
        html += '<div class="tbl-row" style="grid-template-columns: repeat(' + colCount + ', minmax(0,1fr))">';
        html += '<div class="td">' + cell(values[ids[0]], true) + '</div>';
        for (var j = 1; j < colCount; j++) html += '<div class="td"></div>';
        html += '</div>';
        continue;
      }

      // normal row — pad to colCount
      html += '<div class="tbl-row" style="grid-template-columns: repeat(' + colCount + ', minmax(0,1fr))">';
      for (var j = 0; j < colCount; j++) {
        if (j < ids.length && ids[j] != null) {
          var isNum = (j === 0);
          html += '<div class="td">' + cell(values[ids[j]], isNum) + '</div>';
        } else {
          html += '<div class="td"></div>';
        }
      }
      html += '</div>';
    }
    html += '</div>';

    // bulk copy button
    if (mode.copyMode === 'nodes' || mode.copyMode === 'mufty') {
      var btnLabel = mode.copyMode === 'mufty' ? 'Копировать муфты' : 'Копировать';
      html += '<div class="copy-all-wrap">' +
        '<button class="btn btn-all" onclick="copyBulk(this)">' + btnLabel + '</button>' +
        '<span class="copy-all-hint">' +
        (mode.copyMode === 'mufty'
          ? 'номера муфт [79]–[89], без пустых'
          : 'номера узлов, без пустых, по порядку') +
        '</span></div>';
    }

    output.innerHTML = html;
  }

  function showLoaded(name) {
    fileLoaded = true;
    dropZone.classList.add('hidden');
    toolbar.classList.remove('hidden');
    fileNameEl.textContent = name;
  }

  function resetFile() {
    fileLoaded = false;
    lastValues = null;
    dropZone.classList.remove('hidden');
    toolbar.classList.add('hidden');
    output.innerHTML = '';
    fileNameEl.textContent = '';
    setStatus('');
    fileInput.value = '';
  }

  window.resetFile = resetFile;

  async function loadTemplate() {
    try {
      var res = await fetch(TEMPLATE_URL + '?t=' + Date.now());
      if (!res.ok) throw new Error('Не удалось загрузить шаблон');
      currentTemplate = await res.json();
      currentMode = currentTemplate.defaultMode || 'epos';
      // fill mode select
      modeSelect.innerHTML = '';
      var modes = currentTemplate.modes || {};
      Object.keys(modes).forEach(function (k) {
        var opt = document.createElement('option');
        opt.value = k;
        opt.textContent = modes[k].name || k;
        if (k === currentMode) opt.selected = true;
        modeSelect.appendChild(opt);
      });
    } catch (err) {
      setStatus('Ошибка шаблона: ' + err.message, 'err');
    }
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      setStatus('Нужен файл .docx', 'err');
      return;
    }
    if (!currentTemplate) {
      setStatus('Шаблон ещё не загружен', 'err');
      return;
    }
    setStatus('Читаю ' + file.name + '…', '');
    try {
      var buf = await file.arrayBuffer();
      var parsed = await UecnParser.parseDocx(buf);
      var values = UecnParser.extract(parsed, currentTemplate);
      showLoaded(file.name);
      render(values);
      setStatus('Готово', 'ok');
    } catch (err) {
      setStatus('Ошибка: ' + err.message, 'err');
      console.error(err);
    }
  }

  modeSelect.addEventListener('change', function () {
    currentMode = modeSelect.value;
    if (lastValues) render(lastValues);
  });

  dropZone.addEventListener('click', function () { fileInput.click(); });
  fileInput.addEventListener('change', function (e) { handleFile(e.target.files[0]); });

  window.addEventListener('dragenter', function (e) {
    e.preventDefault();
    dragCounter++;
    overlay.classList.add('show');
  });
  window.addEventListener('dragleave', function (e) {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) { dragCounter = 0; overlay.classList.remove('show'); }
  });
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    e.preventDefault();
    dragCounter = 0;
    overlay.classList.remove('show');
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  loadTemplate();
})();
