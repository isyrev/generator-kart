/**
 * UECN Extractor UI — GitHub Pages ready
 */
(function () {
  'use strict';

  var TEMPLATES_INDEX = 'templates/index.json';
  var currentTemplate = null;
  var currentLayoutKey = 'full';
  var lastValues = null;
  var templatesList = [];

  var dropZone = document.getElementById('dropZone');
  var fileInput = document.getElementById('fileInput');
  var statusEl = document.getElementById('status');
  var output = document.getElementById('output');
  var overlay = document.getElementById('dragOverlay');
  var templateSelect = document.getElementById('templateSelect');
  var layoutSelect = document.getElementById('layoutSelect');
  var dragCounter = 0;

  function setStatus(msg, cls) {
    statusEl.textContent = msg;
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

  window.copyAll = function (btn) {
    var layout = getLayout();
    if (!layout || !lastValues) return;
    var ids = layout.copyIds || [];
    var lines = [];
    for (var i = 0; i < ids.length; i++) {
      var v = UecnParser.clean(lastValues[ids[i]]);
      // только реальные номера, без «-» и пустых
      if (isEmpty(v)) continue;
      lines.push(v);
    }
    var text = lines.join('\n');
    navigator.clipboard.writeText(text).then(function () {
      var n = lines.length;
      btn.textContent = '✓ ' + n + ' шт.';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = 'Копировать всё (1–30)';
        btn.classList.remove('copied');
      }, 1500);
    });
  };

  function cellCopy(v) {
    v = UecnParser.clean(v);
    if (isEmpty(v)) return '<span class="val empty">—</span>';
    var id = 'c' + Math.random().toString(36).slice(2, 9);
    return '<span class="cell"><span class="val" id="' + id + '">' + esc(v) + '</span>' +
      '<button class="btn" onclick="copyVal(document.getElementById(\'' + id + '\').textContent,this)">Копировать</button></span>';
  }

  function cellStatic(v) {
    v = UecnParser.clean(v);
    if (isEmpty(v)) return '<span class="static empty">—</span>';
    return '<span class="static">' + esc(v) + '</span>';
  }

  function getLayout() {
    if (!currentTemplate) return null;
    var layouts = currentTemplate.layouts || {};
    return layouts[currentLayoutKey] || layouts[currentTemplate.defaultLayout] || layouts.full || null;
  }

  function render(values, template) {
    lastValues = values;
    var layout = getLayout();
    if (!layout) {
      output.innerHTML = '<p class="err">Нет layout в шаблоне</p>';
      return;
    }
    var fields = template.fields || {};
    var copyMode = layout.copyMode || 'perField';
    var rows = layout.rows || [];
    var html = '';

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row.gap) {
        html += '<div class="block-gap"></div>';
        continue;
      }
      // labelIf: строка только если поле непустое (ППО)
      var labelIfActive = false;
      if (row.labelIf) {
        var check = UecnParser.clean(values[row.labelIf]);
        if (isEmpty(check)) continue;
        labelIfActive = true;
      }

      var cols = row.cols || 1;
      var ids = row.ids || [];
      var cls = cols === 3 ? 'row3' : cols === 2 ? 'row2' : 'row1';

      // Компакт + ППО: [номер] | «ППО»  — подпись во 2-й колонке
      if (copyMode === 'all' && labelIfActive && row.label) {
        html += '<div class="row2">';
        var id0 = String(ids[0] || row.labelIf);
        var f0 = fields[id0] || {};
        var val0 = values[id0] != null ? values[id0] : '-';
        html += cellStatic(val0);
        html += '<span class="row-label col2">' + esc(row.label) + '</span>';
        html += '</div>';
        continue;
      }

      html += '<div class="' + cls + '">';

      // Полный режим: «Комплект №» / «ППО» слева от значения
      if (row.label && copyMode === 'perField') {
        html += '<span class="row-label">' + esc(row.label) + '</span>';
      } else if (row.label && !labelIfActive) {
        // компакт: «Комплект №» тоже показываем
        html += '<span class="row-label">' + esc(row.label) + '</span>';
      }

      for (var j = 0; j < ids.length; j++) {
        var id = String(ids[j]);
        var f = fields[id] || {};
        var val = values[id] != null ? values[id] : '-';
        if (copyMode === 'perField' && f.copyable) {
          html += cellCopy(val);
        } else {
          html += cellStatic(val);
        }
      }
      html += '</div>';
    }

    if (copyMode === 'all') {
      html += '<div class="copy-all-wrap">' +
        '<button class="btn btn-all" onclick="copyAll(this)">Копировать всё (1–30)</button>' +
        '<span class="copy-all-hint">в порядке строк, через перевод строки</span></div>';
    }

    output.innerHTML = '<div class="block">' + html + '</div>';
  }

  function fillLayoutSelect() {
    layoutSelect.innerHTML = '';
    if (!currentTemplate || !currentTemplate.layouts) return;
    var keys = Object.keys(currentTemplate.layouts);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      var opt = document.createElement('option');
      opt.value = k;
      opt.textContent = currentTemplate.layouts[k].name || k;
      if (k === (currentTemplate.defaultLayout || 'full')) opt.selected = true;
      layoutSelect.appendChild(opt);
    }
    currentLayoutKey = layoutSelect.value;
  }

  async function loadTemplatesIndex() {
    try {
      var res = await fetch(TEMPLATES_INDEX + '?t=' + Date.now());
      if (!res.ok) throw new Error('index ' + res.status);
      templatesList = await res.json();
    } catch (e) {
      templatesList = [{ id: 'uecn-v10', file: 'templates/uecn-v10.json', name: 'УЭЦН v10' }];
    }
    templateSelect.innerHTML = '';
    for (var i = 0; i < templatesList.length; i++) {
      var t = templatesList[i];
      var opt = document.createElement('option');
      opt.value = t.file || ('templates/' + t.id + '.json');
      opt.textContent = t.name || t.id;
      templateSelect.appendChild(opt);
    }
    if (templatesList.length) await loadTemplate(templateSelect.value);
  }

  async function loadTemplate(path) {
    setStatus('Загрузка шаблона…', '');
    try {
      var res = await fetch(path + '?t=' + Date.now());
      if (!res.ok) throw new Error('Шаблон не найден: ' + path);
      currentTemplate = await res.json();
      fillLayoutSelect();
      setStatus('Шаблон: ' + (currentTemplate.name || currentTemplate.id), 'ok');
      if (lastValues) render(lastValues, currentTemplate);
      else output.innerHTML = '';
    } catch (err) {
      setStatus('Ошибка шаблона: ' + err.message, 'err');
      currentTemplate = null;
    }
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      setStatus('Нужен файл .docx', 'err');
      return;
    }
    if (!currentTemplate) {
      setStatus('Сначала выберите шаблон', 'err');
      return;
    }
    setStatus('Читаю ' + file.name + '…', '');
    try {
      var buf = await file.arrayBuffer();
      var parsed = await UecnParser.parseDocx(buf);
      var values = UecnParser.extract(parsed, currentTemplate);
      render(values, currentTemplate);
      setStatus('Готово: ' + file.name + '  ·  ' + (currentTemplate.name || currentTemplate.id), 'ok');
    } catch (err) {
      setStatus('Ошибка: ' + err.message, 'err');
      console.error(err);
    }
  }

  templateSelect.addEventListener('change', function () {
    loadTemplate(templateSelect.value);
  });
  layoutSelect.addEventListener('change', function () {
    currentLayoutKey = layoutSelect.value;
    if (lastValues && currentTemplate) render(lastValues, currentTemplate);
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

  loadTemplatesIndex();
})();
