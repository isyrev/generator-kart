/**
 * UECN Extractor UI — GitHub Pages ready
 */
(function () {
  'use strict';

  var TEMPLATES_INDEX = 'templates/index.json';
  var currentTemplate = null;
  var templatesList = [];

  var dropZone = document.getElementById('dropZone');
  var fileInput = document.getElementById('fileInput');
  var statusEl = document.getElementById('status');
  var output = document.getElementById('output');
  var overlay = document.getElementById('dragOverlay');
  var templateSelect = document.getElementById('templateSelect');
  var dragCounter = 0;

  function setStatus(msg, cls) {
    statusEl.textContent = msg;
    statusEl.className = cls || '';
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  window.copyVal = function (text, btn) {
    navigator.clipboard.writeText(text).then(function () {
      btn.textContent = '✓';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = 'Копировать';
        btn.classList.remove('copied');
      }, 1100);
    });
  };

  function cellCopy(v) {
    v = UecnParser.clean(v);
    if (UecnParser.isEmpty(v)) return '<span class="val empty">—</span>';
    var id = 'c' + Math.random().toString(36).slice(2, 9);
    return '<span class="cell"><span class="val" id="' + id + '">' + esc(v) + '</span>' +
      '<button class="btn" onclick="copyVal(document.getElementById(\'' + id + '\').textContent,this)">Копировать</button></span>';
  }

  function cellStatic(v) {
    v = UecnParser.clean(v);
    if (UecnParser.isEmpty(v)) return '<span class="static empty">—</span>';
    return '<span class="static">' + esc(v) + '</span>';
  }

  function render(values, template) {
    var layout = template.layout || [];
    var fields = template.fields || {};
    var html = '';
    for (var i = 0; i < layout.length; i++) {
      var row = layout[i];
      if (row.gap) {
        html += '<div class="block-gap"></div>';
        continue;
      }
      var cols = row.cols || 1;
      var ids = row.ids || [];
      var cls = cols === 3 ? 'row3' : cols === 2 ? 'row2' : 'row1';
      html += '<div class="' + cls + '">';
      for (var j = 0; j < ids.length; j++) {
        var id = String(ids[j]);
        var f = fields[id] || {};
        var val = values[id] != null ? values[id] : '-';
        if (f.copyable) html += cellCopy(val);
        else html += cellStatic(val);
      }
      html += '</div>';
    }
    output.innerHTML = '<div class="block">' + html + '</div>';
  }

  async function loadTemplatesIndex() {
    try {
      var res = await fetch(TEMPLATES_INDEX + '?t=' + Date.now());
      if (!res.ok) throw new Error('index ' + res.status);
      templatesList = await res.json();
    } catch (e) {
      // fallback: single known template
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
      setStatus('Шаблон: ' + (currentTemplate.name || currentTemplate.id), 'ok');
      output.innerHTML = '';
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

  /* events */
  templateSelect.addEventListener('change', function () {
    loadTemplate(templateSelect.value);
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
