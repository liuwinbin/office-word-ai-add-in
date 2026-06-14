/* global Office, Word, console */

(function () {
  'use strict';

  // ── DOM references ──────────────────────────────────────────
  var el = {
    provider:      document.getElementById('modelProvider'),
    ollamaForm:    document.getElementById('ollamaForm'),
    deepseekForm:  document.getElementById('deepseekForm'),
    ollamaUrl:     document.getElementById('ollamaBaseUrl'),
    ollamaModel:   document.getElementById('ollamaModel'),
    deepseekUrl:   document.getElementById('deepseekBaseUrl'),
    deepseekKey:   document.getElementById('deepseekApiKey'),
    testBtn:       document.getElementById('testConnectionBtn'),
    testSpinner:   document.getElementById('testSpinner'),
    testStatus:    document.getElementById('testStatus'),
    selectedText:  document.getElementById('selectedText'),
    extractBtn:    document.getElementById('extractBtn'),
    instruction:   document.getElementById('instruction'),
    executeBtn:    document.getElementById('executeBtn'),
    executeSpinner:document.getElementById('executeSpinner'),
    executeStatus: document.getElementById('executeStatus')
  };

  // ── Configuration persistence ───────────────────────────────
  var STORAGE_KEY = 'officeai_config';

  function loadConfig() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    } catch (e) { /* quota exceeded – silently ignore */ }
  }

  function applyConfig(cfg) {
    el.provider.value = cfg.provider || 'ollama';
    el.ollamaUrl.value = cfg.ollamaUrl || 'http://localhost:11434';
    el.ollamaModel.value = cfg.ollamaModel || 'deepseek-r1:latest';
    el.deepseekUrl.value = cfg.deepseekUrl || 'https://api.deepseek.com';
    el.deepseekKey.value = cfg.deepseekKey || '';
    switchForm(cfg.provider || 'ollama');
  }

  function persistCurrentConfig() {
    saveConfig({
      provider:    el.provider.value,
      ollamaUrl:   el.ollamaUrl.value,
      ollamaModel: el.ollamaModel.value,
      deepseekUrl: el.deepseekUrl.value,
      deepseekKey: el.deepseekKey.value
    });
  }

  // ── Dynamic form switching ──────────────────────────────────
  function switchForm(provider) {
    if (provider === 'deepseek') {
      el.ollamaForm.classList.remove('active');
      el.deepseekForm.classList.add('active');
    } else {
      el.deepseekForm.classList.remove('active');
      el.ollamaForm.classList.add('active');
    }
  }

  el.provider.addEventListener('change', function () {
    switchForm(el.provider.value);
    persistCurrentConfig();
  });

  // Persist on input changes
  [el.ollamaUrl, el.ollamaModel, el.deepseekUrl, el.deepseekKey].forEach(function (input) {
    input.addEventListener('change', persistCurrentConfig);
    input.addEventListener('blur', persistCurrentConfig);
  });

  // ── Status helpers ──────────────────────────────────────────
  function showStatus(elStatus, type, message) {
    elStatus.className = 'status ' + type;
    elStatus.textContent = message;
  }

  function clearStatus(elStatus) {
    elStatus.className = 'status';
    elStatus.textContent = '';
  }

  function setSpinner(spinner, active) {
    if (active) {
      spinner.classList.add('active');
    } else {
      spinner.classList.remove('active');
    }
  }

  function setButtonsDisabled(disabled) {
    el.testBtn.disabled = disabled;
    el.extractBtn.disabled = disabled;
    el.executeBtn.disabled = disabled;
  }

  // ── Build API request ───────────────────────────────────────
  var SYSTEM_PROMPT =
    '你是一个Word排版与文本处理专家。用户会给你一段文本和一个指令。' +
    '请直接输出处理后的文本。如果涉及格式调整，请用HTML标签包裹文本以表示格式' +
    '（例如：<b>加粗</b>, <font size=''5''>三号字</font>），' +
    '不要解释，只输出结果。';

  function buildOllamaRequest(model, text, instruction) {
    return {
      model: model,
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: '文本：\n' + text + '\n\n指令：' + instruction }
      ],
      stream: false
    };
  }

  function buildDeepseekRequest(text, instruction) {
    return {
      model: 'deepseek-chat',
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: '文本：\n' + text + '\n\n指令：' + instruction }
      ],
      stream: false,
      temperature: 0.3
    };
  }

  // ── API call helpers ────────────────────────────────────────
  function parseOllamaResponse(data) {
    if (data && data.message && data.message.content) {
      return data.message.content;
    }
    throw new Error('Ollama返回数据格式异常');
  }

  function parseDeepseekResponse(data) {
    if (data && data.choices && data.choices.length > 0 &&
        data.choices[0].message && data.choices[0].message.content) {
      return data.choices[0].message.content;
    }
    throw new Error('DeepSeek返回数据格式异常');
  }

  // ── Test connection ─────────────────────────────────────────
  el.testBtn.addEventListener('click', function () {
    clearStatus(el.testStatus);
    setSpinner(el.testSpinner, true);
    el.testBtn.disabled = true;

    var provider = el.provider.value;
    var promise;

    if (provider === 'ollama') {
      var baseUrl = el.ollamaUrl.value.replace(/\/+$/, '');
      promise = fetch(baseUrl + '/api/tags', {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' }
      });
    } else {
      var deepseekBase = el.deepseekUrl.value.replace(/\/+$/, '');
      var apiKey = el.deepseekKey.value.trim();
      if (!apiKey) {
        showStatus(el.testStatus, 'error', '请先填写 API Key');
        setSpinner(el.testSpinner, false);
        el.testBtn.disabled = false;
        return;
      }
      promise = fetch(deepseekBase + '/v1/models', {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + apiKey
        }
      });
    }

    promise.then(function (response) {
      if (!response.ok) {
        return response.text().then(function (body) {
          throw new Error('HTTP ' + response.status + ': ' + body.slice(0, 200));
        });
      }
      return response.json();
    }).then(function () {
      showStatus(el.testStatus, 'success', '连接成功！服务可用。');
      persistCurrentConfig();
    }).catch(function (err) {
      console.error('Connection test failed:', err);
      showStatus(el.testStatus, 'error', '连接失败: ' + err.message);
    }).finally(function () {
      setSpinner(el.testSpinner, false);
      el.testBtn.disabled = false;
    });
  });

  // ── Extract selection ───────────────────────────────────────
  el.extractBtn.addEventListener('click', function () {
    clearStatus(el.executeStatus);
    el.extractBtn.disabled = true;

    Word.run(function (context) {
      var selection = context.document.getSelection();
      context.load(selection, 'text');
      return context.sync().then(function () {
        if (selection.text.trim().length === 0) {
          showStatus(el.executeStatus, 'info', '当前未选中任何文本，请先在Word中选中内容。');
        } else {
          el.selectedText.value = selection.text;
          showStatus(el.executeStatus, 'success', '已提取选区文本（' + selection.text.length + ' 字符）');
        }
      });
    }).catch(function (err) {
      console.error('Extract selection failed:', err);
      if (err instanceof OfficeExtension.Error) {
        showStatus(el.executeStatus, 'error', '提取失败: ' + err.debugInfo.message);
      } else {
        showStatus(el.executeStatus, 'error', '提取失败: ' + err.message);
      }
    }).finally(function () {
      el.extractBtn.disabled = false;
    });
  });

  // ── Execute AI instruction ──────────────────────────────────
  el.executeBtn.addEventListener('click', function () {
    clearStatus(el.executeStatus);

    var text = el.selectedText.value.trim();
    var instruction = el.instruction.value.trim();

    if (!text) {
      showStatus(el.executeStatus, 'error', '请先提取或输入待处理的文本。');
      return;
    }
    if (!instruction) {
      showStatus(el.executeStatus, 'error', '请输入AI处理指令。');
      return;
    }

    setSpinner(el.executeSpinner, true);
    setButtonsDisabled(true);

    var provider = el.provider.value;
    var fetchPromise;

    if (provider === 'ollama') {
      var baseUrl = el.ollamaUrl.value.replace(/\/+$/, '');
      var model = el.ollamaModel.value.trim() || 'deepseek-r1:latest';
      var body = buildOllamaRequest(model, text, instruction);

      fetchPromise = fetch(baseUrl + '/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            throw new Error('Ollama API错误 HTTP ' + res.status + ': ' + t.slice(0, 200));
          });
        }
        return res.json();
      }).then(function (data) {
        return parseOllamaResponse(data);
      });
    } else {
      var deepseekBase = el.deepseekUrl.value.replace(/\/+$/, '');
      var apiKey = el.deepseekKey.value.trim();
      if (!apiKey) {
        throw new Error('请先填写 DeepSeek API Key');
      }
      var dsBody = buildDeepseekRequest(text, instruction);

      fetchPromise = fetch(deepseekBase + '/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + apiKey
        },
        body: JSON.stringify(dsBody)
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            throw new Error('DeepSeek API错误 HTTP ' + res.status + ': ' + t.slice(0, 200));
          });
        }
        return res.json();
      }).then(function (data) {
        return parseDeepseekResponse(data);
      });
    }

    fetchPromise.then(function (aiContent) {
      showStatus(el.executeStatus, 'info', 'AI处理完成，正在写回Word文档...');
      return insertContentToWord(aiContent);
    }).then(function () {
      showStatus(el.executeStatus, 'success', '处理完成！结果已替换原选区文本。');
    }).catch(function (err) {
      console.error('AI execution failed:', err);
      showStatus(el.executeStatus, 'error', '处理失败: ' + err.message);
    }).finally(function () {
      setSpinner(el.executeSpinner, false);
      setButtonsDisabled(false);
    });
  });

  // ── Insert HTML back into Word ──────────────────────────────
  function insertContentToWord(htmlContent) {
    return Word.run(function (context) {
      var selection = context.document.getSelection();
      // Replace the selection content with the AI-generated HTML
      selection.insertHtml(htmlContent, Word.InsertLocation.replace);
      return context.sync();
    });
  }

  // ── Office initialization ───────────────────────────────────
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      console.log('OfficeAI: Word host detected, initializing...');
      var cfg = loadConfig();
      applyConfig(cfg);
      persistCurrentConfig();

      // Clear any stale statuses
      clearStatus(el.testStatus);
      clearStatus(el.executeStatus);
    } else {
      console.warn('OfficeAI: Running in unsupported host: ' + info.host);
      showStatus(el.executeStatus, 'error', '此加载项仅支持 Microsoft Word。');
    }
  });

})();
