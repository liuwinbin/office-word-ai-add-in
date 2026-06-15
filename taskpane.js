/* global Office, Word, console */

(function () {
  'use strict';

  // ── DOM 引用 ──────────────────────────────────────────────
  var el = {
    // 服务商选择
    provider:       document.getElementById('modelProvider'),
    // DeepSeek 表单
    deepseekForm:   document.getElementById('deepseekForm'),
    deepseekUrl:    document.getElementById('deepseekBaseUrl'),
    deepseekKey:    document.getElementById('deepseekApiKey'),
    deepseekModel:  document.getElementById('deepseekModel'),
    toggleKeyBtn:   document.getElementById('toggleKeyBtn'),
    // Ollama 表单
    ollamaForm:     document.getElementById('ollamaForm'),
    ollamaUrl:      document.getElementById('ollamaBaseUrl'),
    ollamaModel:    document.getElementById('ollamaModel'),
    // 通用参数
    systemPrompt:   document.getElementById('systemPrompt'),
    tempSlider:     document.getElementById('temperatureSlider'),
    tempLabel:      document.getElementById('temperatureLabel'),
    // 按钮 & 状态
    saveConfigBtn:  document.getElementById('saveConfigBtn'),
    testBtn:        document.getElementById('testConnectionBtn'),
    testSpinner:    document.getElementById('testSpinner'),
    testStatus:     document.getElementById('testStatus'),
    saveIndicator:  document.getElementById('saveIndicator'),
    // 处理区
    selectedText:   document.getElementById('selectedText'),
    extractBtn:     document.getElementById('extractBtn'),
    instruction:    document.getElementById('instruction'),
    executeBtn:     document.getElementById('executeBtn'),
    executeSpinner: document.getElementById('executeSpinner'),
    executeStatus:  document.getElementById('executeStatus')
  };

  // ── 持久化 Key ────────────────────────────────────────────
  var STORAGE_KEY = 'officeai_config_v2';

  var DEFAULT_CONFIG = {
    provider:       'deepseek',
    deepseekUrl:    'https://api.deepseek.com',
    deepseekKey:    '',
    deepseekModel:  'deepseek-chat',
    temperature:    0.7,
    systemPrompt:   '你是一个Word排版与文本处理专家。用户会给你一段文本和一个指令。请直接输出处理后的文本。如果涉及格式调整，请用HTML标签包裹文本以表示格式（例如：<b>加粗</b>）。不要解释，只输出结果。',
    ollamaUrl:      'http://localhost:11434',
    ollamaModel:    'deepseek-r1:latest'
  };

  function loadConfig() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var saved = JSON.parse(raw);
        // 合并默认值（新加字段不会丢）
        return Object.assign({}, DEFAULT_CONFIG, saved);
      }
    } catch (e) { /* ignore */ }
    return Object.assign({}, DEFAULT_CONFIG);
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
      el.saveIndicator.classList.add('visible');
      setTimeout(function () {
        el.saveIndicator.classList.remove('visible');
      }, 2000);
    } catch (e) {
      showStatus(el.testStatus, 'error', '保存失败：本地存储已满，请清理浏览器数据。');
    }
  }

  function applyConfig(cfg) {
    el.provider.value       = cfg.provider;
    el.deepseekUrl.value    = cfg.deepseekUrl;
    el.deepseekKey.value    = cfg.deepseekKey;
    el.deepseekModel.value  = cfg.deepseekModel;
    el.systemPrompt.value   = cfg.systemPrompt;
    el.ollamaUrl.value      = cfg.ollamaUrl;
    el.ollamaModel.value    = cfg.ollamaModel;

    // Temperature slider
    var tempVal = Math.round(cfg.temperature * 10); // 0.7 → 7
    el.tempSlider.value = tempVal;
    el.tempLabel.textContent = cfg.temperature.toFixed(1);

    switchForm(cfg.provider);
  }

  function getCurrentConfig() {
    return {
      provider:       el.provider.value,
      deepseekUrl:    el.deepseekUrl.value.replace(/\/+$/, ''),
      deepseekKey:    el.deepseekKey.value.trim(),
      deepseekModel:  el.deepseekModel.value,
      temperature:    parseFloat(el.tempLabel.textContent),
      systemPrompt:   el.systemPrompt.value,
      ollamaUrl:      el.ollamaUrl.value.replace(/\/+$/, ''),
      ollamaModel:    el.ollamaModel.value.trim()
    };
  }

  // ── 表单切换 ──────────────────────────────────────────────
  function switchForm(provider) {
    if (provider === 'ollama') {
      el.deepseekForm.classList.remove('active');
      el.ollamaForm.classList.add('active');
    } else {
      el.ollamaForm.classList.remove('active');
      el.deepseekForm.classList.add('active');
    }
  }

  el.provider.addEventListener('change', function () {
    switchForm(el.provider.value);
  });

  // ── Temperature 滑块 ──────────────────────────────────────
  el.tempSlider.addEventListener('input', function () {
    var val = parseInt(el.tempSlider.value, 10) / 10;
    el.tempLabel.textContent = val.toFixed(1);

    // 颜色提示
    if (val <= 0.3) el.tempLabel.style.color = '#107c10';      // 绿 - 严谨
    else if (val <= 0.7) el.tempLabel.style.color = '#0078d4'; // 蓝 - 平衡
    else el.tempLabel.style.color = '#d83b01';                  // 橙 - 创意
  });

  // ── API Key 显隐切换 ──────────────────────────────────────
  el.toggleKeyBtn.addEventListener('click', function () {
    var input = el.deepseekKey;
    if (input.type === 'password') {
      input.type = 'text';
      el.toggleKeyBtn.innerHTML = '&#128064;'; // 睁眼
    } else {
      input.type = 'password';
      el.toggleKeyBtn.innerHTML = '&#128065;'; // 闭眼
    }
  });

  // ── Status 工具 ───────────────────────────────────────────
  function showStatus(elStatus, type, message) {
    elStatus.className = 'status ' + type;
    elStatus.textContent = message;
  }

  function clearStatus(elStatus) {
    elStatus.className = 'status';
    elStatus.textContent = '';
  }

  function setSpinner(spinner, active) {
    if (active) spinner.classList.add('active');
    else spinner.classList.remove('active');
  }

  function setButtonsDisabled(disabled) {
    el.testBtn.disabled = disabled;
    el.extractBtn.disabled = disabled;
    el.executeBtn.disabled = disabled;
  }

  // ══════════════════════════════════════════════════════════
  //  保存配置
  // ══════════════════════════════════════════════════════════
  el.saveConfigBtn.addEventListener('click', function () {
    var cfg = getCurrentConfig();

    // 基本校验
    if (cfg.provider === 'deepseek' && !cfg.deepseekKey) {
      showStatus(el.testStatus, 'error',
        '请在 API Key 字段中填写 DeepSeek 密钥。' +
        '获取地址: platform.deepseek.com → API Keys');
      return;
    }

    saveConfig(cfg);
    showStatus(el.testStatus, 'success', '配置已保存。可以开始使用 AI 文本处理功能。');
  });

  // ══════════════════════════════════════════════════════════
  //  测试连接 — DeepSeek API
  // ══════════════════════════════════════════════════════════
  el.testBtn.addEventListener('click', function () {
    clearStatus(el.testStatus);
    setSpinner(el.testSpinner, true);
    el.testBtn.disabled = true;

    var cfg = getCurrentConfig();

    if (cfg.provider === 'deepseek') {
      if (!cfg.deepseekKey) {
        showStatus(el.testStatus, 'error',
          '请先填写 DeepSeek API Key 再测试。');
        setSpinner(el.testSpinner, false);
        el.testBtn.disabled = false;
        return;
      }

      // DeepSeek API: POST /chat/completions（无 /v1 前缀！）
      var endpoint = cfg.deepseekUrl + '/chat/completions';

      fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + cfg.deepseekKey
        },
        body: JSON.stringify({
          model: cfg.deepseekModel,
          messages: [
            { role: 'user', content: 'Hi' }
          ],
          max_tokens: 5,
          stream: false
        })
      })
      .then(function (res) {
        if (!res.ok) {
          return res.text().then(function (body) {
            var msg = 'HTTP ' + res.status;
            try {
              var errData = JSON.parse(body);
              if (errData.error && errData.error.message) {
                msg = errData.error.message;
              }
            } catch (e) { /* use raw body */ }
            throw new Error(msg);
          });
        }
        return res.json();
      })
      .then(function () {
        showStatus(el.testStatus, 'success',
          '连接成功！DeepSeek API (' + cfg.deepseekModel + ') 可用。');
        // 测试成功同时保存
        saveConfig(cfg);
      })
      .catch(function (err) {
        console.error('DeepSeek test error:', err);
        showStatus(el.testStatus, 'error',
          '连接失败: ' + err.message);
      })
      .finally(function () {
        setSpinner(el.testSpinner, false);
        el.testBtn.disabled = false;
      });

    } else {
      // Ollama 测试
      var ollamaEndpoint = cfg.ollamaUrl + '/api/tags';
      fetch(ollamaEndpoint, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' }
      })
      .then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.json();
      })
      .then(function () {
        showStatus(el.testStatus, 'success',
          'Ollama 连接成功！请确保模型 ' + cfg.ollamaModel + ' 已下载。');
        saveConfig(cfg);
      })
      .catch(function (err) {
        showStatus(el.testStatus, 'error',
          'Ollama 连接失败: ' + err.message +
          '。请确认 Ollama 正在运行。');
      })
      .finally(function () {
        setSpinner(el.testSpinner, false);
        el.testBtn.disabled = false;
      });
    }
  });

  // ══════════════════════════════════════════════════════════
  //  提取 Word 选区文本
  // ══════════════════════════════════════════════════════════
  el.extractBtn.addEventListener('click', function () {
    clearStatus(el.executeStatus);
    el.extractBtn.disabled = true;

    Word.run(function (context) {
      var selection = context.document.getSelection();
      context.load(selection, 'text');
      return context.sync().then(function () {
        if (!selection.text || selection.text.trim().length === 0) {
          showStatus(el.executeStatus, 'info',
            '当前未选中任何文本，请先在 Word 文档中用鼠标选中一段文字。');
        } else {
          el.selectedText.value = selection.text;
          showStatus(el.executeStatus, 'success',
            '已提取选区文本（' + selection.text.length + ' 字符）。');
        }
      });
    }).catch(function (err) {
      console.error('Extract error:', err);
      var msg = '提取失败：';
      if (err instanceof OfficeExtension.Error) {
        msg += err.debugInfo ? err.debugInfo.message : err.message;
      } else {
        msg += err.message || '未知错误';
      }
      showStatus(el.executeStatus, 'error', msg);
    }).finally(function () {
      el.extractBtn.disabled = false;
    });
  });

  // ══════════════════════════════════════════════════════════
  //  执行 AI 指令
  // ══════════════════════════════════════════════════════════
  el.executeBtn.addEventListener('click', function () {
    clearStatus(el.executeStatus);

    var text = el.selectedText.value.trim();
    var instruction = el.instruction.value.trim();
    var cfg = getCurrentConfig();

    if (!text) {
      showStatus(el.executeStatus, 'error',
        '请先提取或粘贴待处理的文本。');
      return;
    }
    if (!instruction) {
      showStatus(el.executeStatus, 'error',
        '请输入 AI 指令，例如"翻译为英文"或"润色优化"。');
      return;
    }

    setSpinner(el.executeSpinner, true);
    setButtonsDisabled(true);

    var messages = [
      { role: 'system', content: cfg.systemPrompt },
      { role: 'user',  content: '文本：\n' + text + '\n\n指令：' + instruction }
    ];

    var fetchPromise;

    if (cfg.provider === 'ollama') {
      // ── Ollama ──
      fetchPromise = fetch(cfg.ollamaUrl + '/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: cfg.ollamaModel,
          messages: messages,
          stream: false
        })
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            throw new Error('Ollama HTTP ' + res.status + ': ' + t.slice(0, 300));
          });
        }
        return res.json();
      }).then(function (data) {
        if (data && data.message && data.message.content) {
          return data.message.content;
        }
        throw new Error('Ollama 返回数据格式异常，未找到 message.content');
      });

    } else {
      // ── DeepSeek API ──
      if (!cfg.deepseekKey) {
        throw new Error('请先在"大模型配置"中填写 DeepSeek API Key 并保存。');
      }

      var endpoint = cfg.deepseekUrl + '/chat/completions';
      var body = {
        model: cfg.deepseekModel,
        messages: messages,
        stream: false,
        temperature: cfg.temperature
      };

      // DeepSeek Reasoner (R1) 不支持 temperature 参数
      if (cfg.deepseekModel === 'deepseek-reasoner') {
        delete body.temperature;
      }

      fetchPromise = fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + cfg.deepseekKey
        },
        body: JSON.stringify(body)
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            var msg = 'HTTP ' + res.status;
            try {
              var errData = JSON.parse(t);
              if (errData.error && errData.error.message) {
                msg = errData.error.message;
              }
            } catch (e) { /* keep raw */ }
            throw new Error(msg);
          });
        }
        return res.json();
      }).then(function (data) {
        if (data && data.choices && data.choices.length > 0 &&
            data.choices[0].message && data.choices[0].message.content) {
          return data.choices[0].message.content;
        }
        throw new Error('DeepSeek 返回数据格式异常，未找到 choices[0].message.content');
      });
    }

    fetchPromise.then(function (aiContent) {
      showStatus(el.executeStatus, 'info', 'AI 处理完成，正在写回 Word 文档...');
      return insertHtmlToWord(aiContent);
    }).then(function () {
      showStatus(el.executeStatus, 'success', '处理完成！AI 生成的内容已替换原选区。');
    }).catch(function (err) {
      console.error('AI execution failed:', err);
      showStatus(el.executeStatus, 'error', '处理失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      setSpinner(el.executeSpinner, false);
      setButtonsDisabled(false);
    });
  });

  // ══════════════════════════════════════════════════════════
  //  将 HTML 写回 Word
  // ══════════════════════════════════════════════════════════
  function insertHtmlToWord(htmlContent) {
    return Word.run(function (context) {
      var selection = context.document.getSelection();
      selection.insertHtml(htmlContent, Word.InsertLocation.replace);
      return context.sync();
    });
  }

  // ══════════════════════════════════════════════════════════
  //  Office 初始化
  // ══════════════════════════════════════════════════════════
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      console.log('OfficeAI: Word host detected, initializing...');
      var cfg = loadConfig();
      applyConfig(cfg);

      // 初次加载时触发一次滑块颜色
      var tempVal = parseFloat(el.tempLabel.textContent);
      if (tempVal <= 0.3) el.tempLabel.style.color = '#107c10';
      else if (tempVal <= 0.7) el.tempLabel.style.color = '#0078d4';
      else el.tempLabel.style.color = '#d83b01';

      clearStatus(el.testStatus);
      clearStatus(el.executeStatus);
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
      showStatus(el.executeStatus, 'error',
        '此加载项仅支持 Microsoft Word。当前宿主: ' + info.host);
    }
  });

})();
