/* global Office, Word, console */

(function () {
  'use strict';

  /* ═══════════════════════════════════════════════════════════
     DOM 引用
     ═══════════════════════════════════════════════════════════ */
  var el = {};

  function cacheDom() {
    // 主界面
    el.mainPage         = document.getElementById('mainPage');
    el.settingsPage     = document.getElementById('settingsPage');
    el.openSettingsBtn  = document.getElementById('openSettingsBtn');
    el.backToMainBtn    = document.getElementById('backToMainBtn');
    el.readFullDocBtn   = document.getElementById('readFullDocBtn');
    el.extractSelBtn    = document.getElementById('extractSelectionBtn');
    el.readSpinner      = document.getElementById('readSpinner');
    el.documentText     = document.getElementById('documentText');
    el.instructionInput = document.getElementById('instructionInput');
    el.executeBtn       = document.getElementById('executeBtn');
    el.insertBtn        = document.getElementById('insertBtn');
    el.mainStatus       = document.getElementById('mainStatus');
    el.resultSection    = document.getElementById('resultSection');
    el.resultText       = document.getElementById('resultText');
    el.replaceBtn       = document.getElementById('replaceBtn');
    el.appendBtn        = document.getElementById('appendBtn');
    el.savedToast       = document.getElementById('savedToast');

    // 设置界面
    el.modelProvider    = document.getElementById('modelProvider');
    el.apiBaseUrl       = document.getElementById('apiBaseUrl');
    el.apiKey           = document.getElementById('apiKey');
    el.toggleKeyBtn     = document.getElementById('toggleKeyBtn');
    el.fetchModelsBtn   = document.getElementById('fetchModelsBtn');
    el.fetchModelsSpinner = document.getElementById('fetchModelsSpinner');
    el.fetchModelsStatus = document.getElementById('fetchModelsStatus');
    el.modelListGroup   = document.getElementById('modelListGroup');
    el.modelSelect      = document.getElementById('modelSelect');
    el.modelCountHint   = document.getElementById('modelCountHint');
    el.ollamaModelRow   = document.getElementById('ollamaModelRow');
    el.ollamaModelInput = document.getElementById('ollamaModelInput');
    el.tempSlider       = document.getElementById('temperatureSlider');
    el.tempDisplay      = document.getElementById('tempDisplay');
    el.systemPromptInput = document.getElementById('systemPromptInput');
    el.saveSettingsBtn  = document.getElementById('saveSettingsBtn');
    el.testConnBtn      = document.getElementById('testConnBtn');
    el.settingsStatus   = document.getElementById('settingsStatus');

    // 预设按钮
    el.presetBtns = document.querySelectorAll('.preset-btn[data-preset]');
  }

  /* ═══════════════════════════════════════════════════════════
     持久化配置
     ═══════════════════════════════════════════════════════════ */
  var STORAGE_KEY = 'officeai_config_v3';

  var DEFAULT_CONFIG = {
    provider:      'deepseek',
    apiBaseUrl:    'https://api.deepseek.com',
    apiKey:        '',
    model:         'deepseek-chat',
    modelList:     ['deepseek-chat', 'deepseek-reasoner'],
    temperature:   0.7,
    systemPrompt:  '你是一个集成于 Microsoft Word 的专业排版与文本处理专家。你的任务是对用户提供的文本执行精准操作。' +
                   '如果涉及格式调整，请使用 HTML 标签（如 <b>加粗</b>、<i>斜体</i>、<p style="text-indent:2em">首行缩进</p>、' +
                   '<h2>标题</h2> 等）包裹文本。请直接返回处理后的结果，不要解释，不要添加前缀说明。',
    ollamaModel:   'deepseek-r1:latest'
  };

  function loadConfig() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return Object.assign({}, DEFAULT_CONFIG, JSON.parse(raw));
    } catch (e) { /* ignore */ }
    return Object.assign({}, DEFAULT_CONFIG);
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
      showToast();
    } catch (e) {
      showStatus(el.settingsStatus, 'error', '保存失败：本地存储空间不足。');
    }
  }

  function applyConfigToUI(cfg) {
    el.modelProvider.value    = cfg.provider;
    el.apiBaseUrl.value       = cfg.apiBaseUrl;
    el.apiKey.value           = cfg.apiKey;
    el.systemPromptInput.value = cfg.systemPrompt;
    el.tempSlider.value       = Math.round(cfg.temperature * 10);
    el.tempDisplay.textContent = cfg.temperature.toFixed(1);
    el.ollamaModelInput.value = cfg.ollamaModel;

    // 填充模型下拉
    populateModelDropdown(cfg.modelList, cfg.model);
    updateProviderUI(cfg.provider);
  }

  function getConfigFromUI() {
    return {
      provider:      el.modelProvider.value,
      apiBaseUrl:    el.apiBaseUrl.value.replace(/\/+$/, ''),
      apiKey:        el.apiKey.value.trim(),
      model:         el.modelSelect.value,
      modelList:     getModelListFromDropdown(),
      temperature:   parseFloat(el.tempDisplay.textContent),
      systemPrompt:  el.systemPromptInput.value,
      ollamaModel:   el.ollamaModelInput.value.trim()
    };
  }

  function populateModelDropdown(modelList, selectedModel) {
    el.modelSelect.innerHTML = '';
    if (!modelList || modelList.length === 0) {
      el.modelSelect.innerHTML = '<option value="">请先获取模型列表...</option>';
      return;
    }
    modelList.forEach(function (m) {
      var opt = document.createElement('option');
      opt.value = m;
      opt.textContent = m;
      if (m === selectedModel) opt.selected = true;
      el.modelSelect.appendChild(opt);
    });
    el.modelCountHint.textContent = '共 ' + modelList.length + ' 个可用模型';
    el.modelListGroup.classList.add('visible');
  }

  function getModelListFromDropdown() {
    var list = [];
    for (var i = 0; i < el.modelSelect.options.length; i++) {
      var v = el.modelSelect.options[i].value;
      if (v) list.push(v);
    }
    return list.length > 0 ? list : DEFAULT_CONFIG.modelList;
  }

  /* ═══════════════════════════════════════════════════════════
     视图切换
     ═══════════════════════════════════════════════════════════ */
  function showPage(page) {
    if (page === 'main') {
      el.mainPage.classList.add('active');
      el.settingsPage.classList.remove('active');
    } else {
      el.settingsPage.classList.add('active');
      el.mainPage.classList.remove('active');
    }
  }

  el.openSettingsBtn.addEventListener('click', function () {
    showPage('settings');
  });

  el.backToMainBtn.addEventListener('click', function () {
    showPage('main');
  });

  /* ═══════════════════════════════════════════════════════════
     设置界面 — 服务商切换
     ═══════════════════════════════════════════════════════════ */
  function updateProviderUI(provider) {
    if (provider === 'deepseek') {
      el.apiBaseUrl.value = 'https://api.deepseek.com';
      el.apiBaseUrl.placeholder = 'https://api.deepseek.com';
      el.ollamaModelRow.style.display = 'none';
      el.fetchModelsBtn.style.display = '';
      el.modelSelect.style.display = '';
      el.modelCountHint.style.display = '';
    } else if (provider === 'ollama') {
      el.apiBaseUrl.value = 'http://localhost:11434';
      el.apiBaseUrl.placeholder = 'http://localhost:11434';
      el.ollamaModelRow.style.display = '';
      el.fetchModelsBtn.style.display = 'none';
      el.modelSelect.style.display = 'none';
      el.modelCountHint.style.display = 'none';
      el.modelListGroup.classList.add('visible');
    } else {
      // custom
      el.apiBaseUrl.value = '';
      el.apiBaseUrl.placeholder = 'https://your-api.com/v1';
      el.ollamaModelRow.style.display = 'none';
      el.fetchModelsBtn.style.display = '';
      el.modelSelect.style.display = '';
      el.modelCountHint.style.display = '';
    }
  }

  el.modelProvider.addEventListener('change', function () {
    updateProviderUI(el.modelProvider.value);
  });

  /* ═══════════════════════════════════════════════════════════
     设置界面 — 获取模型列表
     ═══════════════════════════════════════════════════════════ */
  el.fetchModelsBtn.addEventListener('click', function () {
    clearStatus(el.fetchModelsStatus);
    var baseUrl = el.apiBaseUrl.value.replace(/\/+$/, '');
    var apiKey = el.apiKey.value.trim();

    if (!apiKey) {
      showStatus(el.fetchModelsStatus, 'error',
        '请先填写 API Key。');
      return;
    }

    setSpinner(el.fetchModelsSpinner, true);
    el.fetchModelsBtn.disabled = true;
    showStatus(el.fetchModelsStatus, 'info', '正在获取模型列表...');

    // 尝试 OpenAI 兼容的 /models 端点
    var endpoint = baseUrl + '/models';

    fetch(endpoint, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      }
    })
    .then(function (res) {
      if (!res.ok) {
        // 如果是 404，可能端点不存在，尝试从 base URL 构造
        if (res.status === 404) {
          // 尝试不带 /v1 的情况：有些 API 的 models 在根路径
          return fetch(baseUrl.replace(/\/v1$/, '') + '/models', {
            method: 'GET',
            headers: {
              'Authorization': 'Bearer ' + apiKey,
              'Content-Type': 'application/json'
            }
          }).then(function (r2) {
            if (!r2.ok) throw new Error('MODELS_NOT_SUPPORTED');
            return r2.json();
          });
        }
        return res.text().then(function (t) {
          var msg = 'HTTP ' + res.status;
          try { var d = JSON.parse(t); if (d.error && d.error.message) msg = d.error.message; } catch (e) {}
          throw new Error(msg);
        });
      }
      return res.json();
    })
    .then(function (data) {
      var models = [];

      // OpenAI 格式: { object: "list", data: [{id: "model-name"}, ...] }
      if (data && data.data && Array.isArray(data.data)) {
        models = data.data
          .map(function (m) { return m.id; })
          .filter(function (id) { return id && typeof id === 'string'; });
      }
      // 数组格式: ["model-a", "model-b"]
      else if (Array.isArray(data)) {
        models = data.filter(function (m) { return typeof m === 'string'; });
      }
      // 对象格式 (Ollama): { models: [{name: "..."}, ...] }
      else if (data && data.models && Array.isArray(data.models)) {
        models = data.models
          .map(function (m) { return m.name || m.model || m.id; })
          .filter(Boolean);
      }

      if (models.length === 0) {
        throw new Error('MODELS_NOT_SUPPORTED');
      }

      // 过滤掉非对话模型（可选）
      var chatModels = models.filter(function (m) {
        var lower = m.toLowerCase();
        // 排除明显非对话模型
        if (lower.includes('embedding')) return false;
        if (lower.includes('moderation')) return false;
        if (lower.includes('whisper')) return false;
        if (lower.includes('tts')) return false;
        if (lower.includes('dall-e')) return false;
        return true;
      });

      if (chatModels.length === 0) chatModels = models;

      // 保存到配置
      var cfg = getConfigFromUI();
      cfg.modelList = chatModels;
      if (!chatModels.includes(cfg.model)) {
        cfg.model = chatModels[0];
      }
      saveConfig(cfg);
      populateModelDropdown(chatModels, cfg.model);

      showStatus(el.fetchModelsStatus, 'success',
        '成功获取 ' + chatModels.length + ' 个模型。');
    })
    .catch(function (err) {
      console.error('Fetch models error:', err);

      var fallbackMsg = '';
      if (err.message === 'MODELS_NOT_SUPPORTED') {
        // API 不支持 /models 端点，使用默认列表
        var provider = el.modelProvider.value;
        var fallbackList;
        if (provider === 'deepseek') {
          fallbackList = ['deepseek-chat', 'deepseek-reasoner'];
          fallbackMsg = 'DeepSeek 不支持模型列表查询。已加载默认模型。';
        } else {
          fallbackList = ['gpt-4o', 'gpt-4o-mini', 'gpt-3.5-turbo'];
          fallbackMsg = '该 API 不支持 /models 端点。已加载常见模型列表，可手动输入。';
        }
        var cfg = getConfigFromUI();
        cfg.modelList = fallbackList;
        saveConfig(cfg);
        populateModelDropdown(fallbackList, fallbackList[0]);
        el.modelListGroup.classList.add('visible');
        showStatus(el.fetchModelsStatus, 'warn', fallbackMsg);
      } else {
        showStatus(el.fetchModelsStatus, 'error',
          '获取失败: ' + (err.message || '网络错误'));
      }
    })
    .finally(function () {
      setSpinner(el.fetchModelsSpinner, false);
      el.fetchModelsBtn.disabled = false;
    });
  });

  /* ═══════════════════════════════════════════════════════════
     设置界面 — 保存 & 测试
     ═══════════════════════════════════════════════════════════ */
  el.saveSettingsBtn.addEventListener('click', function () {
    var cfg = getConfigFromUI();

    if ((cfg.provider === 'deepseek' || cfg.provider === 'custom') && !cfg.apiKey) {
      showStatus(el.settingsStatus, 'error',
        '请填写 API Key。获取地址: platform.deepseek.com → API Keys');
      return;
    }

    saveConfig(cfg);
    showStatus(el.settingsStatus, 'success', '配置已保存。');
  });

  el.testConnBtn.addEventListener('click', function () {
    clearStatus(el.settingsStatus);
    var cfg = getConfigFromUI();

    if ((cfg.provider === 'deepseek' || cfg.provider === 'custom') && !cfg.apiKey) {
      showStatus(el.settingsStatus, 'error', '请先填写 API Key。');
      return;
    }

    el.testConnBtn.disabled = true;
    showStatus(el.settingsStatus, 'info', '正在测试连接...');

    if (cfg.provider === 'ollama') {
      fetch(cfg.apiBaseUrl + '/api/tags', {
        method: 'GET',
        headers: { 'Content-Type': 'application/json' }
      })
      .then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.json();
      })
      .then(function () {
        showStatus(el.settingsStatus, 'success', 'Ollama 连接成功！');
        saveConfig(cfg);
      })
      .catch(function (err) {
        showStatus(el.settingsStatus, 'error', 'Ollama 连接失败: ' + err.message);
      })
      .finally(function () { el.testConnBtn.disabled = false; });
    } else {
      // DeepSeek / Custom: 发送最小对话请求测试
      var endpoint = cfg.apiBaseUrl + '/chat/completions';

      var body = {
        model: cfg.model || 'deepseek-chat',
        messages: [{ role: 'user', content: 'Hi' }],
        max_tokens: 5,
        stream: false
      };

      fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + cfg.apiKey
        },
        body: JSON.stringify(body)
      })
      .then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            var msg = 'HTTP ' + res.status;
            try {
              var d = JSON.parse(t);
              if (d.error && d.error.message) msg = d.error.message;
            } catch (e) {}
            throw new Error(msg);
          });
        }
        return res.json();
      })
      .then(function () {
        showStatus(el.settingsStatus, 'success',
          '连接成功！模型 ' + (cfg.model || 'default') + ' 可用。');
        saveConfig(cfg);
      })
      .catch(function (err) {
        showStatus(el.settingsStatus, 'error', '连接失败: ' + err.message);
      })
      .finally(function () { el.testConnBtn.disabled = false; });
    }
  });

  /* ═══════════════════════════════════════════════════════════
     设置界面 — 辅助交互
     ═══════════════════════════════════════════════════════════ */
  el.toggleKeyBtn.addEventListener('click', function () {
    if (el.apiKey.type === 'password') {
      el.apiKey.type = 'text';
      el.toggleKeyBtn.innerHTML = '&#128064;';
    } else {
      el.apiKey.type = 'password';
      el.toggleKeyBtn.innerHTML = '&#128065;';
    }
  });

  el.tempSlider.addEventListener('input', function () {
    var val = parseInt(el.tempSlider.value, 10) / 10;
    el.tempDisplay.textContent = val.toFixed(1);
    // 颜色提示
    if (val <= 0.3) el.tempDisplay.style.color = '#107c10';
    else if (val <= 0.7) el.tempDisplay.style.color = '#0078d4';
    else el.tempDisplay.style.color = '#d83b01';
  });

  /* ═══════════════════════════════════════════════════════════
     主界面 — 文档读取
     ═══════════════════════════════════════════════════════════ */

  // 读取全文
  el.readFullDocBtn.addEventListener('click', function () {
    clearStatus(el.mainStatus);
    setSpinner(el.readSpinner, true);
    el.readFullDocBtn.disabled = true;
    el.extractSelBtn.disabled = true;

    Word.run(function (context) {
      var body = context.document.body;
      context.load(body, 'text');
      return context.sync().then(function () {
        if (!body.text || body.text.trim().length === 0) {
          showStatus(el.mainStatus, 'info', '文档为空，请先输入内容。');
        } else {
          el.documentText.value = body.text;
          showStatus(el.mainStatus, 'success',
            '已读取全文（约 ' + body.text.length + ' 字符）。');
        }
      });
    }).catch(function (err) {
      showStatus(el.mainStatus, 'error', '读取全文失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      setSpinner(el.readSpinner, false);
      el.readFullDocBtn.disabled = false;
      el.extractSelBtn.disabled = false;
    });
  });

  // 提取选区
  el.extractSelBtn.addEventListener('click', function () {
    clearStatus(el.mainStatus);
    setSpinner(el.readSpinner, true);
    el.readFullDocBtn.disabled = true;
    el.extractSelBtn.disabled = true;

    Word.run(function (context) {
      var selection = context.document.getSelection();
      context.load(selection, 'text');
      return context.sync().then(function () {
        if (!selection.text || selection.text.trim().length === 0) {
          showStatus(el.mainStatus, 'info', '未选中文本。请先用鼠标在文档中选择一段文字。');
        } else {
          el.documentText.value = selection.text;
          showStatus(el.mainStatus, 'success',
            '已提取选区文本（' + selection.text.length + ' 字符）。');
        }
      });
    }).catch(function (err) {
      showStatus(el.mainStatus, 'error', '提取选区失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      setSpinner(el.readSpinner, false);
      el.readFullDocBtn.disabled = false;
      el.extractSelBtn.disabled = false;
    });
  });

  /* ═══════════════════════════════════════════════════════════
     主界面 — 排版预设
     ═══════════════════════════════════════════════════════════ */
  var PRESET_PROMPTS = {
    format_title: {
      label: '标题加粗居中',
      systemAddon: '你是一位 Word 排版专家。请将文档中的标题识别出来，加粗并居中。正文保持原样。请返回完整的、格式化后的文本。',
      prompt: '请将以下文本中的标题（如果有）加粗并居中显示，正文保持原样。返回完整的格式化文本：'
    },
    format_indent: {
      label: '正文首行缩进',
      systemAddon: '你是一位 Word 排版专家。请为正文段落添加首行缩进两字符。',
      prompt: '请为以下文本的正文段落添加首行缩进两字符的格式。标题和列表项保持不变。返回完整的格式化文本：'
    },
    format_spacing: {
      label: '段落间距调整',
      systemAddon: '你是一位 Word 排版专家。请调整段落间距使排版美观。',
      prompt: '请调整以下文本的段落间距，使整体排版更加美观、层次分明。段前距建议0.5行、段后距0.5行，标题上下各增加间距。返回完整的格式化文本：'
    },
    format_font: {
      label: '正文字号行距',
      systemAddon: '你是一位 Word 排版专家。请统一正文字号和行距。',
      prompt: '请将以下文本的正文统一为合适的中文字号（小四或12pt），行距1.5倍。标题用三号或16pt字体。返回完整的格式化文本：'
    },
    polish: {
      label: '校对润色',
      systemAddon: '你是一位专业的文字校对与润色专家。请修正语法错误、错别字，优化遣词造句，提升流畅度和逻辑性。不要改变原意。',
      prompt: '请校对并润色以下文本，修正错别字和语法问题：'
    },
    translate_cn2en: {
      label: '中译英',
      systemAddon: '你是一位专业的中英翻译专家。请准确、地道地将中文翻译为英文。',
      prompt: '请将以下中文翻译为英文：'
    },
    translate_en2cn: {
      label: '英译中',
      systemAddon: '你是一位专业的英中翻译专家。请准确、流畅地将英文翻译为中文。',
      prompt: '请将以下英文翻译为中文：'
    },
    summarize: {
      label: '生成摘要',
      systemAddon: '你是一位专业的文档摘要专家。请生成简洁但全面的摘要，保留核心观点和关键数据。',
      prompt: '请为以下文本生成清晰、条理分明的摘要（使用要点列表）：'
    }
  };

  el.presetBtns.forEach(function (btn) {
    btn.addEventListener('click', function () {
      var presetKey = btn.getAttribute('data-preset');
      var preset = PRESET_PROMPTS[presetKey];
      if (!preset) return;

      var text = el.documentText.value.trim();
      if (!text) {
        showStatus(el.mainStatus, 'error', '请先读取文档或粘贴文本。');
        return;
      }

      // 填入预设提示词
      el.instructionInput.value = preset.prompt;
      // 自动执行
      executeAI(text, preset.systemAddon, preset.prompt, preset.label);
    });
  });

  /* ═══════════════════════════════════════════════════════════
     主界面 — 自定义指令执行
     ═══════════════════════════════════════════════════════════ */
  el.executeBtn.addEventListener('click', function () {
    var text = el.documentText.value.trim();
    var instruction = el.instructionInput.value.trim();
    var cfg = getConfigFromUI();

    if (!text) {
      showStatus(el.mainStatus, 'error', '请先读取文档或粘贴文本。');
      return;
    }
    if (!instruction) {
      showStatus(el.mainStatus, 'error', '请输入 AI 指令或点击排版预设按钮。');
      return;
    }

    executeAI(text, cfg.systemPrompt, instruction, '自定义指令');
  });

  function executeAI(text, systemPrompt, userPrompt, actionLabel) {
    clearStatus(el.mainStatus);
    el.executeBtn.disabled = true;
    showStatus(el.mainStatus, 'info', 'AI 正在处理（' + actionLabel + '）...');

    var cfg = getConfigFromUI();
    var messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt + '\n\n"""\n' + text + '\n"""' }
    ];

    var fetchPromise;

    if (cfg.provider === 'ollama') {
      fetchPromise = fetch(cfg.apiBaseUrl + '/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: cfg.ollamaModel,
          messages: messages,
          stream: false
        })
      }).then(function (res) {
        if (!res.ok) return res.text().then(function (t) { throw new Error('HTTP ' + res.status + ': ' + t.slice(0, 300)); });
        return res.json();
      }).then(function (data) {
        if (data && data.message && data.message.content) return data.message.content;
        throw new Error('Ollama 返回格式异常');
      });
    } else {
      if (!cfg.apiKey) throw new Error('请先在设置中配置 API Key。');

      var endpoint = cfg.apiBaseUrl + '/chat/completions';
      var body = {
        model: cfg.model || 'deepseek-chat',
        messages: messages,
        stream: false,
        temperature: cfg.temperature
      };

      // R1 推理模型不支持 temperature
      if (cfg.model && (cfg.model.includes('reasoner') || cfg.model.includes('r1'))) {
        delete body.temperature;
      }

      fetchPromise = fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + cfg.apiKey
        },
        body: JSON.stringify(body)
      }).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (t) {
            var msg = 'HTTP ' + res.status;
            try { var d = JSON.parse(t); if (d.error && d.error.message) msg = d.error.message; } catch (e) {}
            throw new Error(msg);
          });
        }
        return res.json();
      }).then(function (data) {
        if (data && data.choices && data.choices[0] && data.choices[0].message) {
          return data.choices[0].message.content || '';
        }
        throw new Error('API 返回格式异常，未找到 choices[0].message.content');
      });
    }

    fetchPromise.then(function (content) {
      el.resultText.value = content;
      el.resultSection.style.display = '';
      showStatus(el.mainStatus, 'success', actionLabel + ' — 处理完成！可在下方预览结果。');
      el.resultSection.scrollIntoView({ behavior: 'smooth' });
    }).catch(function (err) {
      console.error('AI error:', err);
      showStatus(el.mainStatus, 'error', '处理失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      el.executeBtn.disabled = false;
    });
  }

  /* ═══════════════════════════════════════════════════════════
     主界面 — 结果操作（替换选区 / 追加到文末）
     ═══════════════════════════════════════════════════════════ */
  el.replaceBtn.addEventListener('click', function () {
    var content = el.resultText.value;
    if (!content) return;

    Word.run(function (context) {
      var selection = context.document.getSelection();
      selection.insertHtml(content, Word.InsertLocation.replace);
      return context.sync();
    }).then(function () {
      showStatus(el.mainStatus, 'success', '已替换选区。');
    }).catch(function (err) {
      showStatus(el.mainStatus, 'error', '替换失败: ' + (err.message || '未知错误'));
    });
  });

  el.appendBtn.addEventListener('click', function () {
    var content = el.resultText.value;
    if (!content) return;

    Word.run(function (context) {
      var body = context.document.body;
      // 在文档末尾插入
      body.insertHtml(content, Word.InsertLocation.end);
      return context.sync();
    }).then(function () {
      showStatus(el.mainStatus, 'success', '已追加到文档末尾。');
    }).catch(function (err) {
      showStatus(el.mainStatus, 'error', '追加失败: ' + (err.message || '未知错误'));
    });
  });

  /* ═══════════════════════════════════════════════════════════
     工具函数
     ═══════════════════════════════════════════════════════════ */
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

  function showToast() {
    el.savedToast.classList.add('show');
    setTimeout(function () { el.savedToast.classList.remove('show'); }, 1800);
  }

  /* ═══════════════════════════════════════════════════════════
     Office 初始化
     ═══════════════════════════════════════════════════════════ */
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      console.log('OfficeAI v1.2: Word host detected');
      cacheDom();

      // 加载配置
      var cfg = loadConfig();
      applyConfigToUI(cfg);

      // 初始化温度颜色
      var t = parseFloat(el.tempDisplay.textContent);
      if (t <= 0.3) el.tempDisplay.style.color = '#107c10';
      else if (t <= 0.7) el.tempDisplay.style.color = '#0078d4';
      else el.tempDisplay.style.color = '#d83b01';

      // 默认显示主界面
      showPage('main');
      clearStatus(el.mainStatus);
      clearStatus(el.settingsStatus);
      clearStatus(el.fetchModelsStatus);
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
    }
  });

})();
