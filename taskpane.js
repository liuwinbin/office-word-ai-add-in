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

    // 一键排版模块
    el.autoFormatBtn     = document.getElementById('autoFormatBtn');
    el.autoFormatSpinner = document.getElementById('autoFormatSpinner');
    el.formatProgress    = document.getElementById('formatProgress');
    el.formatProgressBar = document.getElementById('formatProgressBar');
    el.formatProgressText = document.getElementById('formatProgressText');
    el.formatStatus      = document.getElementById('formatStatus');

    // 排版参数设置
    el.formatCnFont       = document.getElementById('formatCnFont');
    el.formatEnFont       = document.getElementById('formatEnFont');
    el.formatFontSize     = document.getElementById('formatFontSize');
    el.formatLineSpacing  = document.getElementById('formatLineSpacing');
    el.formatIndent       = document.getElementById('formatIndent');
    el.formatCnEnSpacing  = document.getElementById('formatCnEnSpacing');
    el.formatRemoveEmpty  = document.getElementById('formatRemoveEmpty');
    el.formatSaveBtn      = document.getElementById('formatSaveBtn');
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
    ollamaModel:   'deepseek-r1:latest',

    // 原生排版参数
    formatOptions: {
      cnFont:           '微软雅黑',
      enFont:           'Times New Roman',
      fontSize:         12,
      lineSpacing:      1.5,
      indentChars:      2,
      enableCnEnSpacing: true,
      removeEmptyLines:  true
    }
  };

  function loadConfig() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var saved = JSON.parse(raw);
        var merged = Object.assign({}, DEFAULT_CONFIG, saved);
        if (saved.formatOptions) {
          merged.formatOptions = Object.assign({}, DEFAULT_CONFIG.formatOptions, saved.formatOptions);
        }
        return merged;
      }
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

    populateModelDropdown(cfg.modelList, cfg.model);
    updateProviderUI(cfg.provider);

    var fo = cfg.formatOptions || DEFAULT_CONFIG.formatOptions;
    el.formatCnFont.value       = fo.cnFont || '';
    el.formatEnFont.value       = fo.enFont || '';
    el.formatFontSize.value     = fo.fontSize || '';
    el.formatLineSpacing.value  = fo.lineSpacing || '';
    el.formatIndent.value       = fo.indentChars || '';
    el.formatCnEnSpacing.checked = fo.enableCnEnSpacing !== false;
    el.formatRemoveEmpty.checked = fo.removeEmptyLines !== false;
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
      ollamaModel:   el.ollamaModelInput.value.trim(),
      formatOptions: {
        cnFont:           el.formatCnFont.value.trim() || DEFAULT_CONFIG.formatOptions.cnFont,
        enFont:           el.formatEnFont.value.trim() || DEFAULT_CONFIG.formatOptions.enFont,
        fontSize:         parseFloat(el.formatFontSize.value) || DEFAULT_CONFIG.formatOptions.fontSize,
        lineSpacing:      parseFloat(el.formatLineSpacing.value) || DEFAULT_CONFIG.formatOptions.lineSpacing,
        indentChars:      parseFloat(el.formatIndent.value) || DEFAULT_CONFIG.formatOptions.indentChars,
        enableCnEnSpacing: el.formatCnEnSpacing.checked,
        removeEmptyLines:  el.formatRemoveEmpty.checked
      }
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
      el.apiBaseUrl.value = '';
      el.apiBaseUrl.placeholder = 'https://your-api.com/v1';
      el.ollamaModelRow.style.display = 'none';
      el.fetchModelsBtn.style.display = '';
      el.modelSelect.style.display = '';
      el.modelCountHint.style.display = '';
    }
  }

  /* ═══════════════════════════════════════════════════════════
     排版预设定义 (AI)
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

  /* ═══════════════════════════════════════════════════════════
     AI 执行核心
     ═══════════════════════════════════════════════════════════ */
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

  function showFormatToast(message) {
    el.savedToast.textContent = message;
    el.savedToast.classList.add('show');
    setTimeout(function () {
      el.savedToast.classList.remove('show');
      el.savedToast.textContent = '✓ 已保存';
    }, 3500);
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — 配置与状态
     ═══════════════════════════════════════════════════════════ */

  function getFormatOptions() {
    var cfg = loadConfig();
    return cfg.formatOptions || DEFAULT_CONFIG.formatOptions;
  }

  function createStats() {
    return { headings: 0, body: 0, cjkSpacing: 0, codeBlock: 0, list: 0, quote: 0, empties: 0, total: 0 };
  }

  function updateFormatProgress(stats, current, total, phase) {
    if (!el.formatProgress.classList.contains('active')) {
      el.formatProgress.classList.add('active');
    }
    var pct = total > 0 ? Math.round((current / total) * 100) : 0;
    el.formatProgressBar.style.width = pct + '%';
    el.formatProgressText.textContent = (phase || '') +
      ' 第 ' + current + '/' + total + ' 段' +
      ' (标题' + stats.headings + ' 正文' + stats.body +
      ' 代码' + stats.codeBlock + ' 列表' + stats.list +
      ' 间距' + stats.cjkSpacing + ' 引用' + stats.quote + ')';
  }

  function checkDocumentProtection(context) {
    var doc = context.document;
    doc.load('properties/protectionType');
    return context.sync().then(function () {
      var protType = doc.properties.protectionType;
      var typeMap = {
        2: '文档已启用「修订」模式，请先接受/拒绝所有修订',
        3: '文档仅允许批注，无法修改正文',
        4: '文档仅允许填写窗体',
        5: '文档已设为只读，请先取消文档保护'
      };
      if (protType && protType !== 1 && protType !== 0) {
        return { protected: true, reason: typeMap[protType] || '文档受保护，无法修改 (类型: ' + protType + ')' };
      }
      return { protected: false, reason: '' };
    });
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — AI 分析 Prompt
     ═══════════════════════════════════════════════════════════ */

  /**
   * AI 语义分析 System Prompt
   * AI 仅作为"文档结构理解引擎"，输出结构化 JSON 指令，绝不修改原文。
   */
  var SMART_FORMAT_SYSTEM_PROMPT =
    '你是文档结构分析引擎。你只分析文档的语义结构，不修改任何文本内容。\n\n' +
    '## 你的任务\n' +
    '分析以下段落数组，为每个段落生成一条排版指令。输出严格的 JSON 数组。\n\n' +
    '## 输出 JSON Schema\n' +
    '[{"index": 段落索引, "action": "操作名", ...额外参数}]\n\n' +
    '## 可用操作\n' +
    '1. setHeading — 语义标题（章/节/编号/摘要/引言/结论/参考文献/致谢/附录等）\n' +
    '   参数: "level": 1|2|3\n' +
    '   - level 1: 主标题（第X章、第X部分、摘要、绪论、结论、参考文献、致谢）\n' +
    '   - level 2: 次级标题（第X节、一、二、1. 数字编号）\n' +
    '   - level 3: 三级标题（1.1、(一)、小标题）\n' +
    '2. setBody — 普通正文段落\n' +
    '3. addSpaceBetweenCnEn — 段落中同时包含中文和英文/数字（需要加空格）\n' +
    '4. setCodeBlock — 代码块（等宽字体特征、缩进4空格、含代码关键词）\n' +
    '5. setList — 列表项\n' +
    '   参数: "type": "bullet"|"number"\n' +
    '6. setQuote — 引用块（引文、注释、说明性文字）\n\n' +
    '## 关键规则\n' +
    '- 每个非空段落必须恰好有一条指令\n' +
    '- 空段落（纯空白/零长度）跳过，不输出指令\n' +
    '- 中英文/数字混排的正文段落：优先标记 addSpaceBetweenCnEn\n' +
    '- 有中英文间距问题的标题：标记 addSpaceBetweenCnEn\n' +
    '- 标题不应同时标记为 setBody\n' +
    '- 列表项不标记为 setBody\n' +
    '- 输出必须是纯 JSON 数组，不要用 ```json 包裹，不要加任何解释文字';

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — CJK 间距处理（保留核心算法）
     ═══════════════════════════════════════════════════════════ */

  var CJK_RE = /[一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ]/;

  /**
   * 在中文与英文/数字之间插入半角空格。
   * 边界规则：引号/括号内侧不插入空格。
   * @param {string} text
   * @returns {string}
   */
  function addCnEnSpacingToText(text) {
    if (!text || text.length < 2) return text;

    text = text.replace(/([一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ])([a-zA-Z0-9])/g, '$1 $2');
    text = text.replace(/([a-zA-Z0-9])([一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ])/g, '$1 $2');

    var pairs = [
      ['“', '”'], ['‘', '’'],
      ['「', '」'], ['『', '』'],
      ['（', '）'], ['《', '》'],
      ['“', '”'], ['‘', '’']
    ];
    for (var i = 0; i < pairs.length; i++) {
      var left = pairs[i][0];
      var right = pairs[i][1];
      var reLeft = new RegExp('(' + left + ') ([a-zA-Z])', 'g');
      text = text.replace(reLeft, '$1$2');
      var reRight = new RegExp('([a-zA-Z0-9]) (' + right + ')', 'g');
      text = text.replace(reRight, '$1$2');
    }
    return text;
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — AI 调用与 JSON 解析
     ═══════════════════════════════════════════════════════════ */

  /**
   * 用段落文本数组构建给 AI 的用户 prompt
   * @param {Array<{index: number, text: string}>} paragraphTexts
   * @returns {string}
   */
  function buildSmartFormatPrompt(paragraphTexts) {
    var lines = ['请分析以下段落数组，输出 JSON 排版指令：', ''];
    for (var i = 0; i < paragraphTexts.length; i++) {
      var item = paragraphTexts[i];
      lines.push('[' + item.index + '] ' + item.text);
    }
    lines.push('');
    lines.push('直接输出 JSON 数组（不要 markdown 包裹）：');
    return lines.join('\n');
  }

  /**
   * 三层容错解析 AI 返回的 JSON
   * @param {string} rawText - AI 原始输出
   * @returns {Array|null} 解析后的指令数组，失败返回 null
   */
  function parseAIResponse(rawText) {
    if (!rawText || typeof rawText !== 'string') return null;

    // 第1层：直接 JSON.parse
    try {
      var result = JSON.parse(rawText.trim());
      if (Array.isArray(result)) return result;
    } catch (e) { /* continue */ }

    // 第2层：提取 ```json ... ``` 代码块
    var m = rawText.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (m) {
      try {
        var result2 = JSON.parse(m[1].trim());
        if (Array.isArray(result2)) return result2;
      } catch (e) { /* continue */ }
    }

    // 第3层：提取文本中最长的 [...] 数组
    var arrMatch = rawText.match(/\[[\s\S]*\]/);
    if (arrMatch) {
      try {
        var result3 = JSON.parse(arrMatch[0]);
        if (Array.isArray(result3)) return result3;
      } catch (e) { /* continue */ }
    }

    return null;
  }

  /**
   * 调用 AI 进行文档结构分析
   * @param {Array<{index: number, text: string}>} paragraphTexts
   * @returns {Promise<Array|null>} 指令数组，失败返回 null
   */
  function callAIForAnalysis(paragraphTexts) {
    var cfg = getConfigFromUI();

    if (cfg.provider !== 'ollama' && !cfg.apiKey) {
      showStatus(el.formatStatus, 'error', '请先在设置中配置 API Key。');
      return Promise.resolve(null);
    }

    var prompt = buildSmartFormatPrompt(paragraphTexts);
    var messages = [
      { role: 'system', content: SMART_FORMAT_SYSTEM_PROMPT },
      { role: 'user', content: prompt }
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
        if (!res.ok) {
          return res.text().then(function (t) {
            throw new Error('Ollama HTTP ' + res.status + ': ' + t.slice(0, 200));
          });
        }
        return res.json();
      }).then(function (data) {
        return (data && data.message && data.message.content) ? data.message.content : '';
      });
    } else {
      var endpoint = cfg.apiBaseUrl + '/chat/completions';
      var body = {
        model: cfg.model || 'deepseek-chat',
        messages: messages,
        stream: false,
        temperature: 0.1
      };
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
        return (data && data.choices && data.choices[0] && data.choices[0].message)
          ? data.choices[0].message.content || ''
          : '';
      });
    }

    return fetchPromise.then(function (rawText) {
      console.log('AI raw response length:', rawText.length);
      var instructions = parseAIResponse(rawText);
      if (!instructions || !Array.isArray(instructions) || instructions.length === 0) {
        console.warn('AI returned unparseable response, will use fallback');
        return null;
      }
      return instructions;
    }).catch(function (err) {
      console.error('AI analysis failed:', err);
      return null;
    });
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — 指令执行引擎
     ═══════════════════════════════════════════════════════════ */

  /**
   * executeSetHeading — 设置段落为 Word 内置标题样式
   * @param {Word.Paragraph} p
   * @param {number} level - 1/2/3
   * @param {object} stats
   */
  function executeSetHeading(p, level, stats) {
    var styleMap = { 1: 'Heading 1', 2: 'Heading 2', 3: 'Heading 3' };
    var styleName = styleMap[level] || 'Heading 2';
    try {
      p.style = styleName;
      stats.headings++;
    } catch (e) {
      console.warn('executeSetHeading failed:', e);
    }
  }

  /**
   * executeSetBody — 对单个正文段落应用字体/字号/行距/缩进/段间距
   * @param {Word.Paragraph} p
   * @param {object} stats
   * @param {object} opts - 排版参数
   */
  function executeSetBody(p, stats, opts) {
    try {
      var indentPt = opts.fontSize * opts.indentChars;
      var lineSpacingPt = opts.fontSize * opts.lineSpacing * 1.2;

      p.font.name = opts.cnFont;
      p.font.size = opts.fontSize;
      p.paragraphFormat.firstLineIndent = indentPt;
      p.paragraphFormat.spaceBefore = 0;
      p.paragraphFormat.spaceAfter = 0;
      p.paragraphFormat.lineSpacing = lineSpacingPt;
      stats.body++;
    } catch (e) {
      console.warn('executeSetBody failed:', e);
    }
  }

  /**
   * executeAddCnEnSpacing — 对单个段落执行中英文间距插入
   * @param {Word.Paragraph} p
   * @param {object} stats
   */
  function executeAddCnEnSpacing(p, stats) {
    var originalText = p.text || '';
    if (originalText.trim().length === 0) return;
    var corrected = addCnEnSpacingToText(originalText);
    if (corrected !== originalText) {
      stats.cjkSpacing++;
      try {
        p.insertText(corrected, Word.InsertLocation.replace);
      } catch (e) {
        console.warn('executeAddCnEnSpacing insertText failed:', e);
      }
    }
  }

  /**
   * executeSetCodeBlock — 对段落应用代码块格式（等宽字体、小字号、左缩进）
   * @param {Word.Paragraph} p
   * @param {object} stats
   */
  function executeSetCodeBlock(p, stats) {
    try {
      p.font.name = 'Consolas';
      p.font.size = 10.5;
      p.paragraphFormat.firstLineIndent = 0;
      p.paragraphFormat.leftIndent = 28; // ~4字符缩进
      p.paragraphFormat.spaceBefore = 2;
      p.paragraphFormat.spaceAfter = 2;
      stats.codeBlock++;
    } catch (e) {
      console.warn('executeSetCodeBlock failed:', e);
    }
  }

  /**
   * executeSetList — 对段落应用项目符号或编号列表
   * @param {Word.Paragraph} p
   * @param {string} type - "bullet" | "number"
   * @param {object} stats
   */
  function executeSetList(p, type, stats) {
    try {
      if (type === 'number') {
        p.startNewNumberedList();
      } else {
        p.startNewBullet();
      }
      stats.list++;
    } catch (e) {
      console.warn('executeSetList failed:', e);
    }
  }

  /**
   * executeSetQuote — 对段落应用引用格式（斜体、左右缩进、灰色）
   * @param {Word.Paragraph} p
   * @param {object} stats
   */
  function executeSetQuote(p, stats) {
    try {
      p.font.italic = true;
      p.paragraphFormat.leftIndent = 36;
      p.paragraphFormat.rightIndent = 36;
      p.paragraphFormat.firstLineIndent = 0;
      // 灰色字体
      try { p.font.color = '#666666'; } catch (e) { /* 某些文档不支持颜色 */ }
      stats.quote++;
    } catch (e) {
      console.warn('executeSetQuote failed:', e);
    }
  }

  /**
   * 指令分发器：根据 action 类型调用对应的执行函数
   * @param {Word.RequestContext} context
   * @param {Array<Word.Paragraph>} items - 当前批次的段落对象数组
   * @param {Array<object>} instructions - AI 返回的指令数组
   * @param {number} batchStart - 当前批次的起始段落索引
   * @param {number} batchEnd - 当前批次的结束段落索引（不含）
   * @param {object} stats
   * @param {object} opts
   */
  function executeInstructions(context, items, instructions, batchStart, batchEnd, stats, opts) {
    // 建立指令索引：paragraph index → instruction
    var instructionMap = {};
    for (var i = 0; i < instructions.length; i++) {
      var instr = instructions[i];
      if (instr && typeof instr.index === 'number' && instr.action) {
        instructionMap[instr.index] = instr;
      }
    }

    // 遍历当前批次的所有段落
    for (var pIdx = batchStart; pIdx < batchEnd; pIdx++) {
      var localIdx = pIdx - batchStart;
      var p = items[localIdx];
      var text = (p.text || '').trim();

      // 跳过空段落
      if (text.length === 0) continue;

      var instr = instructionMap[pIdx];

      if (!instr) {
        // AI 未给出指令 → 默认按正文处理
        if (opts.enableCnEnSpacing) {
          executeAddCnEnSpacing(p, stats);
        }
        executeSetBody(p, stats, opts);
        continue;
      }

      // 先执行中英文间距（如果启用且在指令中）
      if (opts.enableCnEnSpacing && instr.action === 'addSpaceBetweenCnEn') {
        executeAddCnEnSpacing(p, stats);
      }

      // 执行指定操作
      switch (instr.action) {
        case 'setHeading':
          executeSetHeading(p, instr.level || 2, stats);
          // 标题也需要正文格式兜底（字体统一）
          try {
            p.font.name = opts.cnFont;
            p.font.size = opts.fontSize + 2; // 标题字号略大
          } catch (e) {}
          break;

        case 'setBody':
          executeSetBody(p, stats, opts);
          break;

        case 'addSpaceBetweenCnEn':
          // 已在上方处理
          if (!opts.enableCnEnSpacing) break;
          // 如果 AI 只标记了 addSpaceBetweenCnEn 没有正文格式，补充正文格式
          executeSetBody(p, stats, opts);
          break;

        case 'setCodeBlock':
          executeSetCodeBlock(p, stats);
          break;

        case 'setList':
          executeSetList(p, instr.type || 'bullet', stats);
          break;

        case 'setQuote':
          executeSetQuote(p, stats);
          break;

        default:
          // 未知操作 → 默认正文
          executeSetBody(p, stats, opts);
          break;
      }
    }
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — 清除连续空行（保留）
     ═══════════════════════════════════════════════════════════ */

  /**
   * 删除连续空段落，保留最多 1 个空行。从后往前遍历以确保安全删除。
   * @param {Word.RequestContext} context
   * @param {Word.ParagraphCollection} paragraphs
   * @param {object} stats
   */
  function removeEmptyLines(context, paragraphs, stats) {
    var items = paragraphs.items;
    var consecutiveEmpty = 0;

    for (var i = items.length - 1; i >= 0; i--) {
      var text = (items[i].text || '').trim();
      if (text.length === 0) {
        consecutiveEmpty++;
        if (consecutiveEmpty > 1) {
          stats.empties++;
          try {
            items[i].delete();
          } catch (e) {
            console.warn('removeEmptyLines: paragraph ' + i + ' delete failed', e);
          }
        }
      } else {
        consecutiveEmpty = 0;
      }
    }
  }

  /* ═══════════════════════════════════════════════════════════
     智能排版模块 — 主编排器（混合架构：AI 分析 + Office.js 执行）
     ═══════════════════════════════════════════════════════════ */

  var BATCH_SIZE = 80;

  /**
   * 一键智能排版：
   *   Phase 0: 加载段落，检查文档保护
   *   Phase 1: AI 语义分析 → 输出 JSON 指令数组
   *   Phase 2: Office.js 原生执行指令（分批 + 进度条）
   *   Phase 3: 清除连续空行 → 显示统计
   */
  function smartFormatDocument() {
    clearStatus(el.formatStatus);
    clearStatus(el.mainStatus);

    var opts = getFormatOptions();

    el.autoFormatBtn.disabled = true;
    setSpinner(el.autoFormatSpinner, true);
    showStatus(el.formatStatus, 'info', '正在加载文档...');

    // Phase 0: 加载段落 + 检查保护
    Word.run(function (context) {
      var body = context.document.body;
      var paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        var items = paragraphs.items;
        var total = items.length;

        if (total === 0) {
          showStatus(el.formatStatus, 'info', '文档无内容，无需排版。');
          return;
        }

        return checkDocumentProtection(context).then(function (protResult) {
          if (protResult.protected) {
            showStatus(el.formatStatus, 'error', protResult.reason);
            return;
          }

          var stats = createStats();
          stats.total = total;

          // 提取段落文本用于 AI 分析（最多取前120字符、最多200段避免 token 爆炸）
          var MAX_AI_PARAGRAPHS = 200;
          var paragraphTexts = [];
          var limit = Math.min(total, MAX_AI_PARAGRAPHS);
          for (var i = 0; i < limit; i++) {
            var t = (items[i].text || '').trim();
            if (t.length > 0) {
              paragraphTexts.push({ index: i, text: t.slice(0, 120) });
            }
          }

          showStatus(el.formatStatus, 'info', '文档共 ' + total + ' 段' +
            (total > MAX_AI_PARAGRAPHS ? ' (前' + MAX_AI_PARAGRAPHS + '段用于AI分析)' : '') +
            '，AI 正在分析结构...');
          updateFormatProgress(stats, 0, total, 'AI 分析中...');

          // Phase 1: AI 语义分析
          return callAIForAnalysis(paragraphTexts).then(function (instructions) {
            if (!instructions || instructions.length === 0) {
              showStatus(el.formatStatus, 'warn',
                'AI 分析失败，使用规则引擎回退。将基于正则匹配执行基本排版...');
              instructions = generateFallbackInstructions(items, total);
            }

            showStatus(el.formatStatus, 'info',
              'AI 已生成 ' + instructions.length + ' 条排版指令，正在原生执行...');

            // Phase 2: 分批执行
            var batchCount = Math.ceil(total / BATCH_SIZE);
            return executeBatches(0, batchCount, total, instructions, stats, opts);
          });
        });
      });
    }).catch(function (err) {
      console.error('smartFormatDocument error:', err);
      showStatus(el.formatStatus, 'error', '排版失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      el.autoFormatBtn.disabled = false;
      setSpinner(el.autoFormatSpinner, false);
    });
  }

  /**
   * 递归分批执行排版指令
   */
  function executeBatches(batchIndex, batchCount, total, instructions, stats, opts) {
    if (batchIndex >= batchCount) {
      return finalizeSmartFormat(stats, opts);
    }

    var start = batchIndex * BATCH_SIZE;
    var end = Math.min(start + BATCH_SIZE, total);

    return Word.run(function (context) {
      var body = context.document.body;
      var paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        var allItems = paragraphs.items;

        // 执行指令
        executeInstructions(context, allItems, instructions, start, end, stats, opts);

        // 清除空行（在每批中执行，确保增量清理）
        if (opts.removeEmptyLines) {
          removeEmptyLines(context, paragraphs, stats);
        }

        return context.sync();
      });
    }).then(function () {
      updateFormatProgress(stats, end, total, '执行中...');

      return executeBatches(batchIndex + 1, batchCount, total, instructions, stats, opts);
    }).catch(function (err) {
      console.error('Batch ' + batchIndex + ' error:', err);
      showStatus(el.formatStatus, 'warn',
        '第 ' + (batchIndex + 1) + ' 批出错: ' + (err.message || '未知') + '，继续...');
      return executeBatches(batchIndex + 1, batchCount, total, instructions, stats, opts);
    });
  }

  /**
   * AI 失败时的回退方案：基于简单正则规则生成指令
   */
  function generateFallbackInstructions(items, total) {
    var instructions = [];
    var HEADING_RE = [
      { regex: /^第[一二三四五六七八九十百千\d]+[章篇部]/, level: 1 },
      { regex: /^第[一二三四五六七八九十百千\d]+节/, level: 2 },
      { regex: /^[一二三四五六七八九十]+[、．.]/, level: 2 },
      { regex: /^\d+[\.\、]/, level: 2 },
      { regex: /^\d+\.\d+/, level: 3 },
      { regex: /^（[一二三四五六七八九十\d]+）/, level: 3 }
    ];

    for (var i = 0; i < total; i++) {
      var text = (items[i].text || '').trim();
      if (text.length === 0) continue;

      var isHeading = false;
      for (var j = 0; j < HEADING_RE.length; j++) {
        if (HEADING_RE[j].regex.test(text.replace(/^\s+/, '').slice(0, 60))) {
          instructions.push({ index: i, action: 'setHeading', level: HEADING_RE[j].level });
          isHeading = true;
          break;
        }
      }
      if (!isHeading) {
        instructions.push({ index: i, action: 'setBody' });
      }
    }
    return instructions;
  }

  /**
   * 排版完成后的收尾工作
   */
  function finalizeSmartFormat(stats, opts) {
    setTimeout(function () {
      el.formatProgress.classList.remove('active');
      el.formatProgressBar.style.width = '0%';
    }, 600);

    var parts = [];
    if (stats.headings > 0) parts.push('标题 ' + stats.headings + ' 处');
    if (stats.body > 0) parts.push('正文 ' + stats.body + ' 段');
    if (stats.cjkSpacing > 0) parts.push('中英文间距 ' + stats.cjkSpacing + ' 处');
    if (stats.codeBlock > 0) parts.push('代码块 ' + stats.codeBlock + ' 处');
    if (stats.list > 0) parts.push('列表 ' + stats.list + ' 项');
    if (stats.quote > 0) parts.push('引用 ' + stats.quote + ' 处');
    if (stats.empties > 0) parts.push('清除空行 ' + stats.empties + ' 行');

    var msg;
    if (parts.length > 0) {
      msg = '✓ 排版完成：已规范 ' + parts.join('，');
    } else {
      msg = '✓ 排版完成：文档格式已符合规范';
    }

    showStatus(el.formatStatus, 'success', msg);
    showFormatToast(msg);
  }

  /* ═══════════════════════════════════════════════════════════
     事件绑定 — 所有事件监听器集中注册
     必须在 cacheDom() 之后调用！
     ═══════════════════════════════════════════════════════════ */
  function bindEvents() {

    // --- 视图切换 ---
    el.openSettingsBtn.addEventListener('click', function () {
      showPage('settings');
    });

    el.backToMainBtn.addEventListener('click', function () {
      showPage('main');
    });

    // --- 服务商切换 ---
    el.modelProvider.addEventListener('change', function () {
      updateProviderUI(el.modelProvider.value);
    });

    // --- 获取模型列表 ---
    el.fetchModelsBtn.addEventListener('click', function () {
      clearStatus(el.fetchModelsStatus);
      var baseUrl = el.apiBaseUrl.value.replace(/\/+$/, '');
      var apiKey = el.apiKey.value.trim();

      if (!apiKey) {
        showStatus(el.fetchModelsStatus, 'error', '请先填写 API Key。');
        return;
      }

      setSpinner(el.fetchModelsSpinner, true);
      el.fetchModelsBtn.disabled = true;
      showStatus(el.fetchModelsStatus, 'info', '正在获取模型列表...');

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
          if (res.status === 404) {
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

        if (data && data.data && Array.isArray(data.data)) {
          models = data.data
            .map(function (m) { return m.id; })
            .filter(function (id) { return id && typeof id === 'string'; });
        }
        else if (Array.isArray(data)) {
          models = data.filter(function (m) { return typeof m === 'string'; });
        }
        else if (data && data.models && Array.isArray(data.models)) {
          models = data.models
            .map(function (m) { return m.name || m.model || m.id; })
            .filter(Boolean);
        }

        if (models.length === 0) {
          throw new Error('MODELS_NOT_SUPPORTED');
        }

        var chatModels = models.filter(function (m) {
          var lower = m.toLowerCase();
          if (lower.includes('embedding')) return false;
          if (lower.includes('moderation')) return false;
          if (lower.includes('whisper')) return false;
          if (lower.includes('tts')) return false;
          if (lower.includes('dall-e')) return false;
          return true;
        });

        if (chatModels.length === 0) chatModels = models;

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

    // --- 保存配置 ---
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

    // --- 测试连接 ---
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

    // --- API Key 显隐切换 ---
    el.toggleKeyBtn.addEventListener('click', function () {
      if (el.apiKey.type === 'password') {
        el.apiKey.type = 'text';
        el.toggleKeyBtn.innerHTML = '&#128064;';
      } else {
        el.apiKey.type = 'password';
        el.toggleKeyBtn.innerHTML = '&#128065;';
      }
    });

    // --- Temperature 滑块 ---
    el.tempSlider.addEventListener('input', function () {
      var val = parseInt(el.tempSlider.value, 10) / 10;
      el.tempDisplay.textContent = val.toFixed(1);
      if (val <= 0.3) el.tempDisplay.style.color = '#107c10';
      else if (val <= 0.7) el.tempDisplay.style.color = '#0078d4';
      else el.tempDisplay.style.color = '#d83b01';
    });

    // --- 读取全文 ---
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

    // --- 提取选区 ---
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

    // --- 排版预设按钮 ---
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

        el.instructionInput.value = preset.prompt;
        executeAI(text, preset.systemAddon, preset.prompt, preset.label);
      });
    });

    // --- 自定义指令执行 ---
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

    // --- 替换选区 ---
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

    // --- 追加到文末 ---
    el.appendBtn.addEventListener('click', function () {
      var content = el.resultText.value;
      if (!content) return;

      Word.run(function (context) {
        var body = context.document.body;
        body.insertHtml(content, Word.InsertLocation.end);
        return context.sync();
      }).then(function () {
        showStatus(el.mainStatus, 'success', '已追加到文档末尾。');
      }).catch(function (err) {
        showStatus(el.mainStatus, 'error', '追加失败: ' + (err.message || '未知错误'));
      });
    });

    // --- 一键智能排版（混合架构：AI 语义分析 + Office.js 原生执行） ---
    el.autoFormatBtn.addEventListener('click', function () {
      smartFormatDocument();
    });

    // --- 保存排版参数 ---
    el.formatSaveBtn.addEventListener('click', function () {
      var cfg = getConfigFromUI();
      saveConfig(cfg);
      showStatus(el.settingsStatus, 'success', '排版参数已保存。');
    });

  } // end bindEvents

  /* ═══════════════════════════════════════════════════════════
     Office 初始化
     ═══════════════════════════════════════════════════════════ */
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      console.log('OfficeAI v1.5: Word host detected');

      // 1. 缓存 DOM 引用（必须在 bindEvents 之前）
      cacheDom();

      // 2. 注册所有事件监听器
      bindEvents();

      // 3. 加载配置
      var cfg = loadConfig();
      applyConfigToUI(cfg);

      // 4. 初始化温度颜色
      var t = parseFloat(el.tempDisplay.textContent);
      if (t <= 0.3) el.tempDisplay.style.color = '#107c10';
      else if (t <= 0.7) el.tempDisplay.style.color = '#0078d4';
      else el.tempDisplay.style.color = '#d83b01';

      // 5. 默认显示主界面
      showPage('main');
      clearStatus(el.mainStatus);
      clearStatus(el.settingsStatus);
      clearStatus(el.fetchModelsStatus);
      clearStatus(el.formatStatus);

      console.log('OfficeAI v1.5: Initialization complete');
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
    }
  });

})();
