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
        // 深度合并 formatOptions
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

    // 填充模型下拉
    populateModelDropdown(cfg.modelList, cfg.model);
    updateProviderUI(cfg.provider);

    // 填充排版参数
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
      // custom
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

  /**
   * 显示带统计信息的排版完成 Toast
   * @param {string} message - 统计信息文本
   */
  function showFormatToast(message) {
    el.savedToast.textContent = message;
    el.savedToast.classList.add('show');
    setTimeout(function () {
      el.savedToast.classList.remove('show');
      el.savedToast.textContent = '✓ 已保存'; // 恢复默认文本
    }, 3500);
  }

  /* ═══════════════════════════════════════════════════════════
     原生排版模块 — 配置与状态
     ═══════════════════════════════════════════════════════════ */

  /**
   * 从已保存配置中获取排版参数，合并默认值
   * @returns {object} 排版参数对象
   */
  function getFormatOptions() {
    var cfg = loadConfig();
    return cfg.formatOptions || DEFAULT_CONFIG.formatOptions;
  }

  /**
   * 排版统计收集器
   * @typedef {{headings: number, body: number, cjkSpacing: number, empties: number, total: number}} FormatStats
   */

  /** @returns {FormatStats} */
  function createStats() {
    return { headings: 0, body: 0, cjkSpacing: 0, empties: 0, total: 0 };
  }

  /** @param {FormatStats} stats @param {number} current @param {number} total */
  function updateFormatProgress(stats, current, total) {
    if (!el.formatProgress.classList.contains('active')) {
      el.formatProgress.classList.add('active');
    }
    var pct = total > 0 ? Math.round((current / total) * 100) : 0;
    el.formatProgressBar.style.width = pct + '%';
    el.formatProgressText.textContent =
      '正在处理 第 ' + current + '/' + total + ' 段...' +
      ' (标题' + stats.headings + ' 正文' + stats.body +
      ' 间距' + stats.cjkSpacing + ' 空行' + stats.empties + ')';
  }

  /**
   * 检查文档是否受保护（只读/修订/填写窗体等）
   * @param {Word.RequestContext} context
   * @returns {Promise<{protected: boolean, reason: string}>}
   */
  function checkDocumentProtection(context) {
    var doc = context.document;
    doc.load('properties/protectionType');
    return context.sync().then(function () {
      var protType = doc.properties.protectionType;
      // Word.ProtectionType: 0=Unknown, 1=NoProtection, 2=AllowOnlyRevisions,
      //   3=AllowOnlyComments, 4=AllowOnlyFormFields, 5=ReadOnly
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
     原生排版模块 — 标题检测
     ═══════════════════════════════════════════════════════════ */

  /**
   * 标题检测正则规则（优先级从高到低）
   * 匹配「第X章/第X节」、中文数字编号、阿拉伯数字编号等模式
   */
  var HEADING_PATTERNS = [
    // 第X章 / 第X篇 / 第X部 → Heading 1
    { regex: /^第[一二三四五六七八九十百千\d]+[章篇部]/, level: 1 },
    // 第X节 → Heading 2
    { regex: /^第[一二三四五六七八九十百千\d]+节/, level: 2 },
    // 一、 二、 三、 → Heading 2
    { regex: /^[一二三四五六七八九十]+[、．.]/, level: 2 },
    // 1. 2. 3. → Heading 2
    { regex: /^\d+[\.\、]/, level: 2 },
    // 1.1  1.2.3 → Heading 3
    { regex: /^\d+\.\d+/, level: 3 },
    // （一）（二）→ Heading 3
    { regex: /^（[一二三四五六七八九十\d]+）/, level: 3 }
  ];

  /**
   * 判断段落文本是否匹配标题模式
   * @param {string} text - 段落纯文本（trim 后）
   * @returns {{isHeading: boolean, level: number}}
   */
  function matchHeading(text) {
    for (var i = 0; i < HEADING_PATTERNS.length; i++) {
      if (HEADING_PATTERNS[i].regex.test(text)) {
        return { isHeading: true, level: HEADING_PATTERNS[i].level };
      }
    }
    return { isHeading: false, level: 0 };
  }

  /**
   * 检测并应用标题样式
   * @param {Word.RequestContext} context
   * @param {Word.ParagraphCollection} paragraphs
   * @param {FormatStats} stats
   * @param {object} opts - 排版参数
   * @returns {Set<number>} 被标记为标题的段落索引集合
   */
  function detectAndStyleHeadings(context, paragraphs, stats, opts) {
    var items = paragraphs.items;
    var headingIndices = new Array(items.length); // sparse: true at heading positions

    for (var i = 0; i < items.length; i++) {
      // 只读取前60个字符进行匹配（标题通常较短）
      var rawText = items[i].text || '';
      var trimmed = rawText.replace(/^\s+/, '').slice(0, 60);

      var match = matchHeading(trimmed);
      if (match.isHeading) {
        headingIndices[i] = true;
        stats.headings++;

        var styleName;
        if (match.level === 1) styleName = 'Heading 1';
        else if (match.level === 2) styleName = 'Heading 2';
        else styleName = 'Heading 3';

        try {
          items[i].style = styleName;
        } catch (e) {
          // 某些文档可能不支持该样式名，静默跳过
          headingIndices[i] = false;
          stats.headings--;
        }
      }
    }
    return headingIndices;
  }

  /* ═══════════════════════════════════════════════════════════
     原生排版模块 — 正文格式
     ═══════════════════════════════════════════════════════════ */

  /**
   * 对非标题段落应用正文格式：字体、字号、行距、首行缩进、段间距
   * @param {Word.RequestContext} context
   * @param {Word.ParagraphCollection} paragraphs
   * @param {FormatStats} stats
   * @param {object} opts - 排版参数
   * @param {Array<boolean>} headingIndices - 标题索引标记
   */
  function applyBodyFormat(context, paragraphs, stats, opts, headingIndices) {
    var items = paragraphs.items;
    var indentPt = opts.fontSize * opts.indentChars;   // 缩进点数
    var lineSpacingPt = opts.fontSize * opts.lineSpacing * 1.2; // 行距点数（×1.2 Word行距系数）

    for (var i = 0; i < items.length; i++) {
      if (headingIndices[i]) continue; // 跳过标题段落

      var p = items[i];
      var text = (p.text || '').trim();

      // 跳过空段落
      if (text.length === 0) continue;

      stats.body++;

      try {
        // 字体（中文字体名同时作用于 CJK 和 Latin，Office.js 会自动处理）
        p.font.name = opts.cnFont;
        p.font.size = opts.fontSize;
        // 段落格式
        p.paragraphFormat.firstLineIndent = indentPt;
        p.paragraphFormat.spaceBefore = 0;
        p.paragraphFormat.spaceAfter = 0;
        p.paragraphFormat.lineSpacing = lineSpacingPt;
      } catch (e) {
        // 单个段落失败不影响整体
        console.warn('applyBodyFormat: paragraph ' + i + ' failed', e);
      }
    }
  }

  /* ═══════════════════════════════════════════════════════════
     原生排版模块 — 中英文间距
     ═══════════════════════════════════════════════════════════ */

  /**
   * CJK 字符 Unicode 范围
   */
  var CJK_RE = /[一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ]/;

  /**
   * 在中文与英文/数字之间插入半角空格
   *
   * 边界规则（避免错误加空格）：
   * - 不在引号/括号内侧加空格（如 "中文"abc → 不加）
   * - 中文标点后跟英文不加空格（如 "中文。English" → "中文。 English"）
   *   实际上中文句号后加英文应该加空格，这里指的是中文引号内侧
   *
   * @param {string} text - 原始文本
   * @returns {string} 处理后的文本
   */
  function addCnEnSpacingToText(text) {
    if (!text || text.length < 2) return text;

    // CJK 后跟拉丁/数字 → 加空格
    // 使用 Unicode 属性：CJK 统一汉字范围
    text = text.replace(/([一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ])([a-zA-Z0-9])/g, '$1 $2');

    // 拉丁/数字后跟 CJK → 加空格
    text = text.replace(/([a-zA-Z0-9])([一-鿿㐀-䶿⺀-⻿　-〿＀-￯㇀-㇯⼀-⿟㄀-ㄯㆠ-ㆿ])/g, '$1 $2');

    // 修正引号/括号内侧的多余空格（中文引号 "「『"』」""''  括号 （） 内侧不空格）
    var pairs = [
      ['“', '”'], // ""
      ['‘', '’'], // ''
      ['「', '」'], // 「」
      ['『', '』'], // 『』
      ['（', '）'], // （）
      ['《', '》'], // 《》
      ['“', '”'], // "" (U+201C U+201D)
      ['‘', '’']  // '' (U+2018 U+2019)
    ];
    for (var i = 0; i < pairs.length; i++) {
      var left = pairs[i][0];
      var right = pairs[i][1];
      // 左引号/括号后紧跟英文 → 去掉多余空格
      var reLeft = new RegExp('(' + left + ') ([a-zA-Z])', 'g');
      text = text.replace(reLeft, '$1$2');
      // 英文后紧跟右引号/括号 → 去掉多余空格
      var reRight = new RegExp('([a-zA-Z0-9]) (' + right + ')', 'g');
      text = text.replace(reRight, '$1$2');
    }

    return text;
  }

  /**
   * 处理中英文间距：逐段落获取文本 → 计算正确间距 → 替换
   *
   * 注意：使用 range.insertText 会丢失段落内的字符级格式（加粗/斜体等），
   * 因此在 applyBodyFormat 之后执行，后续会重新应用段落级格式。
   *
   * @param {Word.RequestContext} context
   * @param {Word.ParagraphCollection} paragraphs
   * @param {FormatStats} stats
   * @param {object} opts
   */
  function processCnEnSpacing(context, paragraphs, stats, opts) {
    var items = paragraphs.items;

    for (var i = 0; i < items.length; i++) {
      var p = items[i];
      var originalText = p.text || '';

      if (originalText.trim().length === 0) continue;

      var corrected = addCnEnSpacingToText(originalText);

      if (corrected !== originalText) {
        stats.cjkSpacing++;
        try {
          // 使用 insertText + Replace 替换整个段落内容
          // 这会丢失字符级格式，但保持段落样式
          p.insertText(corrected, Word.InsertLocation.replace);
        } catch (e) {
          console.warn('processCnEnSpacing: paragraph ' + i + ' insertText failed', e);
        }
      }
    }
  }

  /* ═══════════════════════════════════════════════════════════
     原生排版模块 — 清除连续空行
     ═══════════════════════════════════════════════════════════ */

  /**
   * 删除连续空段落，保留最多 1 个空行
   * @param {Word.RequestContext} context
   * @param {Word.ParagraphCollection} paragraphs
   * @param {FormatStats} stats
   */
  function removeEmptyLines(context, paragraphs, stats) {
    var items = paragraphs.items;
    var consecutiveEmpty = 0;

    // 从后往前遍历，以便安全删除
    for (var i = items.length - 1; i >= 0; i--) {
      var text = (items[i].text || '').trim();

      if (text.length === 0) {
        consecutiveEmpty++;
        // 第一个空行保留（即最后一个连续空行），其余删除
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
     原生排版模块 — 主编排器（支持分批处理）
     ═══════════════════════════════════════════════════════════ */

  var BATCH_SIZE = 80;  // 每批处理的段落数

  /**
   * 一键全文排版：加载文档 → 检查保护 → 分批执行排版管道
   *
   * 管道步骤：
   *   1. 标题检测 → apply style
   *   2. 正文格式 → font/size/spacing/indent
   *   3. 中英文间距 → CN-EN spacing
   *   4. 清除空行 → remove consecutive empties
   *
   * 超过 BATCH_SIZE 时自动分批，每批独立 Word.run 确保事务完整性。
   */
  function autoFormatDocument() {
    clearStatus(el.formatStatus);
    clearStatus(el.mainStatus);

    // 解析排版参数
    var opts = getFormatOptions();

    // 禁用按钮，显示进度
    el.autoFormatBtn.disabled = true;
    setSpinner(el.autoFormatSpinner, true);
    showStatus(el.formatStatus, 'info', '正在加载文档...');

    // 第一阶段：加载段落并检查保护
    Word.run(function (context) {
      var body = context.document.body;
      var paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        var total = paragraphs.items.length;

        if (total === 0) {
          showStatus(el.formatStatus, 'info', '文档无内容，无需排版。');
          return;
        }

        // 检查文档保护
        return checkDocumentProtection(context).then(function (protResult) {
          if (protResult.protected) {
            showStatus(el.formatStatus, 'error', protResult.reason);
            return;
          }

          // 初始化
          var stats = createStats();
          stats.total = total;

          // 分批处理
          var batchCount = Math.ceil(total / BATCH_SIZE);

          showStatus(el.formatStatus, 'info', '文档共 ' + total + ' 段，分 ' + batchCount + ' 批处理...');
          updateFormatProgress(stats, 0, total);

          // 递归处理每一批
          return processBatch(0, batchCount, total, stats, opts);
        });
      });
    }).catch(function (err) {
      console.error('autoFormatDocument error:', err);
      showStatus(el.formatStatus, 'error', '排版失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      el.autoFormatBtn.disabled = false;
      setSpinner(el.autoFormatSpinner, false);
    });
  }

  /**
   * 递归分批处理入口
   * @param {number} batchIndex - 当前批次索引 (0-based)
   * @param {number} batchCount - 总批次数
   * @param {number} total - 总段落数
   * @param {FormatStats} stats - 排版统计
   * @param {object} opts - 排版参数
   * @returns {Promise}
   */
  function processBatch(batchIndex, batchCount, total, stats, opts) {
    if (batchIndex >= batchCount) {
      // 全部完成
      return finalizeFormat(stats);
    }

    var start = batchIndex * BATCH_SIZE;
    var end = Math.min(start + BATCH_SIZE, total);

    return Word.run(function (context) {
      var body = context.document.body;
      var paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        // 获取当前批次的段落切片
        var allItems = paragraphs.items;
        var batchParagraphs = {
          items: allItems.slice(start, end)
        };

        // 步骤 1: 标题检测
        var headingIndices = detectAndStyleHeadings(context, batchParagraphs, stats, opts);

        // 步骤 2: 正文格式
        applyBodyFormat(context, batchParagraphs, stats, opts, headingIndices);

        // 步骤 3: 中英文间距
        if (opts.enableCnEnSpacing) {
          processCnEnSpacing(context, batchParagraphs, stats, opts);
        }

        // 步骤 4: 清除连续空行
        if (opts.removeEmptyLines) {
          removeEmptyLines(context, batchParagraphs, stats);
        }

        return context.sync();
      });
    }).then(function () {
      // 更新进度
      updateFormatProgress(stats, end, total);

      // 处理下一批
      return processBatch(batchIndex + 1, batchCount, total, stats, opts);
    }).catch(function (err) {
      console.error('Batch ' + batchIndex + ' error:', err);
      // 单批失败不终止，继续处理下一批
      showStatus(el.formatStatus, 'warn',
        '第 ' + (batchIndex + 1) + ' 批处理出错: ' + (err.message || '未知') + '，继续处理后续...');
      return processBatch(batchIndex + 1, batchCount, total, stats, opts);
    });
  }

  /**
   * 排版完成后收尾：隐藏进度条，显示统计 Toast
   * @param {FormatStats} stats
   */
  function finalizeFormat(stats) {
    // 隐藏进度条
    setTimeout(function () {
      el.formatProgress.classList.remove('active');
      el.formatProgressBar.style.width = '0%';
    }, 600);

    // 生成统计消息
    var parts = [];
    if (stats.headings > 0) parts.push('标题 ' + stats.headings + ' 处');
    if (stats.body > 0) parts.push('正文 ' + stats.body + ' 段');
    if (stats.cjkSpacing > 0) parts.push('中英文间距 ' + stats.cjkSpacing + ' 处');
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

    // --- 一键全文排版（新增） ---
    el.autoFormatBtn.addEventListener('click', function () {
      autoFormatDocument();
    });

    // --- 保存排版参数（新增） ---
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
      console.log('OfficeAI v1.4: Word host detected');

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

      console.log('OfficeAI v1.4: Initialization complete');
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
    }
  });

})();
