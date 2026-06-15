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
     排版预设定义
     ═══════════════════════════════════════════════════════════ */

  /**
   * 原生格式预设（Office.js 直接操作，不经过 AI）
   * 映射到 data-preset 属性值
   */
  var NATIVE_PRESETS = ['format_title', 'format_indent', 'format_spacing', 'format_font'];

  /**
   * 预设标签名映射
   */
  var PRESET_LABELS = {
    format_title:      '标题加粗居中',
    format_indent:     '正文首行缩进',
    format_spacing:    '段落间距调整',
    format_font:       '正文字号行距',
    polish:            '校对润色',
    translate_cn2en:   '中译英',
    translate_en2cn:   '英译中',
    summarize:         '生成摘要'
  };

  /**
   * AI 预设（需要 AI 重写文本内容）
   */
  var AI_PRESET_PROMPTS = {
    polish: {
      label: '校对润色',
      systemAddon: '你是一位专业文字校对与润色专家。请修正语法错误、错别字，优化遣词造句，提升流畅度和逻辑性。\n' +
                   '⚠ 重要：保持原文的段落结构和层级不变，每个段落独立处理，段落之间用空行分隔。不要改变原意。',
      prompt: '请校对并润色以下文本，修正错别字和语法问题。保持段落结构不变：'
    },
    translate_cn2en: {
      label: '中译英',
      systemAddon: '你是一位专业中英翻译专家。请准确、地道地将中文翻译为英文。\n' +
                   '⚠ 重要：保持原文的段落结构，每个段落独立翻译，段落之间用空行分隔。',
      prompt: '请将以下中文翻译为英文，保持段落结构：'
    },
    translate_en2cn: {
      label: '英译中',
      systemAddon: '你是一位专业英中翻译专家。请准确、流畅地将英文翻译为中文。\n' +
                   '⚠ 重要：保持原文的段落结构，每个段落独立翻译，段落之间用空行分隔。',
      prompt: '请将以下英文翻译为中文，保持段落结构：'
    },
    summarize: {
      label: '生成摘要',
      systemAddon: '你是一位专业文档摘要专家。请生成简洁但全面的摘要，保留核心观点和关键数据。使用要点列表格式输出。',
      prompt: '请为以下文本生成清晰、条理分明的摘要（使用要点列表）：'
    }
  };

  /* ═══════════════════════════════════════════════════════════
     原生格式预设 — 标题检测
     ═══════════════════════════════════════════════════════════ */

  /**
   * 使用正则匹配段落文本的标题级别
   * @param {string} text - 段落文本（已 trim）
   * @returns {number} 1/2/3 = 标题级别，0 = 非标题（正文）
   */
  function detectHeadingLevel(text) {
    var clean = text.replace(/^\s+/, '').slice(0, 80);

    // Level 1: 第X章/摘要/绪论/引言/前言/结论/参考文献/致谢/附录
    if (/^(摘要|绪论|引言|前言|结论|参考文献|致谢|附录)[\s:：]*$/.test(clean)) return 1;
    if (/^第[一二三四五六七八九十百千\d]+[章篇部]/.test(clean)) return 1;

    // Level 2: 第X节 / 一、二、三 / 1. 2. 数字编号
    if (/^第[一二三四五六七八九十百千\d]+节/.test(clean)) return 2;
    if (/^[一二三四五六七八九十]+[、．.\s]/.test(clean)) return 2;
    if (/^\d+[\.\、\s)]/.test(clean)) return 2;

    // Level 3: 1.1 / (一) / 小标题
    if (/^\d+\.\d+/.test(clean)) return 3;
    if (/^（[一二三四五六七八九十\d]+）/.test(clean)) return 3;

    return 0;
  }

  /* ═══════════════════════════════════════════════════════════
     原生格式预设 — Office.js 执行引擎
     ═══════════════════════════════════════════════════════════ */

  /**
   * 原生格式预设执行器（不经过 AI，直接操作文档）
   *
   * 流程：
   *   1. Word.run → 获取选区段落
   *   2. 选区为空则回退到全文
   *   3. 按 presetKey 用 Office.js API 逐段应用格式
   *   4. context.sync → 显示结果
   *
   * @param {string} presetKey - 预设键名（format_title|format_indent|format_spacing|format_font）
   */
  function executeNativeFormatPreset(presetKey) {
    clearStatus(el.mainStatus);
    var label = PRESET_LABELS[presetKey] || presetKey;
    showStatus(el.mainStatus, 'info', '正在应用「' + label + '」...');

    var opts = getFormatOptions();

    Word.run(function (context) {
      // 获取选区段落集合（与 smartFormatDocument 使用相同的加载模式）
      var selParagraphs = context.document.getSelection().paragraphs;
      context.load(selParagraphs, 'items');

      return context.sync().then(function () {
        var paragraphs = selParagraphs.items;

        // 选区无有效内容 → 回退到全文
        var hasContent = false;
        if (paragraphs && paragraphs.length > 0) {
          for (var i = 0; i < paragraphs.length; i++) {
            if ((paragraphs[i].text || '').trim().length > 0) { hasContent = true; break; }
          }
        }

        if (!hasContent) {
          // 回退到全文（复用已验证的加载模式）
          var bodyParagraphs = context.document.body.paragraphs;
          context.load(bodyParagraphs, 'items');
          return context.sync().then(function () {
            applyFormatToParagraphs(bodyParagraphs.items, presetKey, opts);
            return context.sync();
          });
        }

        applyFormatToParagraphs(paragraphs, presetKey, opts);
        return context.sync();
      });
    }).then(function () {
      showStatus(el.mainStatus, 'success', '✓ 「' + label + '」已应用。');
    }).catch(function (err) {
      console.error('executeNativeFormatPreset error:', err);
      showStatus(el.mainStatus, 'error', '格式应用失败: ' + (err.message || '未知错误'));
    });
  }

  /**
   * 对段落数组应用指定的原生格式
   *
   * 格式映射：
   *   format_title    → 标题: Heading 样式 + 居中 + 加粗
   *   format_indent   → 正文: 首行缩进 2 字符；标题不变
   *   format_spacing  → 标题: 段前12pt 段后6pt；正文: 段后6pt
   *   format_font     → 正文: 12pt/1.5倍行距；标题: H1=18pt H2=16pt H3=14pt
   *
   * @param {Array<Word.Paragraph>} paragraphs - Office.js 段落对象数组
   * @param {string} presetKey - 格式预设键名
   * @param {object} opts - 排版参数配置
   */
  function applyFormatToParagraphs(paragraphs, presetKey, opts) {
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var text = (p.text || '').trim();
      if (text.length === 0) continue;

      var hl = detectHeadingLevel(text); // 0=正文, 1/2/3=标题级别

      switch (presetKey) {
        case 'format_title':
          // ★ 标题: 应用 Word 内置 Heading 样式（不是 <b> 标签！）
          if (hl > 0) {
            try { p.style = 'Heading ' + hl; } catch (e) {}
            try { p.alignment = 'Centered'; } catch (e) {}
            try { p.font.bold = true; } catch (e) {}
          }
          // 正文段落保持原样
          break;

        case 'format_indent':
          // 正文: 首行缩进 2 字符（约 24pt @ 12pt 字号）
          if (hl === 0) {
            try {
              p.paragraphFormat.firstLineIndent = opts.fontSize * opts.indentChars;
            } catch (e) {}
          }
          break;

        case 'format_spacing':
          // 标题: 段前距 12pt, 段后距 6pt
          // 正文: 段后距 6pt
          try {
            p.paragraphFormat.spaceBefore = (hl >= 1) ? 12 : 0;
            p.paragraphFormat.spaceAfter = 6;
          } catch (e) {}
          break;

        case 'format_font':
          // 正文: 12pt 字号, 1.5 倍行距
          // 标题: 按级别递增 H1=18pt, H2=16pt, H3=14pt
          try {
            p.font.name = opts.cnFont;
            if (hl >= 1) {
              p.font.size = opts.fontSize + (4 - hl) * 2; // H1=18, H2=16, H3=14
            } else {
              p.font.size = opts.fontSize; // 正文 12pt
              p.paragraphFormat.lineSpacing = opts.fontSize * opts.lineSpacing * 1.2;
            }
          } catch (e) {}
          break;
      }
    }
  }

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
   * AI 语义分析 System Prompt（精简版）
   *
   * 设计原则：
   *   - 最小化 token 消耗（~350 tokens vs 旧版 ~600 tokens）
   *   - 6 种操作覆盖全部排版场景
   *   - 明确禁止 markdown 包裹，强制纯 JSON 输出
   *   - 列表项/标题不与 setBody 重复标记
   */
  var SMART_FORMAT_SYSTEM_PROMPT =
    '你是文档结构分析引擎。分析段落数组，输出 JSON 排版指令。\n\n' +
    '## 指令类型\n' +
    '- setHeading: 标题。level:1(第X章/摘要/绪论/结论/参考文献/致谢) 2(第X节/一、/1.) 3(1.1/(一)/小节)\n' +
    '- setBody: 普通正文段落\n' +
    '- addSpaceBetweenCnEn: 同时包含中文和英文/数字（需加空格，可独立或与 setBody 合并）\n' +
    '- setCodeBlock: 代码块（等宽字体/缩进/含关键词如 function var class def import）\n' +
    '- setList: 列表项。type:"bullet"|"number"\n' +
    '- setQuote: 引用/注释性文字\n\n' +
    '## 规则\n' +
    '1. 每个非空段落恰好 1 条指令，空段落跳过\n' +
    '2. 标题不标记为 setBody；列表项不标记为 setBody\n' +
    '3. 中英文/数字混排优先标记 addSpaceBetweenCnEn\n' +
    '4. 纯输出 JSON 数组，不加解释和 markdown 包裹';

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
   * Promise 超时包装
   * @param {Promise} promise - 原始 promise
   * @param {number} ms - 超时毫秒
   * @returns {Promise}
   */
  function withTimeout(promise, ms) {
    return Promise.race([
      promise,
      new Promise(function (_, reject) {
        setTimeout(function () {
          reject(new Error('请求超时（' + (ms / 1000).toFixed(0) + ' 秒）'));
        }, ms);
      })
    ]);
  }

  /**
   * 单批 AI 分析调用
   * @param {Array<{index: number, text: string}>} paragraphTexts - 段落文本（已截断 120 字符）
   * @param {object} cfg - 当前配置
   * @returns {Promise<Array|null>} 指令数组
   */
  function callAIForBatch(paragraphTexts, cfg) {
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

    return withTimeout(fetchPromise, AI_TIMEOUT_MS).then(function (rawText) {
      console.log('AI batch response length:', rawText.length);
      var instructions = parseAIResponse(rawText);
      if (!instructions || !Array.isArray(instructions) || instructions.length === 0) {
        console.warn('AI batch returned unparseable response');
        return null;
      }
      return instructions;
    }).catch(function (err) {
      console.error('AI batch failed:', err);
      return null; // 单批失败不终止，由上层合并处理
    });
  }

  /**
   * 合并多个批次的指令数组，后批覆盖前批同索引冲突
   * @param {Array<Array<object>>} instructionArrays - 多批指令数组
   * @returns {Array<object>} 合并后按 index 排序的指令数组
   */
  function mergeInstructions(instructionArrays) {
    var merged = {};
    for (var i = 0; i < instructionArrays.length; i++) {
      var batch = instructionArrays[i];
      if (!batch || !Array.isArray(batch)) continue;
      for (var j = 0; j < batch.length; j++) {
        var instr = batch[j];
        if (instr && typeof instr.index === 'number' && instr.action) {
          merged[instr.index] = instr; // 后批覆盖前批
        }
      }
    }
    var result = [];
    var keys = Object.keys(merged);
    for (var k = 0; k < keys.length; k++) {
      result.push(merged[keys[k]]);
    }
    result.sort(function (a, b) { return a.index - b.index; });
    return result;
  }

  /**
   * 调用 AI 进行文档结构分析（支持大文档分批）
   *
   * 策略：
   *   - ≤ AI_BATCH_SIZE 段：单次调用
   *   - > AI_BATCH_SIZE 段：顺序分批发送，每批返回的指令按 index 合并
   *
   * @param {Array<{index: number, text: string}>} paragraphTexts - 已截断的段落文本数组
   * @returns {Promise<Array|null>} 合并后的指令数组，全部失败返回 null
   */
  function callAIForAnalysis(paragraphTexts) {
    var cfg = getConfigFromUI();

    if (cfg.provider !== 'ollama' && !cfg.apiKey) {
      showStatus(el.formatStatus, 'error', '请先在设置中配置 API Key。');
      return Promise.resolve(null);
    }

    // 小文档：单次调用
    if (paragraphTexts.length <= AI_BATCH_SIZE) {
      return callAIForBatch(paragraphTexts, cfg);
    }

    // 大文档：分批发送，顺序执行（避免并发限流）
    var allResults = [];
    var batchCount = Math.ceil(paragraphTexts.length / AI_BATCH_SIZE);
    var failCount = 0;

    showStatus(el.formatStatus, 'info',
      'AI 分 ' + batchCount + ' 批分析中 (1/' + batchCount + ')...');

    /**
     * 递归处理单批
     * @param {number} idx - 当前批次索引（0-based）
     * @returns {Promise<Array|null>}
     */
    function processBatch(idx) {
      if (idx >= batchCount) {
        // 全部完成
        if (allResults.length === 0) {
          showStatus(el.formatStatus, 'warn', '所有 AI 批次均失败，将使用规则引擎回退。');
          return Promise.resolve(null);
        }
        var merged = mergeInstructions(allResults);
        showStatus(el.formatStatus, 'info',
          'AI 分析完成：' + batchCount + ' 批共生成 ' + merged.length + ' 条指令');
        return Promise.resolve(merged);
      }

      var start = idx * AI_BATCH_SIZE;
      var end = Math.min(start + AI_BATCH_SIZE, paragraphTexts.length);
      var batch = paragraphTexts.slice(start, end);

      return callAIForBatch(batch, cfg).then(function (instructions) {
        if (instructions && instructions.length > 0) {
          allResults.push(instructions);
        } else {
          failCount++;
        }
        if (idx + 1 < batchCount) {
          showStatus(el.formatStatus, 'info',
            'AI 分 ' + batchCount + ' 批分析中 (' + (idx + 2) + '/' + batchCount + ')...');
        }
        return processBatch(idx + 1);
      });
    }

    return processBatch(0);
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

  var BATCH_SIZE      = 80;    // 进度条更新单位（段/步）
  var AI_BATCH_SIZE   = 150;   // AI 分析每批段落数
  var AI_TIMEOUT_MS   = 30000; // AI 调用超时（毫秒）

  /**
   * 一键智能排版 — 主编排器（v1.6 混合架构）
   *
   * 三阶段流水线：
   *   Phase 0: Word.run 加载段落 → 检查保护 → 提取文本数组
   *   Phase 1: AI 语义分析（自动分批）→ JSON 指令数组
   *   Phase 2: 单 Word.run 事务执行全部指令 + 清除空行 → 一次 sync
   *
   * 关键改进（v1.5 → v1.6）：
   *   - Phase 2 合并为单个 Word.run + 单次 context.sync → Ctrl+Z 一键撤销全部排版
   *   - AI 分析支持 >150 段自动分批，不再截断大文档
   *   - 每批 AI 调用含 30s 超时保护
   */
  function smartFormatDocument() {
    clearStatus(el.formatStatus);
    clearStatus(el.mainStatus);

    var opts = getFormatOptions();

    el.autoFormatBtn.disabled = true;
    setSpinner(el.autoFormatSpinner, true);
    showStatus(el.formatStatus, 'info', '正在加载文档...');

    /* ═══ Phase 0: 加载段落 + 提取文本 ═══ */
    Word.run(function (context) {
      var body = context.document.body;
      var paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        var items = paragraphs.items;
        var total = items.length;

        if (total === 0) {
          showStatus(el.formatStatus, 'info', '文档无内容，无需排版。');
          return null;
        }

        return checkDocumentProtection(context).then(function (protResult) {
          if (protResult.protected) {
            showStatus(el.formatStatus, 'error', protResult.reason);
            return null;
          }

          // 提取段落文本：保留原始 index，截断 120 字符控制 token 消耗
          var paragraphTexts = [];
          for (var i = 0; i < total; i++) {
            var t = (items[i].text || '').trim();
            if (t.length > 0) {
              paragraphTexts.push({ index: i, text: t.slice(0, 120) });
            }
          }

          return { total: total, paragraphTexts: paragraphTexts };
        });
      });
    }).then(function (paraData) {
      if (!paraData) return; // 空文档或受保护文档

      var total = paraData.total;
      var paragraphTexts = paraData.paragraphTexts;

      var stats = createStats();
      stats.total = total;

      updateFormatProgress(stats, 0, total, 'AI 分析中...');
      showStatus(el.formatStatus, 'info', '文档共 ' + total + ' 段，AI 正在分析结构...');

      /* ═══ Phase 1: AI 语义分析（自动分批） ═══ */
      return callAIForAnalysis(paragraphTexts).then(function (instructions) {
        if (!instructions || instructions.length === 0) {
          showStatus(el.formatStatus, 'warn',
            'AI 分析失败，使用规则引擎回退。将基于正则匹配执行基本排版...');
          instructions = generateFallbackInstructions(paragraphTexts);
        } else {
          showStatus(el.formatStatus, 'info',
            'AI 已生成 ' + instructions.length + ' 条指令，单事务原生执行中...');
        }

        updateFormatProgress(stats, 0, total, '原生执行中...');

        /* ═══ Phase 2: 单 Word.run 事务 — 全部指令 + 清除空行 → 一次 sync ═══ */
        return Word.run(function (context) {
          var body = context.document.body;
          var paragraphs = body.paragraphs;
          context.load(paragraphs, 'items');

          return context.sync().then(function () {
            var items = paragraphs.items;

            // 复用指令执行引擎（全量处理：batchStart=0, batchEnd=items.length）
            executeInstructions(context, items, instructions, 0, items.length, stats, opts);

            // 清除连续空行
            if (opts.removeEmptyLines) {
              removeEmptyLines(context, paragraphs, stats);
            }

            updateFormatProgress(stats, total, total, '同步中...');

            // ★ 单次 sync：整个排版 = 一个撤销单元
            return context.sync();
          });
        }).then(function () {
          finalizeSmartFormat(stats, opts);
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
   * AI 失败回退方案：基于正则规则从文本生成排版指令
   *
   * 新增识别能力（v1.6）：
   *   - 代码块：等宽字体特征、缩进 ≥4 空格、含代码关键词
   *   - 列表：项目符号（• - *）或编号（1. 2.）
   *   - 引用：以 ">" 开头或引号包裹的短段落
   *   - 中英文混排：自动标记 addSpaceBetweenCnEn
   *
   * @param {Array<{index: number, text: string}>} paragraphTexts - 文本数组
   * @returns {Array<object>} 排版指令数组
   */
  function generateFallbackInstructions(paragraphTexts) {
    var instructions = [];

    // 标题正则（优先级从高到低匹配）
    var HEADING_RULES = [
      { regex: /^(摘要|绪论|引言|前言|结论|参考文献|致谢|附录)[\s:：]*$/,  level: 1 },
      { regex: /^第[一二三四五六七八九十百千\d]+[章篇部]/,              level: 1 },
      { regex: /^第[一二三四五六七八九十百千\d]+节/,                    level: 2 },
      { regex: /^[一二三四五六七八九十]+[、．.\s]/,                     level: 2 },
      { regex: /^\d+[\.\、\s)]/,                                         level: 2 },
      { regex: /^\d+\.\d+[\s\.\、)]/,                                    level: 3 },
      { regex: /^（[一二三四五六七八九十\d]+）/,                        level: 3 }
    ];

    // 代码块特征：等宽字体常见关键词 + 缩进特征
    var CODE_KEYWORDS = /\b(function|var |let |const |class |def |import |from |return |if |for |while |public |private |static |void |int |string |console\.|System\.|#include|#define)\b/;
    var CODE_INDENT_RE = /^(    |\t)/; // 4空格缩进或 Tab

    // 列表特征
    var BULLET_RE = /^[•\-\*▪▸●○]\s/;
    var NUMBER_RE = /^\d+[\.\)、]\s/;

    // 引用特征
    var QUOTE_RE = /^[>」」"'"『"].*[」"'"』"]$/;

    for (var i = 0; i < paragraphTexts.length; i++) {
      var item = paragraphTexts[i];
      var text = item.text;
      if (text.length === 0) continue;

      var cleanText = text.replace(/^\s+/, '').slice(0, 80);

      // 1. 标题检测
      var isHeading = false;
      for (var j = 0; j < HEADING_RULES.length; j++) {
        if (HEADING_RULES[j].regex.test(cleanText)) {
          instructions.push({ index: item.index, action: 'setHeading', level: HEADING_RULES[j].level });
          isHeading = true;
          break;
        }
      }
      if (isHeading) continue;

      // 2. 代码块检测
      if (CODE_INDENT_RE.test(text) || CODE_KEYWORDS.test(text.slice(0, 120))) {
        instructions.push({ index: item.index, action: 'setCodeBlock' });
        continue;
      }

      // 3. 列表检测
      if (BULLET_RE.test(cleanText)) {
        instructions.push({ index: item.index, action: 'setList', type: 'bullet' });
        continue;
      }
      if (NUMBER_RE.test(cleanText) && cleanText.length < 60) {
        instructions.push({ index: item.index, action: 'setList', type: 'number' });
        continue;
      }

      // 4. 引用检测
      if (QUOTE_RE.test(cleanText) && cleanText.length < 100) {
        instructions.push({ index: item.index, action: 'setQuote' });
        continue;
      }

      // 5. 中英文混排检测
      var hasCJK = CJK_RE.test(text);
      var hasLatin = /[a-zA-Z0-9]/.test(text);
      if (hasCJK && hasLatin) {
        instructions.push({ index: item.index, action: 'addSpaceBetweenCnEn' });
      }

      // 6. 默认正文
      instructions.push({ index: item.index, action: 'setBody' });
    }

    return instructions;
  }

  /**
   * 排版完成后的统计与反馈
   *
   * 展示 Toast 含：标题/正文/中英文间距/代码块/列表/引用/空行的修改计数。
   * 600ms 后自动隐藏进度条。
   *
   * @param {object} stats - 修改统计 {headings, body, cjkSpacing, codeBlock, list, quote, empties, total}
   * @param {object} opts - 排版参数
   */
  function finalizeSmartFormat(stats, opts) {
    setTimeout(function () {
      el.formatProgress.classList.remove('active');
      el.formatProgressBar.style.width = '0%';
    }, 600);

    var parts = [];
    if (stats.headings > 0)   parts.push('标题 ' + stats.headings + ' 处');
    if (stats.body > 0)       parts.push('正文 ' + stats.body + ' 段');
    if (stats.cjkSpacing > 0) parts.push('中英文间距 ' + stats.cjkSpacing + ' 处');
    if (stats.codeBlock > 0)  parts.push('代码块 ' + stats.codeBlock + ' 处');
    if (stats.list > 0)       parts.push('列表 ' + stats.list + ' 项');
    if (stats.quote > 0)      parts.push('引用 ' + stats.quote + ' 处');
    if (stats.empties > 0)    parts.push('清除空行 ' + stats.empties + ' 行');

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

    // --- 排版预设按钮（原生格式 + AI 分流） ---
    el.presetBtns.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var presetKey = btn.getAttribute('data-preset');
        if (!presetKey) return;

        // ★ 原生格式预设：直接操作文档，不经过 AI
        if (NATIVE_PRESETS.indexOf(presetKey) >= 0) {
          executeNativeFormatPreset(presetKey);
          return;
        }

        // AI 预设：需要 AI 重写文本内容
        var preset = AI_PRESET_PROMPTS[presetKey];
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
      console.log('OfficeAI v1.7: Word host detected');

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

      console.log('OfficeAI v1.7: Initialization complete');
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
    }
  });

})();
