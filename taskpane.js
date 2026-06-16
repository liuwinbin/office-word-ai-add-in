/* global Office, Word, console */

(function () {
  'use strict';

  /* ═══════════════════════════════════════════════════════════
     DOM 引用
     ═══════════════════════════════════════════════════════════ */
  var el = {};

  /* ═══════════════════════════════════════════════════════════
     会话状态（IIFE 内部，不挂载 window）
     ═══════════════════════════════════════════════════════════ */

  /** @type {Array<{role: string, content: string}>} 对话历史（最多 10 轮 = 20 条消息） */
  var conversationHistory = [];
  var MAX_HISTORY_ROUNDS = 10;

  /** @type {string|null} 当前光标所在段落的文本（截断 500 字） */
  var currentCursorContext = null;

  /** @type {string} 最后一条 AI 回复的完整内容（供替换/追加使用） */
  var _lastAIResponse = '';

  /** 文档全文缓存 */
  var _docTextCache = { text: '', timestamp: 0, ttl: 30000 };

  /** SelectionChanged 防抖计时器 */
  var _selectionDebounceTimer = null;
  var SELECTION_DEBOUNCE_MS = 300;

  /** 上次已知文档 URL（用于检测文档切换） */
  var _lastDocumentUrl = null;

  function cacheDom() {
    // 主界面
    el.mainPage         = document.getElementById('mainPage');
    el.settingsPage     = document.getElementById('settingsPage');
    el.presetsPage      = document.getElementById('presetsPage');
    el.openSettingsBtn  = document.getElementById('openSettingsBtn');
    el.openPresetsBtn   = document.getElementById('openPresetsBtn');
    el.backToMainBtn    = document.getElementById('backToMainBtn');
    el.backFromPresetsBtn = document.getElementById('backFromPresetsBtn');
    el.readFullDocBtn   = document.getElementById('readFullDocBtn');
    el.extractSelBtn    = document.getElementById('extractSelectionBtn');
    el.readSpinner      = document.getElementById('readSpinner');
    el.documentText     = document.getElementById('documentText');
    el.mainStatus       = document.getElementById('mainStatus');

    // 聊天 UI
    el.contextBar       = document.getElementById('contextBar');
    el.contextText      = document.getElementById('contextText');
    el.clearContextBtn  = document.getElementById('clearContextBtn');
    el.chatMessages     = document.getElementById('chatMessages');
    el.chatEmptyHint    = document.getElementById('chatEmptyHint');
    el.chatTyping       = document.getElementById('chatTyping');
    el.chatInput        = document.getElementById('chatInput');
    el.sendBtn          = document.getElementById('sendBtn');
    el.clearChatBtn     = document.getElementById('clearChatBtn');
    el.insertLastBtn    = document.getElementById('insertLastBtn');
    el.chatSpinner      = document.getElementById('chatSpinner');
    el.chatStatus       = document.getElementById('chatStatus');
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
    el.presetsStatus     = document.getElementById('presetsStatus');

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

  /**
   * 获取文档全文（带 30s 缓存，避免大文档重复抓取）
   *
   * 缓存策略：
   *   - 命中缓存（30s 内）→ 直接返回 Promise.resolve(cached)
   *   - 缓存过期/不存在 → Word.run 抓取 → 更新缓存
   *
   * @returns {Promise<string>} 文档全文
   */
  function getDocumentText() {
    var now = Date.now();
    if (_docTextCache.text && (now - _docTextCache.timestamp) < _docTextCache.ttl) {
      return Promise.resolve(_docTextCache.text);
    }

    return Word.run(function (context) {
      var body = context.document.body;
      context.load(body, 'text');
      return context.sync().then(function () {
        _docTextCache.text = body.text || '';
        _docTextCache.timestamp = Date.now();
        return _docTextCache.text;
      });
    }).catch(function (err) {
      console.warn('getDocumentText failed:', err);
      // 返回过期缓存（如果有）
      return _docTextCache.text || '';
    });
  }

  /**
   * 获取当前选区文本（带段落样式标注）
   *
   * @returns {Promise<{plainText: string, styledText: string}>}
   */
  function getSelectedTextWithStyles() {
    return Word.run(function (context) {
      var selection = context.document.getSelection();
      var paragraphs = selection.paragraphs;
      context.load(paragraphs, 'items');
      return context.sync().then(function () {
        var items = paragraphs.items;

        // 统计非空段落数：≤1 个 → 光标定位，不加格式标注，沿用旧纯文本行为
        var nonEmpty = 0, onlyText = '';
        for (var i = 0; i < items.length; i++) {
          var t = (items[i].text || '').trim();
          if (t.length > 0) { nonEmpty++; onlyText = t; }
        }

        if (nonEmpty <= 1) {
          return { plainText: onlyText, styledText: '' };
        }

        // ≥2 个段落被选中 → 带格式标注
        var plainLines = [];
        var styledLines = [];
        for (var j = 0; j < items.length; j++) {
          var p = items[j];
          var text = (p.text || '').trim();
          if (text.length === 0) continue;

          var styleName = '';
          try { styleName = p.style || ''; } catch (e) {}

          var styleTag = styleName;
          if (/^Heading 1$|^标题 1$/i.test(styleName)) styleTag = '标题 1';
          else if (/^Heading 2$|^标题 2$/i.test(styleName)) styleTag = '标题 2';
          else if (/^Heading 3$|^标题 3$/i.test(styleName)) styleTag = '标题 3';
          else if (/^Heading 4$|^标题 4$/i.test(styleName)) styleTag = '标题 4';
          else if (/^Heading 5$|^标题 5$/i.test(styleName)) styleTag = '标题 5';
          else if (/^Caption$|^题注$/i.test(styleName)) styleTag = '题注';
          else if (/^Normal$|^正文$/i.test(styleName)) styleTag = '正文';

          plainLines.push(text);
          styledLines.push('[' + styleTag + '] ' + text);
        }
        return {
          plainText: plainLines.join('\n'),
          styledText: styledLines.join('\n')
        };
      });
    }).catch(function (err) {
      console.warn('getSelectedTextWithStyles failed:', err);
      return { plainText: '', styledText: '' };
    });
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
    el.mainPage.classList.remove('active');
    el.settingsPage.classList.remove('active');
    el.presetsPage.classList.remove('active');

    if (page === 'main') {
      el.mainPage.classList.add('active');
    } else if (page === 'settings') {
      el.settingsPage.classList.add('active');
    } else if (page === 'presets') {
      el.presetsPage.classList.add('active');
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
     样式修改模块 — System Prompt 附加段
     ═══════════════════════════════════════════════════════════ */

  /**
   * 样式修改识别 System Prompt 附加段
   *
   * 当用户输入是样式修改指令时，指示 AI 输出纯 JSON。
   * 仅在非 preset 模式下附加到 system prompt 前面。
   */
  var STYLE_MODIFY_SYSTEM_ADDON =
    '## 文档样式实时修改能力\n' +
    '你具备实时修改 Word 文档样式的能力。**绝不输出 HTML 代码**——HTML 只显示在对话框里，无法修改 Word 文档。\n' +
    '样式修改必须输出纯 JSON，系统会自动执行并应用到文档中。\n\n' +
    '### 需要输出 JSON 的场景（全部）：\n' +
    '- 直接指令："把标题3改成宋体小四加粗"\n' +
    '- 行距指令："正文行距改成固定值22磅"\n' +
    '- 段间距指令："标题1段前12磅段后6磅"\n' +
    '- ★ 格式规范列表（批量）：用户贴一段格式要求（如论文排版规范），逐条转换为 JSON\n' +
    '  例：用户说"各章题序及标题 小2号黑体，上下各空一行"\n' +
    '  → {"action":"modify_style","targetStyle":"Heading 1","properties":{"fontCN":"黑体","sizePt":18,"spaceBefore":18,"spaceAfter":18}}\n' +
    '  例：用户说"正文用小4号宋体，行距为20磅"\n' +
    '  → {"action":"modify_style","targetStyle":"Normal","properties":{"fontCN":"宋体","sizePt":12,"lineSpacing":20,"lineSpacingType":"fixed"}}\n\n' +
    '### 中文字号→磅值对照（必须使用磅值，不能写字号名）：\n' +
    '初号=42, 小初=36, 一号=26, 小一=24, 二号=22, 小二=18, 三号=16, 小三=15, 四号=14, 小四=12, 五号=10.5, 小五=9\n\n' +
    '### targetStyle（必填）— 使用文档中实际存在的样式名：\n' +
    '- 中国用户一般使用中文样式名，请优先使用中文名：\n' +
    '  "标题 1"=章标题, "标题 2"=节标题, "标题 3"=小节标题, "标题 4"=小标题\n' +
    '  "正文"=正文段落, "题注"=图题/表题, "标题"=文档标题, "副标题"=副标题\n' +
    '- 英文 Word 才用 "Heading 1"~"Heading 6", "Normal", "Caption"\n\n' +
    '### properties（必填，至少填一个；仅填用户明确提到的属性）：\n' +
    '| 属性 | 类型 | 说明 | 示例值 |\n' +
    '|------|------|------|--------|\n' +
    '| fontCN | string | 中文字体 | "宋体","黑体","楷体","仿宋","微软雅黑" |\n' +
    '| sizePt | number | 字号(磅) | 见上方对照表 |\n' +
    '| bold | boolean | 加粗 | true/false |\n' +
    '| italic | boolean | 斜体 | true/false |\n' +
    '| lineSpacing | number | 行距值 | 配合 lineSpacingType 使用 |\n' +
    '| lineSpacingType | string | 行距类型 | "fixed"(固定值磅), "multiple"(多倍行距) |\n' +
    '| spaceBefore | number | 段前(磅) | "上下各空一行"=该级字号磅值, "段前12磅"=12 |\n' +
    '| spaceAfter | number | 段后(磅) | "上下各空一行"=该级字号磅值, "段后6磅"=6 |\n' +
    '| alignment | string | 对齐 | "left"/"center"/"right"/"justify" |\n' +
    '| firstLineIndent | number | 首行缩进(磅) | 12pt字号×2字符=24 |\n\n' +
    '### 关键规则（务必遵守）：\n' +
    '1. 用户提供的是格式修改请求 → 输出 JSON；不是 → 正常文字回复\n' +
    '2. 输出**纯 JSON**，每行一个样式，不要用数组或代码块包裹\n' +
    '3. properties 只填用户提到的属性\n' +
    '4. ★ 绝不输出 HTML 代码来"展示"格式——HTML 无法修改 Word 文档，必须用 JSON\n' +
    '5. "上下各空X行"=该级字号×X磅，"上下各空X磅"直接用X\n';

  /* ═══════════════════════════════════════════════════════════
     样式修改模块 — JSON 解析与 Office.js 执行
     ═══════════════════════════════════════════════════════════ */

  /**
   * 从 AI 响应中解析样式修改指令（支持多 JSON）
   *
   * 四层容错策略：
   *   1. 整个响应作为单 JSON 解析
   *   2. 按行拆分，逐行 JSON.parse
   *   3. 括号计数提取完整 JSON（正确处理嵌套 {} ）
   *   4. 提取 ```json ... ``` 代码块
   *
   * @param {string} response - AI 原始输出文本
   * @returns {Array<object>} 样式修改 action 对象数组（可能为空）
   */
  function parseStyleModificationJSONs(response) {
    if (!response || typeof response !== 'string') return [];

    var trimmed = response.trim();
    var results = [];

    // 第1层：整个响应就是单 JSON
    try {
      var obj = JSON.parse(trimmed);
      if (obj && obj.action === 'modify_style' && obj.targetStyle && obj.properties) {
        return [obj];
      }
    } catch (e) { /* continue */ }

    // 第2层：按行拆分，逐行解析（处理多 JSON 分行输出的情况）
    var lines = trimmed.split(/\r?\n/);
    for (var li = 0; li < lines.length; li++) {
      var line = lines[li].trim();
      if (!line || line[0] !== '{') continue;
      try {
        var lineObj = JSON.parse(line);
        if (lineObj && lineObj.action === 'modify_style' && lineObj.targetStyle && lineObj.properties) {
          results.push(lineObj);
        }
      } catch (e) { /* skip invalid lines */ }
    }
    if (results.length > 0) return results;

    // 第3层：括号计数提取完整 JSON 对象（正确处理嵌套 {} 和字符串内大括号）
    var depth = 0, inString = false, escape = false;
    var start = -1;
    for (var i = 0; i < trimmed.length; i++) {
      var ch = trimmed[i];
      if (escape) { escape = false; continue; }
      if (ch === '\\') { escape = true; continue; }
      if (ch === '"') { inString = !inString; continue; }
      if (inString) continue;

      if (ch === '{') {
        if (depth === 0) start = i;
        depth++;
      } else if (ch === '}') {
        depth--;
        if (depth === 0 && start >= 0) {
          var candidate = trimmed.substring(start, i + 1);
          try {
            var candObj = JSON.parse(candidate);
            if (candObj && candObj.action === 'modify_style' && candObj.targetStyle && candObj.properties) {
              results.push(candObj);
            }
          } catch (e) { /* skip */ }
          start = -1;
        }
      }
    }
    if (results.length > 0) return results;

    // 第4层：markdown 代码块
    var m = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (m) {
      try {
        var obj4 = JSON.parse(m[1].trim());
        if (obj4 && obj4.action === 'modify_style') return [obj4];
      } catch (e) {}
      // 代码块内可能也有多行 JSON
      var codeLines = m[1].trim().split(/\r?\n/);
      for (var cl = 0; cl < codeLines.length; cl++) {
        var cline = codeLines[cl].trim();
        if (!cline || cline[0] !== '{') continue;
        try {
          var clObj = JSON.parse(cline);
          if (clObj && clObj.action === 'modify_style') results.push(clObj);
        } catch (e) {}
      }
    }

    return results;
  }

  /**
   * 对单个段落对象应用格式属性修改
   *
   * 仅设置 props 中存在的属性，其他属性保持不变。
   * 与 applyStyleProperties 不同，此函数直接操作段落而非样式定义。
   *
   * @param {Word.Paragraph} p - Office.js 段落对象
   * @param {object} props - 属性键值对
   */
  function applyParagraphFormatting(p, props) {
    if (props.fontCN || props.fontEN) {
      p.font.name = props.fontCN || props.fontEN;
    }
    if (typeof props.sizePt === 'number') {
      p.font.size = props.sizePt;
    }
    if (typeof props.bold === 'boolean') {
      p.font.bold = props.bold;
    }
    if (typeof props.italic === 'boolean') {
      p.font.italic = props.italic;
    }
    if (typeof props.lineSpacing === 'number') {
      p.lineSpacing = props.lineSpacing;
    }
    // 段前/段后/对齐/缩进 → 使用标量属性避免 paragraphFormat 导航属性未加载的问题
    // 标量属性写入不需要显式加载；对带样式段落可能被样式定义覆盖，OOXML 注入提供可靠覆盖
    if (typeof props.spaceBefore === 'number') {
      p.spaceBefore = props.spaceBefore;
    }
    if (typeof props.spaceAfter === 'number') {
      p.spaceAfter = props.spaceAfter;
    }
    if (typeof props.alignment === 'string') {
      var alignMap = { left: 'Left', center: 'Centered', right: 'Right', justify: 'Justified' };
      p.alignment = alignMap[props.alignment] || props.alignment;
    }
    if (typeof props.firstLineIndent === 'number') {
      p.firstLineIndent = props.firstLineIndent;
    }
    if (typeof props.color === 'string') {
      p.font.color = props.color;
    }
  }

  /**
   * 检查是否包含需要通过 OOXML 覆盖的段落级属性
   *
   * Office.js 的 p.paragraphFormat.xxx 和 p.xxx 标量属性在带样式段落
   * （如标题）上无法可靠覆盖样式定义的段前段后/行距/对齐/缩进。
   * 这些属性必须通过 OOXML 直接注入 <w:pPr> 才能生效。
   *
   * @param {object} props - 格式属性
   * @returns {boolean}
   */
  
  /**
   * 从 Windows Flat OPC 中提取指定索引的段落 OOXML
   * getOoxml() 在 Mac/Web 返回段落级 OOXML，在 Windows 返回整个文档的 Flat OPC
   */
  function extractParagraphFromFlatOpc(flatOpc, paraIndex) {
    var bodyMatch = flatOpc.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/i);
    if (!bodyMatch) return null;
    var bodyContent = bodyMatch[1];
    var paraRegex = /<w:p\b[\s\S]*?<\/w:p>/gi;
    var paragraphs = [];
    var match;
    while ((match = paraRegex.exec(bodyContent)) !== null) {
      paragraphs.push(match[0]);
    }
    if (paraIndex >= 0 && paraIndex < paragraphs.length) {
      return paragraphs[paraIndex];
    }
    return null;
  }
function needsOoxmlOverride(props) {
    return typeof props.spaceBefore === 'number' ||
           typeof props.spaceAfter === 'number' ||
           typeof props.lineSpacing === 'number' ||
           typeof props.alignment === 'string' ||
           typeof props.firstLineIndent === 'number';
  }

  /**
   * 在段落 OOXML 的 <w:pPr> 中注入/替换段落级格式属性
   *
   * 原理：
   *   1. 获取段落的完整 OOXML（<w:p>...</w:p>）
   *   2. 在 <w:pPr> 内修改 <w:spacing> / <w:jc> / <w:ind>
   *   3. 通过 range.insertOoxml('Replace') 替换整个段落
   *
   * OOXML 间距单位换算：
   *   - 段前/段后: 1pt = 20 twips
   *   - 多倍行距:  1.0x = 240, 1.5x = 360, 2.0x = 480 (w:lineRule="auto")
   *   - 固定行距:  1pt = 20 twips                         (w:lineRule="exact")
   *   - 最小行距:  1pt = 20 twips                         (w:lineRule="atLeast")
   *
   * @param {string} xml - 段落 OOXML
   * @param {object} props - 格式属性 {spaceBefore, spaceAfter, lineSpacing, lineSpacingType, alignment, firstLineIndent}
   * @returns {string} 修改后的 OOXML
   */
  function injectParagraphOoxml(xml, props) {
    var twipsPerPt = 20;
    var insertXml = '';

    // ── 间距 <w:spacing> ──
    if (typeof props.spaceBefore === 'number' ||
        typeof props.spaceAfter === 'number' ||
        typeof props.lineSpacing === 'number') {
      var sa = [];
      if (typeof props.spaceBefore === 'number') {
        sa.push('w:before="' + Math.round(props.spaceBefore * twipsPerPt) + '"');
      }
      if (typeof props.spaceAfter === 'number') {
        sa.push('w:after="' + Math.round(props.spaceAfter * twipsPerPt) + '"');
      }
      if (typeof props.lineSpacing === 'number') {
        var lsType = props.lineSpacingType || 'multiple';
        var lineRule = 'auto';
        var lineVal;
        if (lsType === 'fixed') {
          lineRule = 'exact';
          lineVal = Math.round(props.lineSpacing * twipsPerPt);
        } else if (lsType === 'atLeast') {
          lineRule = 'atLeast';
          lineVal = Math.round(props.lineSpacing * twipsPerPt);
        } else {
          // multiple: 1.0 = 240
          lineRule = 'auto';
          lineVal = Math.round(props.lineSpacing * 240);
        }
        sa.push('w:line="' + lineVal + '"');
        sa.push('w:lineRule="' + lineRule + '"');
      }
      var oldSpacingMatch = xml.match(/<w:spacing\b[^>]*\/?>/i);
      if (oldSpacingMatch) {
        // 合并原有属性，保留不被本次修改涉及的属性（如 w:line/w:lineRule）
        // 防止只改 w:before/w:after 时丢弃原有的行距
        var mergedAttrs = {};
        var attrReSp = /(w:\w+)="([^"]*)"/g;
        var mSp;
        while ((mSp = attrReSp.exec(oldSpacingMatch[0])) !== null) {
          mergedAttrs[mSp[1]] = mSp[2];
        }
        for (var si = 0; si < sa.length; si++) {
          var eqSp = sa[si].indexOf('=');
          if (eqSp > 0) {
            mergedAttrs[sa[si].substring(0, eqSp)] = sa[si].substring(eqSp + 2, sa[si].length - 1);
          }
        }
        var mergedParts = [];
        for (var mk in mergedAttrs) {
          if (mergedAttrs.hasOwnProperty(mk)) {
            mergedParts.push(mk + '="' + mergedAttrs[mk] + '"');
          }
        }
        xml = xml.replace(/<w:spacing\b[^>]*\/?>/i, '<w:spacing ' + mergedParts.join(' ') + '/>');
      } else {
        insertXml += '<w:spacing ' + sa.join(' ') + '/>';
      }
    }

    // ── 对齐 <w:jc> ──
    if (typeof props.alignment === 'string') {
      var alignMap = { left: 'left', center: 'center', right: 'right', justify: 'both' };
      var jcTag = '<w:jc w:val="' + (alignMap[props.alignment] || props.alignment) + '"/>';
      if (/<w:jc\b[^>]*\/?>/i.test(xml)) {
        xml = xml.replace(/<w:jc\b[^>]*\/?>/i, jcTag);
      } else {
        insertXml += jcTag;
      }
    }

    // ── 首行缩进 <w:ind> ──
    if (typeof props.firstLineIndent === 'number') {
      var newIndParts = ['w:firstLine="' + Math.round(props.firstLineIndent * twipsPerPt) + '"'];
      var oldIndMatch = xml.match(/<w:ind\b[^>]*\/?>/i);
      if (oldIndMatch) {
        // 合并原有属性，保留不被本次修改涉及的属性（如 w:left/w:right/w:hanging）
        var mergedInd = {};
        var attrReInd = /(w:\w+)="([^"]*)"/g;
        var mInd;
        while ((mInd = attrReInd.exec(oldIndMatch[0])) !== null) {
          mergedInd[mInd[1]] = mInd[2];
        }
        for (var ii = 0; ii < newIndParts.length; ii++) {
          var eqInd = newIndParts[ii].indexOf('=');
          if (eqInd > 0) {
            mergedInd[newIndParts[ii].substring(0, eqInd)] = newIndParts[ii].substring(eqInd + 2, newIndParts[ii].length - 1);
          }
        }
        var indResult = [];
        for (var ik in mergedInd) {
          if (mergedInd.hasOwnProperty(ik)) {
            indResult.push(ik + '="' + mergedInd[ik] + '"');
          }
        }
        xml = xml.replace(/<w:ind\b[^>]*\/?>/i, '<w:ind ' + indResult.join(' ') + '/>');
      } else {
        insertXml += '<w:ind ' + newIndParts.join(' ') + '/>';
      }
    }

    if (!insertXml) return xml;

    // 插入到 <w:pPr> 内（紧接开始标签之后）
    if (/<w:pPr\b/i.test(xml)) {
      return xml.replace(/(<w:pPr\b[^>]*>)/i, '$1' + insertXml);
    }

    // 无 <w:pPr>：在 <w:p> 开始标签后创建
    return xml.replace(/(<w:p\b[^>]*>)/i, '$1<w:pPr>' + insertXml + '</w:pPr>');
  }

  /**
   * 执行样式修改 — 遍历全文段落，对匹配样式的段落应用格式
   *
   * Office.js 的 Style.font / Style.paragraphFormat 为只读，
   * 无法直接通过 getByName() 修改样式定义。
   * 改为遍历所有段落，检查 style 属性，对匹配的段落直接设格式。
   *
   * @param {object} action - 样式修改指令 {targetStyle: string, properties: object}
   * @returns {Promise<{styleName: string, count: number}>}
   */
  /**
   * 样式名归一化
   *
   * Word 中文版内置样式名使用  （非断行空格），而非普通空格  。
   * 如段落实际样式名为 "标题 2" 而 AI 输出 "标题 2"，精确匹配会失败。
   * 该函数将各种 Unicode 空白字符统一替换为普通空格，确保匹配不受空格编码差异影响。
   *
   * @param {string} s - 原始样式名
   * @returns {string} 归一化后的样式名
   */
  function canonStyleName(s) {
    return (s || '').replace(/[  -  　\s]+/g, ' ').trim();
  }

  /**
   * 样式名匹配归一化：去除所有 Unicode 空白字符并转小写
   * 用于处理 "标题2" vs "标题 2" vs "标题 2" 的互认
   */
  function normStyleKey(s) {
    return (s || '').replace(/[  -  　\s]+/g, '').toLowerCase();
  }

  function executeStyleModification(action) {
    var name = action.targetStyle;
    var props = action.properties;

    // 构建可能的样式名列表（英文 + 中文，双向覆盖）
    var targetNames = [name];
    var cnStyleNameMap = {
      'Heading 1': '标题 1', 'Heading 2': '标题 2', 'Heading 3': '标题 3',
      'Heading 4': '标题 4', 'Heading 5': '标题 5', 'Heading 6': '标题 6',
      'Normal': '正文',
      'Caption': '题注',
      'Title': '标题',
      'Subtitle': '副标题',
      'TOC Heading': '目录标题',
      'TOC 1': '目录 1', 'TOC 2': '目录 2', 'TOC 3': '目录 3',
      'List Paragraph': '列表段落',
      'Table Normal': '普通表格'
    };
    // 双向：英文→中文 和 中文→英文 都加入候选
    if (cnStyleNameMap[name]) targetNames.push(cnStyleNameMap[name]);
    // 反向查：如果 AI 输出了中文名，也加入英文名
    var canonName = canonStyleName(name);
    var compactName = normStyleKey(name);
    for (var enName in cnStyleNameMap) {
      var cnCanon = canonStyleName(cnStyleNameMap[enName]);
      var cnCompact = normStyleKey(cnStyleNameMap[enName]);
      if ((cnCanon === canonName || cnCompact === compactName) && targetNames.indexOf(enName) < 0) {
        targetNames.push(enName);
      }
    }
    // 归一化所有候选名用于匹配（处理 Word 非断行空格 vs 普通空格 vs 无空格差异）
    var targetCanons = {};
    for (var ti = 0; ti < targetNames.length; ti++) {
      targetCanons[canonStyleName(targetNames[ti])] = true;
      targetCanons[normStyleKey(targetNames[ti])] = true;
    }

    // 判断是否需要走 OOXML 路径覆盖段落级属性
    var useOoxml = needsOoxmlOverride(props);

    return Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      // 'items' 加载所有标量属性（含 style，用于样式名匹配）
      // 注意：paragraphFormat 和 font 一样是导航属性，写入不需要显式加载；
      // 不能同时调用 paragraphs.load('paragraphFormat')，会覆盖上面的 items 加载导致 p.style 为空
      context.load(paragraphs, 'items');

      return context.sync().then(function () {
        var modifiedCount = 0;
        // 收集文档中所有唯一样式名（用于无匹配时提示）
        var allStylesInDoc = {};
        // OOXML 队列：{para, ooxml}
        var ooxmlQueue = [];

        console.log('OfficeAI: executeStyleModification target=' + name + ' candidates=' + JSON.stringify(targetNames) + ' useOoxml=' + useOoxml);

        for (var i = 0; i < paragraphs.items.length; i++) {
          var p = paragraphs.items[i];
          var styleName = '';
          try { styleName = p.style || ''; } catch (e) { /* skip */ }

          if (styleName) {
            allStylesInDoc[styleName] = (allStylesInDoc[styleName] || 0) + 1;
            if (targetCanons[canonStyleName(styleName)] || targetCanons[normStyleKey(styleName)]) {
              console.log('OfficeAI: matched paragraph #' + i + ' style="' + styleName + '" applying props:', JSON.stringify(props));
              // 阶段1: 字体属性走 Office.js API（已验证对样式段落有效）
              try {
                applyParagraphFormatting(p, props);
              } catch (e) {
                console.warn('OfficeAI: applyParagraphFormatting threw for paragraph #' + i, e);
              }
              // 阶段1b: 段落级属性（间距/对齐/缩进）排队获取 OOXML
              if (useOoxml) {
                try {
                  var ooxmlResult = p.getOoxml(); console.log('OfficeAI: getOoxml queued for paragraph #' + i + ', result=' + (ooxmlResult ? 'ClientResult' : 'null')); ooxmlQueue.push({ para: p, ooxml: ooxmlResult });
                } catch (e) {
                  console.warn('OfficeAI: getOoxml queue failed for paragraph #' + i, e);
                }
              }
              modifiedCount++;
            }
          }
        }

        console.log('OfficeAI: executeStyleModification loop done, ooxmlQueue.length=' + ooxmlQueue.length + ' modifiedCount=' + modifiedCount);
          return context.sync().then(function () {
          // 阶段2: 通过 OOXML 替换注入段落级格式（间距/对齐/缩进）
          var ooxmlApplied = 0;
          console.log('OfficeAI: OOXML phase entry check, useOoxml=' + useOoxml + ' queueLen=' + ooxmlQueue.length);
          if (useOoxml && ooxmlQueue.length > 0) {
            console.log('OfficeAI: OOXML phase start, queue=' + ooxmlQueue.length + ' props=' + JSON.stringify(props));
            for (var j = 0; j < ooxmlQueue.length; j++) {
              var entry = ooxmlQueue[j];
              var oxml = '';
              try { oxml = entry.ooxml.value || ''; } catch (e) {
                console.warn('OfficeAI: getOoxml read failed for paragraph #' + j, e);
                continue;
              }
              if (!oxml) {
                console.warn('OfficeAI: getOoxml returned empty for paragraph #' + j);
                continue;
              }
              // Windows returns Flat OPC; extract target paragraph OOXML
              if (oxml.indexOf('<pkg:package') >= 0) {
                oxml = extractParagraphFromFlatOpc(oxml, entry.paraIndex) || oxml;
                if (!oxml || oxml.indexOf('<pkg:package') >= 0) {
                  console.warn('OfficeAI: failed to extract paragraph from Flat OPC for entry #' + j);
                  continue;
                }
              }
              if (!oxml) {
                console.warn('OfficeAI: getOoxml returned empty for paragraph #' + j);
                continue;
              }
              try {
                var modified = injectParagraphOoxml(oxml, props);
                // 第一个段落输出前后对比，方便诊断
                if (j === 0) {
                  console.log('OfficeAI: OOXML[0] BEFORE (' + oxml.length + ' chars): ' + oxml.substring(0, 500));
                  console.log('OfficeAI: OOXML[0] AFTER  (' + modified.length + ' chars): ' + modified.substring(0, 500));
                }
                // 用 Range.insertOoxml('Replace') 替换整个段落
                var range = entry.para.getRange('Whole');
                range.insertOoxml(modified, 'Replace');
                ooxmlApplied++;
              } catch (e2) {
                console.warn('OfficeAI: insertOoxml failed for paragraph #' + j, e2);
              }
            }
            console.log('OfficeAI: OOXML phase done, applied=' + ooxmlApplied);
          }

          return context.sync().then(function () {
            var styleList = Object.keys(allStylesInDoc).sort(function (a, b) {
              return (allStylesInDoc[b] || 0) - (allStylesInDoc[a] || 0);
            });
            console.log('OfficeAI: executeStyleModification done. modified=' + modifiedCount + ' ooxmlApplied=' + ooxmlApplied + ' totalStyles=' + styleList.length, styleList.slice(0, 10));
            return { styleName: name, count: modifiedCount, allStyles: styleList };
          });
        });
      });
    });
  }

  /**
   * 构建样式修改成功的人类可读确认消息
   *
   * @param {object} action - 样式修改指令
   * @param {number} count - 实际修改的段落数
   * @returns {string} 格式化的确认消息
   */
  function buildStyleModificationMessage(action, count, allStyles) {
    var props = action.properties;
    var parts = [];

    if (props.fontCN) parts.push('中文字体=' + props.fontCN);
    if (props.fontEN) parts.push('英文字体=' + props.fontEN);
    if (typeof props.sizePt === 'number') parts.push('字号=' + props.sizePt + 'pt');
    if (typeof props.bold === 'boolean') parts.push(props.bold ? '加粗' : '取消加粗');
    if (typeof props.italic === 'boolean') parts.push(props.italic ? '斜体' : '取消斜体');
    if (typeof props.lineSpacing === 'number') {
      var lsLabel = '';
      if (props.lineSpacingType === 'fixed') lsLabel = '固定值';
      else if (props.lineSpacingType === 'multiple') lsLabel = '倍行距';
      else if (props.lineSpacingType === 'atLeast') lsLabel = '最小值';
      var unit = (props.lineSpacingType === 'multiple') ? '倍' : 'pt';
      parts.push('行距=' + (lsLabel ? lsLabel : '') + props.lineSpacing + unit);
    }
    if (typeof props.spaceBefore === 'number') parts.push('段前=' + props.spaceBefore + 'pt');
    if (typeof props.spaceAfter === 'number') parts.push('段后=' + props.spaceAfter + 'pt');
    if (props.alignment) {
      var alignLabel = { left: '左对齐', center: '居中', right: '右对齐', justify: '两端对齐' };
      parts.push(alignLabel[props.alignment] || props.alignment);
    }
    if (typeof props.firstLineIndent === 'number') parts.push('首行缩进=' + props.firstLineIndent + 'pt');
    if (props.color) parts.push('颜色=' + props.color);

    var cnStyleMap = {
      'Heading 1': '标题1', 'Heading 2': '标题2', 'Heading 3': '标题3',
      'Heading 4': '标题4', 'Heading 5': '标题5', 'Heading 6': '标题6',
      'Normal': '正文', 'Caption': '题注（图题/表题）', 'Title': '标题',
      'Subtitle': '副标题'
    };
    var displayName = cnStyleMap[action.targetStyle] || action.targetStyle;

    if (count === 0) {
      var hint = '⚠️ 未找到使用「' + displayName + '」样式的段落。\n\n' +
        '已尝试匹配的样式名：' + action.targetStyle;
      // 显示文档中实际存在的样式
      if (allStyles && allStyles.length > 0) {
        var topStyles = allStyles.slice(0, 12);
        hint += '\n\n📋 文档中实际存在的样式（按使用量排序）：\n   ' + topStyles.join('、');
        if (allStyles.length > 12) hint += ' …等' + allStyles.length + '种';
        hint += '\n\n💡 请从上面列表中复制准确的样式名后重新发送指令。';
      }
      return hint;
    }

    return '✅ 已修改 ' + count + ' 个使用「' + displayName + '」样式的段落：' + parts.join('，') +
      '。\n\n📌 修改已直接应用于段落格式。\n🔄 可按 Ctrl+Z 撤销本次修改。';
  }

  /* ═══════════════════════════════════════════════════════════
     插入后样式归一化 — insertHtml 后对段落补设 Word 样式名
     ═══════════════════════════════════════════════════════════ */

  /**
   * 中英文样式名双向映射表
   */
  var STYLE_NAME_MAP_CN_EN = {
    'Heading 1': '标题 1', 'Heading 2': '标题 2', 'Heading 3': '标题 3',
    'Heading 4': '标题 4', 'Heading 5': '标题 5', 'Heading 6': '标题 6',
    'Normal': '正文',
    'Caption': '题注',
    'Title': '标题',
    'Subtitle': '副标题',
    'TOC 1': '目录 1', 'TOC 2': '目录 2', 'TOC 3': '目录 3',
    'List Paragraph': '列表段落',
    'Table Normal': '普通表格',
    'Quote': '引文',
    'Intense Quote': '明显引用',
    'List Bullet': '要点',
    'List Number': '编号列表'
  };

  /**
   * 根据段落内容和格式，检测应使用的 Word 样式名
   *
   * @param {Word.Paragraph} p - 段落对象
   * @param {string} text - 段落文本
   * @param {number} index - 在插入块中的位置（0-based）
   * @param {number} total - 插入块的总段落数
   * @returns {string} 目标样式名（英文）
   */
  function detectTargetStyle(p, text, index, total) {
    // 读取当前已有样式
    var currentStyle = '';
    try { currentStyle = p.style || ''; } catch (e) {}

    // 已是内置标题 → 保持不变
    if (/^Heading \d$|^标题 \d$/i.test(canonStyleName(currentStyle))) return null;

    // 读取格式属性
    var fontSize = 0, isBold = false, isItalic = false;
    try { fontSize = p.font.size || 0; } catch (e) {}
    try { isBold = p.font.bold; } catch (e) {}
    try { isItalic = p.font.italic; } catch (e) {}

    // --- 题注 (Caption) ---
    // 短文本、以"图/表/Fig/Table"开头
    if (text.length < 100 && /^(图|表|Fig|Figure|Table)\s*[\.\d:：]/.test(text)) {
      return 'Caption';
    }

    // --- 标题 (Title) ---
    // 第一个段落、大字加粗
    if (index === 0 && isBold && fontSize >= 18) {
      return 'Title';
    }

    // --- 副标题 (Subtitle) ---
    // 第二个段落、较大字号
    if (index === 1 && fontSize >= 14 && fontSize < 18 && total > 1) {
      return 'Subtitle';
    }

    // --- 目录 (TOC) ---
    // 含点划线或页码模式
    if (/\.{3,}\s*\d+$|……\s*\d+$/.test(text) && fontSize >= 10 && fontSize <= 14) {
      return 'TOC 1';
    }

    // --- 引文 (Quote) ---
    // 斜体、短段落
    if (isItalic && text.length < 200) {
      return 'Quote';
    }

    // --- 列表段落 ---
    // 项目符号或编号开头
    if (/^[•\-\*▪▸●○]\s/.test(text) || /^\d+[\.\)、]\s/.test(text)) {
      return 'List Paragraph';
    }

    // --- 非标题段落：确认不是 Heading → 设正文 ---
    if (!/Heading|标题/.test(currentStyle)) {
      return 'Normal';
    }

    return null;
  }

  /**
   * 设置段落的样式名（兼容中英文 Office）
   *
   * @param {Word.Paragraph} p - 段落对象
   * @param {string} enStyle - 英文样式名
   */
  function setParagraphStyle(p, enStyle) {
    if (!enStyle) return;
    try {
      p.style = enStyle;
    } catch (e1) {
      var cnStyle = STYLE_NAME_MAP_CN_EN[enStyle] || enStyle;
      try { p.style = cnStyle; } catch (e2) { /* skip */ }
    }
  }

  // ── Default paragraph spacing per style (spaceBefore/spaceAfter in pt) ──
  var STYLE_SPACING = {
    "Heading 1":       { spaceBefore: 12, spaceAfter: 6, lineSpacing: 18, lineSpacingType: "fixed" },
    "Heading 2":       { spaceBefore: 10, spaceAfter: 6, lineSpacing: 16, lineSpacingType: "fixed" },
    "Heading 3":       { spaceBefore: 8,  spaceAfter: 6, lineSpacing: 14, lineSpacingType: "fixed" },
    "Title":            { spaceBefore: 6,  spaceAfter: 12 },
    "Subtitle":         { spaceBefore: 4,  spaceAfter: 8 },
    "Normal":           { spaceBefore: 0,  spaceAfter: 6, lineSpacing: 1.5, lineSpacingType: "multiple" },
    "List Paragraph":   { spaceBefore: 0,  spaceAfter: 3 },
    "Quote":            { spaceBefore: 6,  spaceAfter: 6, lineSpacing: 1.0, lineSpacingType: "multiple" },
    "Caption":          { spaceBefore: 4,  spaceAfter: 2, lineSpacing: 1.2, lineSpacingType: "multiple" }
  };

  /**
   * 对刚插入的段落应用样式归一化
   * 原理: insertHtml 会把 <h1>~<h6> 转为标题样式，但 <p> 不会得到"正文"样式。
   *       此函数遍历最后 N 个非空段落，检测每段内容/格式，补设正确的 Word 样式名。
   *       紧接着设置样式名后，自动应用对应的默认段落间距（spaceBefore/spaceAfter/lineSpacing）。
   *       对带样式的段落（如标题）通过 OOXML 注入确保格式生效。
   * @param {number} paraCountBefore - 插入前的段落总数
   * @returns {Promise<number>} 成功设置样式的段落数
   */
  function normalizeInsertedParagraphs(paraCountBefore) {
    return Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, "items");
      return context.sync().then(function () {
        var items = paragraphs.items;
        var total = items.length;
        var startIdx = Math.max(0, paraCountBefore);
        var applied = 0;
        var ooxmlQueue = [];

        for (var i = startIdx; i < total; i++) {
          var p = items[i];
          var text = "";
          try { text = (p.text || "").trim(); } catch (e) {}
          if (text.length === 0) continue;

          var targetStyle = detectTargetStyle(p, text, i - startIdx, total - startIdx);
          if (targetStyle) {
            setParagraphStyle(p, targetStyle);
            applied++;
            // 应用该样式对应的默认段落间距
            var spacing = STYLE_SPACING[targetStyle];
            if (spacing) {
              try {
                applyParagraphFormatting(p, spacing);
              } catch (e) {
                console.warn("OfficeAI: applyParagraphFormatting threw for paragraph #" + i, e);
              }
              // 对段落级属性（段前/段后/行距/对齐）排队 OOXML 覆盖
              if (needsOoxmlOverride(spacing)) {
                try {
                  var ooxmlResult = p.getOoxml(); console.log("OfficeAI: normalizeInsertedParagraphs getOoxml queued for paragraph #" + i); ooxmlQueue.push({ para: p, ooxml: ooxmlResult, props: spacing, paraIndex: i });
                } catch (e) {
                  console.warn("OfficeAI: getOoxml queue failed for paragraph #" + i, e);
                }
              }
            }
          }
        }

        return context.sync().then(function () {
          var ooxmlApplied = 0;
          if (ooxmlQueue.length > 0) {
            for (var j = 0; j < ooxmlQueue.length; j++) {
              var entry = ooxmlQueue[j];
              var oxml = "";
              try { oxml = entry.ooxml.value || ""; } catch (e) {
                console.warn("OfficeAI: getOoxml read failed for paragraph #" + j, e);
                continue;
              }
              if (!oxml) continue;
 `r`n              // Windows returns Flat OPC; extract target paragraph OOXML`r`n              if (oxml.indexOf('<pkg:package') >= 0) {`r`n                oxml = extractParagraphFromFlatOpc(oxml, entry.paraIndex) || oxml;`r`n                if (!oxml || oxml.indexOf('<pkg:package') >= 0) {`r`n                  console.warn("OfficeAI: normalizeInsertedParagraphs Flat OPC extraction failed for entry #" + j);`r`n                  continue;`r`n                }`r`n              }             try {
                var modified = injectParagraphOoxml(oxml, entry.props);
                var range = entry.para.getRange("Whole");
                range.insertOoxml(modified, "Replace");
                ooxmlApplied++;
              } catch (e2) {
                console.warn("OfficeAI: insertOoxml failed for paragraph #" + j, e2);
              }
            }
            return context.sync().then(function () {
              console.log("OfficeAI: normalizeInsertedParagraphs applied=" + applied + " styles + " + ooxmlApplied + " OOXML overrides over " + (total - startIdx) + " new paragraphs");
              return applied;
            });
          }
          console.log("OfficeAI: normalizeInsertedParagraphs applied=" + applied + " styles over " + (total - startIdx) + " new paragraphs");
          return applied;
        });
      });
    });
  }  /**
   * 获取当前文档段落总数（插入前快照）
   * @returns {Promise<number>}
   */
  function getParagraphCount() {
    return Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, 'items');
      return context.sync().then(function () {
        return paragraphs.items.length;
      });
    }).catch(function () { return 0; });
  }

  /* ═══════════════════════════════════════════════════════════
     原生格式预设 — 标题检测
     ═══════════════════════════════════════════════════════════ */

  /**
   * 使用正则匹配段落文本的标题级别
   *
   * 检测策略（优先级从高到低）：
   *   1. 显式标题关键词
   *   2. 编号模式（第X章、一、、1.、1.1、(一)）
   *   3. 短行启发式：≤25字符且无句末标点 → 可能为标题
   *
   * @param {string} text - 段落文本（已 trim）
   * @returns {number} 1/2/3 = 标题级别，0 = 非标题（正文）
   */
  function detectHeadingLevel(text) {
    var clean = text.replace(/^\s+/, '').slice(0, 80);

    // Level 1: 显式标题关键词
    if (/^(摘要|绪论|引言|前言|概述|介绍|背景|总则|目的|范围|定义|总结|结论|参考文献|致谢|附录)[\s:：]*$/.test(clean)) return 1;
    if (/^第[一二三四五六七八九十百千\d]+[章篇部]/.test(clean)) return 1;

    // Level 2: 第X节 / 一、二、三 / 1. 2. 数字编号
    if (/^第[一二三四五六七八九十百千\d]+节/.test(clean)) return 2;
    if (/^[一二三四五六七八九十]+[、．.\s]/.test(clean)) return 2;
    if (/^\d+[\.\、\s)]/.test(clean)) return 2;

    // Level 3: 1.1 / (一) / 小标题
    if (/^\d+\.\d+/.test(clean)) return 3;
    if (/^（[一二三四五六七八九十\d]+）/.test(clean)) return 3;

    // 短行启发式：≤25 字符、无句末标点 → 疑似标题（level 2）
    if (clean.length <= 25 && !/[。！？；，\.!\?;,]$/.test(clean)) {
      // 排除明显是正文的短行（含多个句子，或以"的""了""是"结尾）
      if (!/[的了是]$/.test(clean)) return 2;
    }

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
    clearStatus(el.presetsStatus);
    var label = PRESET_LABELS[presetKey] || presetKey;
    showStatus(el.presetsStatus, 'info', '正在应用「' + label + '」...');

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
      showStatus(el.presetsStatus, 'success', '✓ 「' + label + '」已应用。');
    }).catch(function (err) {
      console.error('executeNativeFormatPreset [' + presetKey + '] error:', err);
      var detail = err.message || '未知错误';
      if (err.code) { detail = err.code + ': ' + detail; }
      if (err.debugInfo) { console.error('debugInfo:', err.debugInfo); }
      showStatus(el.presetsStatus, 'error', '格式应用失败: ' + detail);
    });
  }

  /**
   * 设置段落的内置标题样式（兼容中英文 Office 版本）
   *
   * Word 内置样式名随 Office 语言变化：
   *   英文版: "Heading 1"  中文版: "标题 1"
   * styleBuiltIn 属性（WordApi 1.3+）使用语言无关的键名如 "Heading1"
   *
   * @param {Word.Paragraph} p - 段落对象
   * @param {number} level - 标题级别 1/2/3
   * @returns {boolean} 是否成功设置
   */
  function setHeadingStyle(p, level) {
    // 方法 1: styleBuiltIn（语言无关，WordApi 1.3+）
    if (p.styleBuiltIn !== undefined && p.styleBuiltIn !== null) {
      try {
        p.styleBuiltIn = 'Heading' + level;
        return true;
      } catch (e) { /* 回退到下一方法 */ }
    }

    // 方法 2: style 属性 — 英文名
    try {
      p.style = 'Heading ' + level;
      return true;
    } catch (e) { /* 回退到下一方法 */ }

    // 方法 3: style 属性 — 中文名
    var cnNames = { 1: '标题 1', 2: '标题 2', 3: '标题 3' };
    try {
      p.style = cnNames[level];
      return true;
    } catch (e) {
      console.warn('setHeadingStyle: 所有方法均失败 (level=' + level + ')');
      return false;
    }
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
          if (hl > 0) {
            setHeadingStyle(p, hl);
            p.font.bold = true;
            p.alignment = 'Centered';
          }
          break;

        case 'format_indent':
          if (hl === 0) {
            p.firstLineIndent = opts.fontSize * opts.indentChars;
          }
          break;

        case 'format_spacing':
          p.spaceBefore = (hl >= 1) ? 12 : 0;
          p.spaceAfter = 6;
          break;

        case 'format_font':
          p.font.name = opts.cnFont;
          if (hl >= 1) {
            p.font.size = opts.fontSize + (4 - hl) * 2;
          } else {
            p.font.size = opts.fontSize;
            p.lineSpacing = opts.fontSize * opts.lineSpacing * 1.2;
          }
          break;
      }
    }
  }

  /**
   * AI API 调用核心（纯函数，不操作 UI）
   *
   * @param {Array<{role: string, content: string}>} messages - 完整的消息数组
   * @param {object} cfg - 配置对象（从 getConfigFromUI() 获取）
   * @returns {Promise<string>} AI 返回的文本内容
   */
  function executeAI(messages, cfg) {
    if (cfg.provider === 'ollama') {
      return fetch(cfg.apiBaseUrl + '/api/chat', {
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
    }

    // OpenAI 兼容 API（DeepSeek / 自定义）
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

    return fetch(endpoint, {
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

  /* ═══════════════════════════════════════════════════════════
     聊天引擎 — 消息发送、UI 渲染、上下文管理
     ═══════════════════════════════════════════════════════════ */

  /**
   * 在聊天区渲染一条消息气泡
   *
   * @param {string} role - 'user' | 'assistant'
   * @param {string} content - 消息内容
   */
  function renderChatMessage(role, content) {
    // 隐藏空状态提示
    el.chatEmptyHint.style.display = 'none';

    var msgDiv = document.createElement('div');
    msgDiv.className = 'chat-message ' + role;

    var label = document.createElement('div');
    label.className = 'chat-label';
    label.textContent = role === 'user' ? '👤 你' : '🤖 AI';

    var bubble = document.createElement('div');
    bubble.className = 'chat-bubble';
    bubble.textContent = content;

    msgDiv.appendChild(label);
    msgDiv.appendChild(bubble);
    el.chatMessages.appendChild(msgDiv);

    // 滚动到底部
    scrollChatToBottom();
  }

  /**
   * 聊天区滚动到底部
   */
  function scrollChatToBottom() {
    // 使用父级 chatSection 的滚动
    var section = el.chatMessages.parentElement;
    if (section) {
      section.scrollTop = section.scrollHeight;
    }
  }

  /**
   * 发送消息 — 聊天核心入口
   *
   * 处理流程：
   *   1. 获取上下文（选区 > 光标段落 > 文档摘要 > 拒绝）
   *   2. 渲染用户消息气泡
   *   3. 构建带 history 的 messages 数组
   *   4. 调用 executeAI
   *   5. 渲染 AI 回复气泡 + 更新 history
   *
   * @param {string}  userMessage   - 用户输入的消息文本
   * @param {string} [systemOverride] - 可选，覆盖默认 system prompt（AI 预设使用）
   * @param {string} [contextText]    - 可选，预确定的上下文文本（AI 预设使用，跳过上下文探测）
   */
  function sendMessage(userMessage, systemOverride, contextText) {
    if (!userMessage || !userMessage.trim()) return;

    clearStatus(el.chatStatus);
    el.sendBtn.disabled = true;
    el.chatInput.disabled = true;
    setSpinner(el.chatSpinner, true);

    // ═══ 第1步：确定上下文 ═══
    var contextPromise;
    if (contextText !== undefined) {
      // AI 预设已自带上下文，直接使用
      contextPromise = Promise.resolve({ text: contextText, hasStyles: false });
    } else {
      // 优先级：选区（带格式） > 光标段落 > 文档摘要 > 空字符串
      contextPromise = getSelectedTextWithStyles().then(function (result) {
        if (result.styledText) {
          // 选区有内容 → 使用带格式标注的文本
          return { text: result.styledText, hasStyles: true };
        }

        if (currentCursorContext) {
          // 无选区但有光标段落
          return { text: currentCursorContext, hasStyles: false };
        }

        // 尝试文档摘要
        return getDocumentText().then(function (docText) {
          if (docText) return { text: docText.slice(0, 3000), hasStyles: false };
          return { text: '', hasStyles: false };
        });
      });
    }

    contextPromise.then(function (ctx) {
      // ═══ 第2步：构建 system prompt ═══
      var cfg = getConfigFromUI();
      var basePrompt = systemOverride || cfg.systemPrompt;
      // 非预设模式：附加样式修改识别能力（预设模式不受影响）
      var systemPrompt = systemOverride ? basePrompt : (STYLE_MODIFY_SYSTEM_ADDON + '\n\n' + basePrompt);

      // 空选区时自动追加指令前缀
      if (!contextText && !ctx.text && !systemOverride) {
        systemPrompt = '请基于当前文档内容回答用户的问题。如果文档内容不足以回答问题，请如实说明。\n\n' + systemPrompt;
      }

      // ═══ 第3步：构建 user message（含上下文注入） ═══
      var finalUserContent = userMessage;
      if (ctx.text && !contextText) {
        if (ctx.hasStyles) {
          // ★ 跨段选中 → 在 user message 中附带格式映射说明
          finalUserContent =
            '【参考上下文 - 段落格式已标注】：\n"""\n' + ctx.text + '\n"""\n\n' +
            '【格式映射】：[标题 1]→<h1>  [标题 2]→<h2>  [标题 3]→<h3>  [正文]→<p style="text-indent:2em">  [题注]→<p>\n' +
            '请保留原文段落结构，每段用对应的 HTML 标签包裹后返回。\n\n' +
            '【用户问题】：' + userMessage;
        } else {
          // 光标段落或全文 → 纯文本上下文
          finalUserContent = '【参考上下文】：\n"""\n' + ctx.text + '\n"""\n\n【用户问题】：' + userMessage;
        }
      }

      // ═══ 第4步：构建完整 messages 数组（system + history + 当前 user） ═══
      var messages = [{ role: 'system', content: systemPrompt }];

      // 追加对话历史
      for (var i = 0; i < conversationHistory.length; i++) {
        messages.push(conversationHistory[i]);
      }

      // 追加当前用户消息
      messages.push({ role: 'user', content: finalUserContent });

      // ═══ 第5步：渲染用户消息 ═══
      renderChatMessage('user', userMessage);

      // ═══ 第6步：调用 AI ═══
      return executeAI(messages, cfg).then(function (response) {
        // ★ 检测是否为样式修改指令（支持多个样式同时修改）
        var styleActions = parseStyleModificationJSONs(response);

        if (styleActions.length > 0) {
          // 样式修改模式：逐个执行，避免 Office.js 并发冲突
          var results = [];
          function executeNext(idx) {
            if (idx >= styleActions.length) {
              // 全部完成 → 聚合确认消息
              var msgs = [];
              for (var ri = 0; ri < results.length; ri++) {
                msgs.push(buildStyleModificationMessage(styleActions[ri], results[ri].count, results[ri].allStyles));
              }
              var combinedMsg = msgs.join('\n\n');

              renderChatMessage('assistant', combinedMsg);
              conversationHistory.push({ role: 'user', content: finalUserContent });
              conversationHistory.push({ role: 'assistant', content: combinedMsg });
              while (conversationHistory.length > MAX_HISTORY_ROUNDS * 2) {
                conversationHistory.shift();
                conversationHistory.shift();
              }
              _lastAIResponse = combinedMsg;
              el.insertLastBtn.style.display = '';
              clearStatus(el.chatStatus);
              return;
            }

            return executeStyleModification(styleActions[idx]).then(function (result) {
              results.push(result);
              return executeNext(idx + 1);
            }).catch(function (err) {
              // 单个样式失败不终止，继续执行后续
              console.error('executeStyleModification [' + styleActions[idx].targetStyle + '] error:', err);
              results.push({ styleName: styleActions[idx].targetStyle, count: 0, error: err.message });
              return executeNext(idx + 1);
            });
          }

          return executeNext(0).catch(function (err) {
            console.error('executeStyleModification chain error:', err);
            var errMsg = '❌ 样式修改失败：' + (err.message || '未知错误') +
              '\n\n请确认样式名称正确。可用的内置样式：Heading 1~6, Normal';
            renderChatMessage('assistant', errMsg);
            conversationHistory.push({ role: 'user', content: finalUserContent });
            conversationHistory.push({ role: 'assistant', content: errMsg });
            while (conversationHistory.length > MAX_HISTORY_ROUNDS * 2) {
              conversationHistory.shift();
              conversationHistory.shift();
            }
            _lastAIResponse = errMsg;
            el.insertLastBtn.style.display = '';
            clearStatus(el.chatStatus);
          });
        }

        // 正常聊天回复
        renderChatMessage('assistant', response);

        // 更新对话历史（最多 10 轮 = 20 条）
        conversationHistory.push({ role: 'user', content: finalUserContent });
        conversationHistory.push({ role: 'assistant', content: response });
        while (conversationHistory.length > MAX_HISTORY_ROUNDS * 2) {
          conversationHistory.shift();
          conversationHistory.shift();
        }

        // 缓存最后一条 AI 回复（供插入/替换使用）
        _lastAIResponse = response;
        el.insertLastBtn.style.display = '';

        clearStatus(el.chatStatus);
      });
    }).catch(function (err) {
      console.error('sendMessage error:', err);
      showStatus(el.chatStatus, 'error', '发送失败: ' + (err.message || '未知错误'));
    }).finally(function () {
      el.sendBtn.disabled = false;
      el.chatInput.disabled = false;
      setSpinner(el.chatSpinner, false);
      el.chatInput.focus();
    });
  }

  /**
   * 清空对话历史与聊天区 DOM
   */
  function clearChat() {
    conversationHistory = [];
    _lastAIResponse = '';
    currentCursorContext = null;

    // 清空聊天 DOM
    el.chatMessages.innerHTML = '';
    // 恢复空状态提示
    var hint = document.createElement('div');
    hint.className = 'chat-empty-hint';
    hint.id = 'chatEmptyHint';
    hint.innerHTML = '&#128172; 输入问题，AI 将基于文档内容回答<br><small>选中文本后提问可精确定位上下文</small>';
    el.chatMessages.appendChild(hint);
    el.chatEmptyHint = hint;

    // 隐藏上下文提示条
    el.contextBar.classList.remove('visible');
    el.insertLastBtn.style.display = 'none';
    clearStatus(el.chatStatus);
  }

  /**
   * 更新光标上下文提示条
   *
   * @param {string} text - 段落文本（原始，将截断显示前 20 字）
   */
  function updateContextIndicator(text, styleDisplay) {
    if (!text) {
      el.contextBar.classList.remove('visible');
      return;
    }
    var display = text.slice(0, 16) + (text.length > 16 ? '...' : '');
    if (styleDisplay) display = '[' + styleDisplay + '] ' + display;
    el.contextText.textContent = display;
    el.contextBar.classList.add('visible');
  }

  /**
   * 清除光标定位上下文，切换回全文模式
   */
  function clearCursorContext() {
    currentCursorContext = null;
    el.contextBar.classList.remove('visible');
    clearStatus(el.chatStatus);
  }

  /**
   * 样式名 → 简短中文显示名
   */
  function mapStyleToDisplayName(styleName) {
    if (!styleName) return '';
    if (/^Heading 1$|^标题 1$/i.test(styleName)) return '标题1';
    if (/^Heading 2$|^标题 2$/i.test(styleName)) return '标题2';
    if (/^Heading 3$|^标题 3$/i.test(styleName)) return '标题3';
    if (/^Heading 4$|^标题 4$/i.test(styleName)) return '标题4';
    if (/^Heading 5$|^标题 5$/i.test(styleName)) return '标题5';
    if (/^Heading 6$|^标题 6$/i.test(styleName)) return '标题6';
    if (/^Normal$|^正文$/i.test(styleName)) return '正文';
    if (/^Caption$|^题注$/i.test(styleName)) return '题注';
    if (/^Title$/i.test(styleName) && !/Heading|标题/.test(styleName)) return '文档标题';
    if (/^Subtitle$|^副标题$/i.test(styleName)) return '副标题';
    if (/^TOC\s*\d$|^目录\s*\d$/i.test(styleName)) return '目录';
    if (/^List\s/i.test(styleName) || /列表/i.test(styleName)) return '列表';
    if (/^Table Normal$|^普通表格$/i.test(styleName)) return '表格';
    return styleName;
  }

  /* ═══════════════════════════════════════════════════════════
     光标位置感知 — SelectionChanged 事件处理
     ═══════════════════════════════════════════════════════════ */

  /**
   * 处理光标位置变更（核心逻辑，不含防抖）
   *
   * 安全策略：
   *   - try-catch 包裹整个 Word.run，防止页眉/页脚/表格等特殊区域抛异常
   *   - 空白段落不覆盖 currentCursorContext（保留上一次有效上下文）
   *   - 检测文档 URL 变化 → 自动清空对话历史
   */
  function handleSelectionChanged() {
    try {
      Word.run(function (context) {
        var doc = context.document;
        doc.load('url');

        var sel = doc.getSelection();
        var paras = sel.paragraphs;
        context.load(paras, 'items');

        return context.sync().then(function () {
          // 检测文档切换 → 清空对话历史
          if (_lastDocumentUrl && doc.url && doc.url !== _lastDocumentUrl) {
            conversationHistory = [];
            currentCursorContext = null;
            _lastAIResponse = '';
            el.insertLastBtn.style.display = 'none';
            console.log('OfficeAI: 检测到文档切换，已清空对话历史');
          }
          _lastDocumentUrl = doc.url;

          // 获取光标所在段落文本 + 样式名
          var items = paras.items;
          if (!items || items.length === 0) return;
          var text = (items[0].text || '').trim();
          if (text.length === 0) return; // 空白处保留上一次有效上下文

          // 读取段落样式名
          var styleName = '';
          try { styleName = items[0].style || ''; } catch (e) {}
          var displayStyle = mapStyleToDisplayName(styleName);

          currentCursorContext = text.slice(0, 500);
          updateContextIndicator(text, displayStyle);
        });
      }).catch(function (err) {
        console.warn('SelectionChanged handler suppressed:', err.message);
      });
    } catch (e) {
      console.warn('SelectionChanged outer error:', e);
    }
  }

  /**
   * 防抖包装的 SelectionChanged 处理器（300ms）
   *
   * 避免快速移动光标时频繁触发 Office.js API 调用。
   * 每次新事件到达时重置计时器，只有 300ms 内无新事件才真正执行。
   */
  function debouncedSelectionChangedHandler() {
    if (_selectionDebounceTimer) clearTimeout(_selectionDebounceTimer);
    _selectionDebounceTimer = setTimeout(handleSelectionChanged, SELECTION_DEBOUNCE_MS);
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
    if (setHeadingStyle(p, level)) {
      stats.headings++;
    } else {
      console.warn('executeSetHeading failed: unable to set heading style (level=' + level + ')');
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
      p.firstLineIndent = indentPt;
      p.spaceBefore = 0;
      p.spaceAfter = 0;
      p.lineSpacing = lineSpacingPt;
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
      p.firstLineIndent = 0;
      p.leftIndent = 28; // ~4字符缩进
      p.spaceBefore = 2;
      p.spaceAfter = 2;
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
      p.leftIndent = 36;
      p.rightIndent = 36;
      p.firstLineIndent = 0;
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

    // --- SelectionChanged 光标位置感知（仅 Word 宿主） ---
    if (Office.context.document && Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        debouncedSelectionChangedHandler,
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('OfficeAI: SelectionChanged handler registered');
          } else {
            console.warn('OfficeAI: SelectionChanged registration failed:', result.error);
          }
        }
      );
    }

    // --- 视图切换 ---
    el.openSettingsBtn.addEventListener('click', function () {
      showPage('settings');
    });

    el.backToMainBtn.addEventListener('click', function () {
      showPage('main');
    });

    el.openPresetsBtn.addEventListener('click', function () {
      showPage('presets');
    });

    el.backFromPresetsBtn.addEventListener('click', function () {
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

      // 强制刷新缓存
      _docTextCache.timestamp = 0;

      getDocumentText().then(function (text) {
        if (!text || text.trim().length === 0) {
          showStatus(el.mainStatus, 'info', '文档为空，请先输入内容。');
        } else {
          el.documentText.value = text;
          showStatus(el.mainStatus, 'success',
            '已读取全文（约 ' + text.length + ' 字符）。');
        }
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

        // AI 预设：通过聊天引擎发送
        var preset = AI_PRESET_PROMPTS[presetKey];
        if (!preset) return;

        var text = el.documentText.value.trim();
        // 空选区/无文档时尝试使用光标上下文或文档缓存
        if (!text && currentCursorContext) {
          text = currentCursorContext;
        }
        if (!text) {
          showStatus(el.presetsStatus, 'error', '请先读取文档、选中文本，或将光标移至目标段落。');
          return;
        }

        // 构建完整提示词（预设 prompt + 文档文本作为上下文）
        var fullPrompt = preset.prompt + '\n\n"""\n' + text + '\n"""';
        // 切换回主界面以显示聊天交互
        showPage('main');
        sendMessage(fullPrompt, preset.systemAddon, text);
      });
    });

    // --- 聊天：发送按钮 ---
    el.sendBtn.addEventListener('click', function () {
      var msg = el.chatInput.value.trim();
      if (!msg) {
        showStatus(el.chatStatus, 'error', '请输入问题或指令。');
        return;
      }
      el.chatInput.value = '';
      sendMessage(msg);
    });

    // --- 聊天：Enter 发送，Shift+Enter 换行 ---
    el.chatInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        el.sendBtn.click();
      }
    });

    // --- 聊天：清空对话 ---
    el.clearChatBtn.addEventListener('click', function () {
      clearChat();
    });

    // --- 上下文提示条：清除光标定位 ---
    el.clearContextBtn.addEventListener('click', function () {
      clearCursorContext();
      showStatus(el.chatStatus, 'info', '已切换为全文模式。');
    });

    // --- 聊天：插入最后一条 AI 回复到文档光标处 ---
    el.insertLastBtn.addEventListener('click', function () {
      if (!_lastAIResponse) return;

      getParagraphCount().then(function (beforeCount) {
        return Word.run(function (context) {
          var selection = context.document.getSelection();
          selection.insertHtml(_lastAIResponse, Word.InsertLocation.replace);
          return context.sync();
        }).then(function () {
          return normalizeInsertedParagraphs(beforeCount);
        }).then(function (applied) {
          showStatus(el.chatStatus, 'success',
            '已插入到文档光标处。' + (applied > 0 ? ' 已为 ' + applied + ' 个段落应用样式。' : ''));
        }).catch(function (err) {
          showStatus(el.chatStatus, 'error', '插入失败: ' + (err.message || '未知错误'));
        });
      });
    });

    // --- 替换选区（使用最后一条 AI 回复） ---
    el.replaceBtn.addEventListener('click', function () {
      if (!_lastAIResponse) {
        showStatus(el.chatStatus, 'error', '暂无 AI 回复可供替换。');
        return;
      }

      getParagraphCount().then(function (beforeCount) {
        return Word.run(function (context) {
          var selection = context.document.getSelection();
          selection.insertHtml(_lastAIResponse, Word.InsertLocation.replace);
          return context.sync();
        }).then(function () {
          // 替换后全文归一化（替换位置不确定，全文扫描稳妥）
          return normalizeInsertedParagraphs(0);
        }).then(function (applied) {
          showStatus(el.chatStatus, 'success',
            '已替换选区。' + (applied > 0 ? ' 已为 ' + applied + ' 个段落应用样式。' : ''));
        }).catch(function (err) {
          showStatus(el.chatStatus, 'error', '替换失败: ' + (err.message || '未知错误'));
        });
      });
    });

    // --- 追加到文末（使用最后一条 AI 回复） ---
    el.appendBtn.addEventListener('click', function () {
      if (!_lastAIResponse) {
        showStatus(el.chatStatus, 'error', '暂无 AI 回复可供追加。');
        return;
      }

      getParagraphCount().then(function (beforeCount) {
        return Word.run(function (context) {
          var body = context.document.body;
          body.insertHtml(_lastAIResponse, Word.InsertLocation.end);
          return context.sync();
        }).then(function () {
          return normalizeInsertedParagraphs(beforeCount);
        }).then(function (applied) {
          showStatus(el.chatStatus, 'success',
            '已追加到文档末尾。' + (applied > 0 ? ' 已为 ' + applied + ' 个段落应用样式。' : ''));
        }).catch(function (err) {
          showStatus(el.chatStatus, 'error', '追加失败: ' + (err.message || '未知错误'));
        });
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
      console.log('OfficeAI v2.4: Word host detected');

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
      clearStatus(el.presetsStatus);

      console.log('OfficeAI v2.4: Initialization complete');
    } else {
      console.warn('OfficeAI: Unsupported host:', info.host);
    }
  });

})();
