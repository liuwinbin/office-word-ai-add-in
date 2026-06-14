import React, { useState } from 'react';
import { 
  Sparkles, Languages, FileJson, ArrowDownToLine, 
  RotateCcw, Clipboard, HelpCircle, ChevronRight, Wand2, RefreshCw
} from 'lucide-react';
import { getSelectedText, insertText } from '../utils/office-utils';
import { AddInSettings } from '../types';

interface DocumentToolsProps {
  settings: AddInSettings;
  onCallAI: (system: string, prompt: string, onProgress?: (text: string) => void) => Promise<string>;
}

export default function DocumentTools({ settings, onCallAI }: DocumentToolsProps) {
  const [selectedText, setSelectedText] = useState<string>('');
  const [isReadingSelection, setIsReadingSelection] = useState<boolean>(false);
  const [aiOutput, setAiOutput] = useState<string>('');
  const [thinkingOutput, setThinkingOutput] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const [customPrompt, setCustomPrompt] = useState<string>('');
  const [lastAction, setLastAction] = useState<string>('');
  const [infoMessage, setInfoMessage] = useState<{ type: 'error' | 'success', text: string } | null>(null);

  const showNotification = (type: 'error' | 'success', text: string) => {
    setInfoMessage({ type, text });
    setTimeout(() => setInfoMessage(null), 4000);
  };

  const handleFetchSelection = async () => {
    setIsReadingSelection(true);
    try {
      const text = await getSelectedText();
      setSelectedText(text);
      if (!text || text.trim() === "") {
        showNotification('error', '未检测到文档选区。请先在 Word 文档里用鼠标选中一些文字。');
      } else {
        showNotification('success', '成功同步当前的文档选中内容！');
      }
    } catch (err: any) {
      showNotification('error', err.message || '读取选区失败');
    } finally {
      setIsReadingSelection(false);
    }
  };

  const executeAction = async (actionType: 'polish' | 'translate' | 'summarize' | 'continue' | 'custom') => {
    setIsLoading(true);
    setAiOutput('');
    setThinkingOutput('');
    setLastAction(actionType);
    
    try {
      // 1. Fetch selection first
      setProcessingStatus('正在读取文档选中文字...');
      const targetText = await getSelectedText();
      setSelectedText(targetText);
      
      if (!targetText || targetText.trim() === "") {
        throw new Error("请先用鼠标选中需要处理的文字！");
      }

      let systemPrompt = settings.systemPrompt || "你是一个集成于 Word 的高效文档助手。";
      let prompt = "";

      // 2. Build prompt based on action
      if (actionType === 'polish') {
        setProcessingStatus('正在进行智能润色纠错...');
        systemPrompt += " 你的任务是润色用户的文字。保证不改变原文主旨，修正语法错误、别字、拼写，优化遣词造句，提升文字的逻辑流畅度与学术公文感，不要输出多余的解释。只返回修饰后的最终文本。";
        prompt = `请润色优化以下文字，只返回润色后的正文：\n\n"""\n${targetText}\n"""`;
      } else if (actionType === 'translate') {
        setProcessingStatus(`正在将文字翻译为 ${settings.languagePair.to}...`);
        systemPrompt += ` 你的任务是翻译文字。请专业的将以下文字从【${settings.languagePair.from}】转换为【${settings.languagePair.to}】，力求信雅达。不要添加任何多余的解释，只返回翻译后的最终结果。`;
        prompt = `请将以下文本翻译为 ${settings.languagePair.to}：\n\n"""\n${targetText}\n"""`;
      } else if (actionType === 'summarize') {
        setProcessingStatus('正在智能生成提炼摘要...');
        systemPrompt += " 你的任务是提炼摘要。请使用清晰的信息列表或核心要点，对用户的文字进行高度凝练，保留核心观点与重点数据，格式整齐美观。";
        prompt = `请为以下文字生成凝练而条理清晰的摘要：\n\n"""\n${targetText}\n"""`;
      } else if (actionType === 'continue') {
        setProcessingStatus('正在智能续写下文...');
        systemPrompt += " 你是小说、文章、公文续写专家。请顺承以下文本的文风、语境和主旨，智能续写 150-300 字。只输出你需要续写的段落本身，不要包含前文和自我说明。";
        prompt = `请顺着以下文本的内容与风格继续往下书写：\n\n"""\n${targetText}\n"""`;
      } else if (actionType === 'custom') {
        if (!customPrompt.trim()) {
          throw new Error("请输入您的自定义修改指令！例如：将这段话改为文言文。");
        }
        setProcessingStatus('正在执行自定义修改指令...');
        systemPrompt += " 你是一个完全听从调配的文字改写专家。请严格按照用户特定的修改意图或写作指令，对输入文本进行针对性修饰重写，只返回处理完成的文本正文。";
        prompt = `修改指令：${customPrompt}\n\n需要处理的源文本：\n"""\n${targetText}\n"""`;
      }

      // 3. Call server AI endpoint
      const result = await onCallAI(systemPrompt, prompt, (partial) => {
        // Simple progress tracking
        setProcessingStatus('AI 正在为您全力生成中...');
      });

      setAiOutput(result);
      showNotification('success', 'AI 处理完成！可在下方预览结果。');

      // 4. Optionally auto-insert back to Word
      if (settings.autoInsertOnPolish) {
        setProcessingStatus('正在全自动向文档中填入内容...');
        await insertText(result, 'replace');
        showNotification('success', 'AI 已自动替换原文档中的选中区域！');
      }

    } catch (err: any) {
      console.error(err);
      showNotification('error', err.message || 'AI 操作执行失败');
    } finally {
      setIsLoading(false);
      setProcessingStatus('');
    }
  };

  const handleWriteToDoc = async (location: 'replace' | 'after' | 'before') => {
    if (!aiOutput) return;
    try {
      setProcessingStatus('正在将结果填入文档中...');
      await insertText(aiOutput, location);
      showNotification('success', '成功插入到 Word 文档！');
    } catch (err: any) {
      showNotification('error', err.message || '无法写入到文档');
    } finally {
      setProcessingStatus('');
    }
  };

  const handleCopyToClipboard = () => {
    if (!aiOutput) return;
    navigator.clipboard.writeText(aiOutput);
    showNotification('success', '已复制到剪贴板！');
  };

  return (
    <div id="quick-tools-container" class="p-4 space-y-4 overflow-y-auto max-h-[85vh] animate-fade-in">
      
      {/* Dynamic Notifications */}
      {infoMessage && (
        <div class={`text-[11px] p-2.5 rounded-lg border flex items-start gap-1.5 transition-all ${
          infoMessage.type === 'error'
            ? 'bg-rose-50 border-rose-100 text-rose-700 dark:bg-rose-950/20 dark:border-rose-900/30 dark:text-rose-300'
            : 'bg-emerald-50 border-emerald-100 text-emerald-700 dark:bg-emerald-950/20 dark:border-emerald-900/30 dark:text-emerald-300'
        }`}>
          <span class="font-bold flex-shrink-0">{infoMessage.type === 'error' ? '提示:' : '成功:'}</span>
          <span>{infoMessage.text}</span>
        </div>
      )}

      {/* Synchronizer Block */}
      <div class="bg-white dark:bg-slate-950 p-3 rounded-lg border border-slate-100 dark:border-slate-800 space-y-2 shadow-sm">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold text-slate-700 dark:text-slate-300 flex items-center gap-1">
            <span class="flex h-2 w-2 relative">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-2 w-2 bg-indigo-500"></span>
            </span>
            文档动态联动
          </span>
          <button 
            type="button"
            onClick={handleFetchSelection}
            disabled={isReadingSelection}
            class="text-[11px] flex items-center gap-1 px-2.5 py-1 text-indigo-600 hover:text-indigo-700 hover:bg-indigo-50 dark:text-indigo-400 dark:hover:bg-indigo-950/20 rounded-md font-medium transition-colors cursor-pointer"
          >
            <RefreshCw class={`h-3 w-3 ${isReadingSelection ? 'animate-spin' : ''}`} />
            同步所选
          </button>
        </div>
        
        <div class="text-xs bg-slate-50 dark:bg-slate-850 p-2.5 rounded border border-slate-100 dark:border-slate-800 text-slate-600 dark:text-slate-400 min-h-[44px] max-h-[80px] overflow-y-auto">
          {selectedText ? (
            <p class="font-mono text-[11px] leading-relaxed break-all whitespace-pre-wrap">{selectedText}</p>
          ) : (
            <span class="text-slate-400 italic">在 Word 中选中文字，或点击右侧“同步所选”拉取当前文字（网页端可直接在模拟器中写字并用鼠标选择）。</span>
          )}
        </div>
      </div>

      {/* Quick Action Matrix */}
      <div class="space-y-2">
        <span class="block text-xs font-semibold text-slate-600 dark:text-slate-400">一键文字智能处理</span>
        <div class="grid grid-cols-2 gap-2">
          
          {/* Polish Button */}
          <button
            type="button"
            onClick={() => executeAction('polish')}
            disabled={isLoading}
            class="flex items-center justify-center gap-1.5 py-2 px-3 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 dark:hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-xs font-medium text-slate-700 dark:text-slate-300 cursor-pointer transition-all disabled:opacity-50"
          >
            <Sparkles class="h-4 w-4 text-amber-500 flex-shrink-0" />
            <span>智能润色</span>
          </button>

          {/* Translate Button */}
          <button
            type="button"
            onClick={() => executeAction('translate')}
            disabled={isLoading}
            class="flex items-center justify-center gap-1.5 py-2 px-3 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 dark:hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-xs font-medium text-slate-700 dark:text-slate-300 cursor-pointer transition-all disabled:opacity-50"
          >
            <Languages class="h-4 w-4 text-indigo-500 flex-shrink-0" />
            <span>极速翻译</span>
          </button>

          {/* Summarize Button */}
          <button
            type="button"
            onClick={() => executeAction('summarize')}
            disabled={isLoading}
            class="flex items-center justify-center gap-1.5 py-2 px-3 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 dark:hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-xs font-medium text-slate-700 dark:text-slate-300 cursor-pointer transition-all disabled:opacity-50"
          >
            <FileJson class="h-4 w-4 text-emerald-500 flex-shrink-0" />
            <span>核心摘要</span>
          </button>

          {/* Continue Writing */}
          <button
            type="button"
            onClick={() => executeAction('continue')}
            disabled={isLoading}
            class="flex items-center justify-center gap-1.5 py-2 px-3 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 dark:hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-xs font-medium text-slate-700 dark:text-slate-300 cursor-pointer transition-all disabled:opacity-50"
          >
            <Wand2 class="h-4 w-4 text-violet-500 flex-shrink-0" />
            <span>行文续写</span>
          </button>
        </div>
      </div>

      {/* Free Custom Command prompt */}
      <div class="space-y-2 border-t border-slate-100 dark:border-slate-800 pt-3">
        <label class="block text-xs font-semibold text-slate-600 dark:text-slate-400">自定义修改意图</label>
        <div class="relative flex items-center">
          <input
            type="text"
            value={customPrompt}
            onChange={(e) => setCustomPrompt(e.target.value)}
            disabled={isLoading}
            placeholder="例如：缩写为更精简的版本、语调改为更委婉等"
            class="w-full text-xs p-2.5 pr-10 border border-slate-200 dark:border-slate-800 rounded-lg bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 disabled:opacity-50"
            onKeyDown={(e) => {
              if (e.key === 'Enter') executeAction('custom');
            }}
          />
          <button
            type="button"
            onClick={() => executeAction('custom')}
            disabled={isLoading || !customPrompt.trim()}
            class="absolute right-1.5 top-1/2 -translate-y-1/2 p-1.5 bg-indigo-500 hover:bg-indigo-600 text-white rounded-md transition-colors disabled:opacity-40 cursor-pointer"
          >
            <ChevronRight class="h-3.5 w-3.5" />
          </button>
        </div>
      </div>

      {/* Loading Block */}
      {isLoading && (
        <div class="bg-indigo-50/50 dark:bg-indigo-950/10 p-3 rounded-lg border border-indigo-100/30 flex flex-col items-center justify-center text-center py-5 space-y-2">
          <div class="h-6 w-6 border-2 border-t-transparent border-indigo-500 rounded-full animate-spin" />
          <span class="text-xs text-indigo-600 dark:text-indigo-400 font-medium">{processingStatus || "正在联系 AI 云脑..."}</span>
          <span class="text-[10px] text-slate-400">使用的是 {settings.model}</span>
        </div>
      )}

      {/* Output Console Console */}
      {aiOutput && !isLoading && (
        <div class="space-y-2 border-t border-slate-100 dark:border-slate-800 pt-3 animate-fade-in">
          <div class="flex items-center justify-between">
            <span class="text-xs font-semibold text-slate-600 dark:text-slate-400">处理结果预览</span>
            <div class="flex items-center gap-1">
              <button
                type="button"
                onClick={handleCopyToClipboard}
                class="p-1.5 text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
                title="复制到剪贴板"
              >
                <Clipboard class="h-3.5 w-3.5" />
              </button>
            </div>
          </div>

          <div class="p-3 bg-indigo-50/20 dark:bg-slate-900 border border-slate-100 dark:border-slate-800 rounded-lg text-xs leading-relaxed max-h-[160px] overflow-y-auto font-mono text-slate-800 dark:text-slate-200 break-all whitespace-pre-wrap">
            {aiOutput}
          </div>

          {/* Word Writing commands */}
          <div class="grid grid-cols-2 gap-2 text-[11px] pt-1.5">
            <button
              type="button"
              onClick={() => handleWriteToDoc('replace')}
              class="flex items-center justify-center gap-1 border border-indigo-500 hover:bg-indigo-500 hover:text-white dark:hover:bg-indigo-600 text-indigo-600 dark:text-indigo-400 font-semibold py-2 px-2 rounded-lg transition-all cursor-pointer"
            >
              <ArrowDownToLine class="h-3 w-3" />
              替换文档原文
            </button>
            <button
              type="button"
              onClick={() => handleWriteToDoc('after')}
              class="flex items-center justify-center gap-1 border border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-600 dark:text-indigo-200 font-semibold py-2 px-2 rounded-lg transition-all cursor-pointer"
            >
              <ChevronRight class="h-3 w-3" />
              插入到原文后
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
