import React, { useState, useRef, useEffect } from 'react';
import { 
  FileText, Bold, Italic, AlignLeft, AlignCenter, AlignRight, 
  HelpCircle, Check, Sparkles, Download, ArrowRight, Layers
} from 'lucide-react';
import { registerMockDocumentBridge } from '../utils/office-utils';

interface WordSimulatorProps {
  onSideloadDownload: () => void;
  appUrl: string;
}

export default function WordSimulator({ onSideloadDownload, appUrl }: WordSimulatorProps) {
  // Prepopulate with a realistic draft having typos, clunky phrases, and translational needs
  const initialText = `2026年智能化行业应用发展规划草案（草稿审查版）

一、发展现状
当前，我司及行业对于大语言模型（LLM）的结合应用呈现爆发式增长。然而在实际公文与学术撰写中，由于缺乏高品质大模型的深度集成，导致写作质量不一。We must find a way to resolve this issue as soon as possible. 为了更好地提高公职和科研人员的效率，亟需引入顶尖AI如DeepSeek作为日常工作的辅助工具。

二、面临瓶颈
目前存在以下明显的问题：
1. 别字与语病较多：在起草长篇大论时，拼写别字（例如：“消售部”写成“销售部”已算万幸，更有“深喉”写成“深厚”、“精练”写成“精炼”等笔误）严重影响公信度。
2. 遣词造句过于平庸大白话：公文及论文需要特定严谨格式与文风，而起草人员往往文笔偏向大白话，显得不专业。
3. 跨国协作翻译耗时：海外项目及科研文献的大量汉英互译过程极其烦琐，耗费了大量人工时间。

三、拟采取对策
我司决定全方位引入 DeepSeek V3 与 R1 (Reasoner) 满血大模型：
[此处需要续写更多关于引入大模型的实施路径与预期收益...]`;

  const [docContent, setDocContent] = useState(initialText);
  const [selectionText, setSelectionText] = useState('');
  const [selectionRange, setSelectionRange] = useState({ start: 0, end: 0 });
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Hook simulated bridge selection when running in web mode
  useEffect(() => {
    registerMockDocumentBridge(
      // Getter for current Selection
      () => {
        if (textareaRef.current) {
          const start = textareaRef.current.selectionStart;
          const end = textareaRef.current.selectionEnd;
          const selected = docContent.substring(start, end);
          return { text: selected, start, end };
        }
        return { text: selectionText, start: selectionRange.start, end: selectionRange.end };
      },
      // Writer/Setter for replacement
      (newText, location) => {
        setDocContent(prev => {
          if (!textareaRef.current) return prev + "\n" + newText;
          const start = textareaRef.current.selectionStart;
          const end = textareaRef.current.selectionEnd;
          
          if (location === 'replace') {
            const updated = prev.substring(0, start) + newText + prev.substring(end);
            // Adjust cursor position after text update
            setTimeout(() => {
              if (textareaRef.current) {
                textareaRef.current.focus();
                textareaRef.current.setSelectionRange(start, start + newText.length);
              }
            }, 50);
            return updated;
          } else if (location === 'after') {
            const updated = prev.substring(0, end) + newText + prev.substring(end);
            setTimeout(() => {
              if (textareaRef.current) {
                textareaRef.current.focus();
                textareaRef.current.setSelectionRange(end, end + newText.length);
              }
            }, 50);
            return updated;
          } else {
            const updated = prev.substring(0, start) + newText + prev.substring(start);
            setTimeout(() => {
              if (textareaRef.current) {
                textareaRef.current.focus();
                textareaRef.current.setSelectionRange(start, start + newText.length);
              }
            }, 50);
            return updated;
          }
        });
      }
    );
  }, [docContent, selectionText, selectionRange]);

  const handleTextareaSelect = () => {
    if (textareaRef.current) {
      const start = textareaRef.current.selectionStart;
      const end = textareaRef.current.selectionEnd;
      const selected = docContent.substring(start, end);
      setSelectionText(selected);
      setSelectionRange({ start, end });
    }
  };

  const selectSuggestedAWK = () => {
    if (textareaRef.current) {
      // Preselect "We must find a way to resolve this issue as soon as possible."
      const textToFind = "We must find a way to resolve this issue as soon as possible.";
      const index = docContent.indexOf(textToFind);
      if (index !== -1) {
        textareaRef.current.focus();
        textareaRef.current.setSelectionRange(index, index + textToFind.length);
        setSelectionText(textToFind);
        setSelectionRange({ start: index, end: index + textToFind.length });
      }
    }
  };

  const selectSuggestedAwkwardChinese = () => {
    if (textareaRef.current) {
      const textToFind = "拼写别字（例如：“消售部”写成“销售部”已算万幸，更有“深喉”写成“深厚”、“精练”写成“精炼”等笔误）严重影响公信度。";
      const index = docContent.indexOf(textToFind);
      if (index !== -1) {
        textareaRef.current.focus();
        textareaRef.current.setSelectionRange(index, index + textToFind.length);
        setSelectionText(textToFind);
        setSelectionRange({ start: index, end: index + textToFind.length });
      }
    }
  };

  return (
    <div class="flex-1 flex flex-col bg-slate-100 dark:bg-slate-950 border-r border-slate-200 dark:border-slate-800/80 h-full overflow-hidden">
      
      {/* MS Word simulated toolbar chrome */}
      <div class="p-3 bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 flex flex-col gap-2 shadow-sm shrink-0">
        <div class="flex items-center justify-between">
          <div class="flex items-center gap-2">
            <div class="p-1 px-1.5 bg-blue-600 rounded text-white font-serif text-[11px] font-bold">W</div>
            <span class="text-xs font-semibold text-slate-800 dark:text-slate-200 font-sans truncate max-w-[200px]">
              DeepSeek_智能化发展规划草稿.docx
            </span>
            <span class="text-[10px] bg-emerald-50 text-emerald-600 dark:bg-emerald-950/20 dark:text-emerald-400 font-medium px-2 py-0.5 rounded flex items-center gap-1">
              <span class="h-1 w-1 rounded-full bg-emerald-500 animate-pulse"></span>
              自动云同步
            </span>
          </div>

          <button
            type="button"
            onClick={onSideloadDownload}
            class="text-[11px] px-2.5 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 dark:bg-indigo-950/40 dark:text-indigo-300 rounded font-medium transition-all flex items-center gap-1 cursor-pointer"
          >
            <Download class="h-3 w-3" />
            下载 Word 清单清单 XML
          </button>
        </div>

        {/* Word toolbar tools */}
        <div class="flex items-center gap-3 border-t border-slate-100 dark:border-slate-800 pt-2 select-none">
          <div class="flex items-center gap-1 text-slate-400 dark:text-slate-600">
            <Bold class="h-3.5 w-3.5" />
            <Italic class="h-3.5 w-3.5" />
          </div>
          <div class="h-4 w-[1px] bg-slate-200 dark:bg-slate-800" />
          <div class="flex items-center gap-1 text-slate-400 dark:text-slate-600">
            <AlignLeft class="h-3.5 w-3.5" />
            <AlignCenter class="h-3.5 w-3.5" />
            <AlignRight class="h-3.5 w-3.5" />
          </div>
          <div class="h-4 w-[1px] bg-slate-200 dark:bg-slate-800" />
          {/* Quick instructions on simulator */}
          <span class="text-[10px] text-indigo-500 dark:text-indigo-400 flex items-center gap-1">
            <Sparkles class="h-3 w-3 animate-pulse" />
            演示指南：选中文字，再在右侧面板点击“一键处理”或“对话”
          </span>
        </div>
      </div>

      {/* A4 Document centered wrapping block */}
      <div class="flex-1 overflow-y-auto p-6 md:p-10 flex flex-col items-center">
        
        {/* Document quick-sampler helpers */}
        <div class="w-full max-w-2xl bg-indigo-50/50 dark:bg-indigo-950/10 border border-indigo-100/30 p-3 rounded-lg mb-4 text-xs space-y-1.5 shadow-sm text-slate-700 dark:text-slate-300">
          <p class="font-semibold flex items-center gap-1">
            💡 快速开始体验选区联动 (Highlight Text to Simulate):
          </p>
          <div class="flex flex-wrap gap-2 pt-1">
            <button
              type="button"
              onClick={selectSuggestedAWK}
              class="bg-white hover:bg-slate-50 dark:bg-slate-900 dark:hover:bg-slate-850 border border-slate-200 dark:border-slate-800 px-2.5 py-1 rounded-md text-[10.5px] text-slate-600 dark:text-slate-400 hover:text-indigo-600 transition-colors cursor-pointer flex items-center gap-1"
            >
              一键高亮英文行并翻译成中文 <ArrowRight class="h-3 w-3" />
            </button>
            <button
              type="button"
              onClick={selectSuggestedAwkwardChinese}
              class="bg-white hover:bg-slate-50 dark:bg-slate-900 dark:hover:bg-slate-850 border border-slate-200 dark:border-slate-800 px-2.5 py-1 rounded-md text-[10.5px] text-slate-600 dark:text-slate-400 hover:text-indigo-600 transition-colors cursor-pointer flex items-center gap-1"
            >
              一键高亮错笔病句并智能润色 <Sparkles class="h-3 w-3 text-amber-500" />
            </button>
          </div>
        </div>

        {/* Word A4 Page simulation */}
        <div class="w-full max-w-2xl aspect-[1/1.4] bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-xl rounded-md flex flex-col p-8 md:p-12 relative overflow-hidden">
          
          {/* Subtle page watermark */}
          <div class="absolute -right-16 -top-16 p-24 rounded-full border border-slate-100 dark:border-slate-800/40 pointer-events-none select-none flex items-center justify-center">
            <Layers class="h-10 w-10 text-slate-100 dark:text-slate-800/40" />
          </div>

          <textarea
            ref={textareaRef}
            value={docContent}
            onChange={(e) => setDocContent(e.target.value)}
            onMouseUp={handleTextareaSelect}
            onKeyUp={handleTextareaSelect}
            onSelect={handleTextareaSelect}
            onFocus={handleTextareaSelect}
            class="flex-1 w-full bg-transparent border-none outline-none font-sans text-xs md:text-sm leading-relaxed text-slate-800 dark:text-slate-200 resize-none font-medium h-full tracking-wide whitespace-pre-wrap select-text word-editor focus:ring-0 focus:outline-none"
            placeholder="在 Word 文档里写点内容吧..."
            spellCheck={false}
          />

          {/* Simulated footer */}
          <div class="mt-4 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-[10px] text-slate-400 dark:text-slate-500 font-mono select-none">
            <span>第 1 页，共 1 页</span>
            <span>字数: {docContent.replace(/\s+/g, '').length} 字</span>
            <span>选区: {selectionText ? `${selectionText.length} 字` : '未选中'}</span>
          </div>
        </div>
      </div>
    </div>
  );
}
