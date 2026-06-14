import React, { useState, useRef, useEffect } from 'react';
import { 
  Send, User, Terminal, ChevronDown, ChevronUp, Copy, Check, 
  ArrowDownToLine, Trash2, HelpCircle, Sparkles, MessageSquareHeart
} from 'lucide-react';
import { Message, AIModel } from '../types';
import { insertText, getSelectedText } from '../utils/office-utils';

interface TaskPaneChatProps {
  messages: Message[];
  onSendMessage: (text: string) => void;
  onClearHistory: () => void;
  isLoading: boolean;
  activeModel: AIModel;
}

export default function TaskPaneChat({ 
  messages, 
  onSendMessage, 
  onClearHistory, 
  isLoading,
  activeModel 
}: TaskPaneChatProps) {
  const [inputValue, setInputValue] = useState('');
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [insertedId, setInsertedId] = useState<string | null>(null);
  const [expandedThinking, setExpandedThinking] = useState<{ [key: string]: boolean }>({});
  
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Auto scroll
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isLoading]);

  const handleSend = () => {
    if (!inputValue.trim() || isLoading) return;
    onSendMessage(inputValue.trim());
    setInputValue('');
  };

  const handleCopy = (id: string, text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const handleInsert = async (id: string, text: string) => {
    try {
      await insertText(text, 'replace');
      setInsertedId(id);
      setTimeout(() => setInsertedId(null), 2000);
    } catch (err) {
      alert("插入失败: " + err);
    }
  };

  const toggleThinking = (id: string) => {
    setExpandedThinking(prev => ({
      ...prev,
      [id]: !prev[id]
    }));
  };

  const handleQuickQuestionFromDoc = async () => {
    try {
      const selected = await getSelectedText();
      if (!selected || selected.trim() === "") {
        setInputValue("帮我帮我写一篇关于人工智能的论文大纲。");
      } else {
        setInputValue(`针对以下文档所选文本：\n"""\n${selected}\n"""\n\n请修改建议并扩写：`);
      }
    } catch {
      setInputValue("帮我润色一些当前的公文大纲。");
    }
  };

  return (
    <div id="chat-section-container" class="flex flex-col h-[78vh] bg-slate-50 dark:bg-slate-900 animate-fade-in">
      
      {/* Header Bar */}
      <div class="flex items-center justify-between p-2.5 bg-white dark:bg-slate-950 border-b border-slate-100 dark:border-slate-800">
        <div class="flex items-center gap-1.5 text-xs text-slate-500 dark:text-slate-400">
          <MessageSquareHeart class="h-4 w-4 text-indigo-500" />
          <span>正在对话: </span>
          <span class="font-mono bg-slate-100 dark:bg-slate-800 px-1.5 py-0.5 rounded text-[10px] text-slate-700 dark:text-slate-300">
            {activeModel === 'custom-model' ? '自定义模型 (Custom)' : activeModel === 'deepseek-chat' ? 'DeepSeek V3' : activeModel === 'deepseek-reasoner' ? 'DeepSeek R1 (Reasoner)' : activeModel}
          </span>
        </div>
        {messages.length > 0 && (
          <button
            type="button"
            onClick={onClearHistory}
            class="text-[11px] text-rose-500 hover:text-rose-600 flex items-center gap-1 hover:bg-rose-50 dark:hover:bg-rose-950/20 px-2 py-1 rounded transition-colors cursor-pointer"
          >
            <Trash2 class="h-3 w-3" />
            清空历史
          </button>
        )}
      </div>

      {/* Messages Feed */}
      <div class="flex-1 overflow-y-auto p-3 space-y-3 scroll-smooth">
        {messages.length === 0 ? (
          <div class="flex flex-col items-center justify-center h-full text-center p-4 space-y-3 dark:opacity-80">
            <Sparkles class="h-10 w-10 text-indigo-400/80 animate-pulse" />
            <div class="space-y-1">
              <p class="text-xs font-semibold text-slate-700 dark:text-slate-300">欢迎使用 AI 写作助手</p>
              <p class="text-[11px] text-slate-400 max-w-[200px] leading-relaxed">您可以直接向我提问，也可以选中文件内容后，让我帮你续写、改写、解释、校验等。</p>
            </div>
            
            {/* Shortcut helper links */}
            <div class="grid grid-cols-1 gap-1.5 w-full max-w-[220px] pt-2">
              <button
                type="button"
                onClick={handleQuickQuestionFromDoc}
                class="text-[11px] text-left p-2 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-slate-600 dark:text-slate-400 flex items-center justify-between cursor-pointer transition-colors"
              >
                <span>📃 联动并提问 Word 选区</span>
                <span class="text-indigo-500 text-[10px]">✨</span>
              </button>
              <button
                type="button"
                onClick={() => setInputValue("帮我把接下来的汇报要点拟定一份公文大纲。")}
                class="text-[11px] text-left p-2 border border-slate-200 dark:border-slate-800 hover:border-indigo-500 bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-850 rounded-lg text-slate-600 dark:text-slate-400 flex items-center justify-between cursor-pointer transition-colors"
              >
                <span>💼 撰写学术/公文工作大纲</span>
                <span class="text-indigo-500 text-[10px]">✏️</span>
              </button>
            </div>
          </div>
        ) : (
          messages.map((msg) => {
            const isAI = msg.role === 'assistant';
            const hasThinking = !!msg.thinking;
            const isMsgExpanded = expandedThinking[msg.id] ?? true;

            return (
              <div 
                key={msg.id} 
                class={`flex flex-col space-y-1.5 ${isAI ? 'items-start' : 'items-end'}`}
              >
                {/* Meta details */}
                <span class="text-[9px] text-slate-400 dark:text-slate-500 px-1 font-mono">
                  {isAI ? `DeepSeek AI / ${msg.timestamp}` : `文档用户 / ${msg.timestamp}`}
                </span>

                {/* Bubble Container */}
                <div class={`max-w-[92%] rounded-xl p-3 text-xs leading-relaxed ${
                  isAI
                    ? msg.isError
                      ? 'bg-rose-50 border border-rose-100 text-rose-700 dark:bg-rose-950/20 dark:border-rose-900/30'
                      : 'bg-white border border-slate-100 dark:border-slate-800 dark:bg-slate-950 text-slate-800 dark:text-slate-200 shadow-sm'
                    : 'bg-indigo-600 text-white dark:bg-indigo-600 shadow-sm'
                }`}>
                  
                  {/* Thinking Accoridon for R1 */}
                  {isAI && hasThinking && (
                    <div class="mb-2 border-b border-slate-100 dark:border-slate-800/80 pb-2">
                      <button
                        type="button"
                        onClick={() => toggleThinking(msg.id)}
                        class="flex items-center gap-1 text-[10px] text-indigo-500 dark:text-indigo-400 hover:text-indigo-600 bg-indigo-50/50 dark:bg-indigo-950/10 px-2 py-1 rounded font-medium cursor-pointer"
                      >
                        <Terminal class="h-3 w-3 animate-pulse" />
                        <span>{isMsgExpanded ? '隐藏 R1 思考轨迹' : '查看 R1 思考轨迹'}</span>
                        {isMsgExpanded ? <ChevronUp class="h-3 w-3" /> : <ChevronDown class="h-3 w-3" />}
                      </button>
                      
                      {isMsgExpanded && (
                        <div class="mt-1.5 p-2 bg-slate-50 dark:bg-slate-900 border border-slate-100 dark:border-slate-800/60 rounded text-[10.5px] text-slate-500 dark:text-slate-400 font-mono flex flex-col gap-1 italic whitespace-pre-wrap max-h-[140px] overflow-y-auto leading-normal">
                          {msg.thinking}
                        </div>
                      )}
                    </div>
                  )}

                  {/* Message main content */}
                  <div class="whitespace-pre-wrap font-sans select-text break-words">
                    {msg.content}
                  </div>

                  {/* Utility bar for AI responses */}
                  {isAI && !msg.isError && (
                    <div class="flex items-center gap-2 mt-2 pt-2 border-t border-slate-50 dark:border-slate-850">
                      
                      {/* Copy action */}
                      <button
                        type="button"
                        onClick={() => handleCopy(msg.id, msg.content)}
                        class="text-[10px] flex items-center gap-0.5 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
                        title="复制到剪贴板"
                      >
                        {copiedId === msg.id ? (
                          <span class="flex items-center gap-0.5 text-emerald-500">
                            <Check class="h-3 w-3" /> 已复制
                          </span>
                        ) : (
                          <>
                            <Copy class="h-3 w-3" /> 复制
                          </>
                        )}
                      </button>

                      {/* Write / Import back into active cursor */}
                      <button
                        type="button"
                        onClick={() => handleInsert(msg.id, msg.content)}
                        class="text-[10px] flex items-center gap-0.5 text-indigo-500 hover:text-indigo-600"
                        title="插入到 Word 光标处"
                      >
                        {insertedId === msg.id ? (
                          <span class="flex items-center gap-0.5 text-emerald-500">
                            <Check class="h-3 w-3" /> 已插入
                          </span>
                        ) : (
                          <>
                            <ArrowDownToLine class="h-3 w-3" /> 插入文档
                          </>
                        )}
                      </button>
                    </div>
                  )}
                </div>
              </div>
            );
          })
        )}
        
        {/* Assistant Processing Skeleton */}
        {isLoading && (
          <div class="flex flex-col space-y-1 items-start animate-pulse">
            <span class="text-[9px] text-slate-400 font-mono">DeepSeek AI 正在演算输入中...</span>
            <div class="max-w-[70%] bg-white dark:bg-slate-950 border border-slate-100 dark:border-slate-800 rounded-xl p-3 space-y-2 shadow-sm">
              <div class="h-2 bg-slate-200 dark:bg-slate-800 rounded w-20" />
              <div class="h-2.5 bg-slate-200 dark:bg-slate-800 rounded w-40" />
              <div class="h-2.5 bg-slate-200 dark:bg-slate-800 rounded w-32" />
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Footer Interface */}
      <div class="p-2.5 bg-white dark:bg-slate-950 border-t border-slate-100 dark:border-slate-800 space-y-2">
        
        {/* Dynamic short helpers */}
        <div class="flex select-none items-center gap-1.5 overflow-x-auto pb-1 text-[10px]">
          <button
            type="button"
            onClick={handleQuickQuestionFromDoc}
            class="flex-shrink-0 border border-slate-200 dark:border-slate-800 bg-slate-50 hover:bg-slate-100 dark:bg-slate-900 hover:dark:bg-slate-850 px-2 py-1 rounded text-slate-600 dark:text-slate-400 flex items-center gap-1 cursor-pointer"
          >
            📂 读取当前选区
          </button>
          <button
            type="button"
            onClick={() => setInputValue("把刚才回答的内容进行整理，写成一份严谨的公文小结。")}
            class="flex-shrink-0 border border-slate-200 dark:border-slate-800 bg-slate-50 hover:bg-slate-100 dark:bg-slate-900 hover:dark:bg-slate-850 px-2 py-1 rounded text-slate-600 dark:text-slate-400 flex items-center gap-1 cursor-pointer"
          >
            📰 整理为公文格式
          </button>
          <button
            type="button"
            onClick={() => setInputValue("将这段文字换个更亲和委婉的风格。")}
            class="flex-shrink-0 border border-slate-200 dark:border-slate-800 bg-slate-50 hover:bg-slate-100 dark:bg-slate-900 hover:dark:bg-slate-850 px-2 py-1 rounded text-slate-600 dark:text-slate-400 flex items-center gap-1 cursor-pointer"
          >
            🧸 转换委婉语气
          </button>
        </div>

        {/* Input Text Form */}
        <div class="relative flex items-center">
          <textarea
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleSend();
              }
            }}
            placeholder="在此向 DeepSeek 输入指令或提问..."
            rows={1}
            disabled={isLoading}
            class="w-full text-xs p-2.5 pr-10 border border-slate-200 dark:border-slate-800 rounded-lg bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 resize-none font-sans max-h-[80px]"
          />
          <button
            type="button"
            onClick={handleSend}
            disabled={isLoading || !inputValue.trim()}
            class="absolute right-1.5 top-1/2 -translate-y-1/2 p-1.5 bg-indigo-500 hover:bg-indigo-600 text-white rounded-md transition-colors disabled:opacity-40 cursor-pointer"
          >
            <Send class="h-3.5 w-3.5" />
          </button>
        </div>
      </div>
    </div>
  );
}
