import React, { useState, useEffect } from 'react';
import { 
  Sparkles, MessageSquare, Settings, HelpCircle, Laptop, 
  Wand2, Moon, Sun, Monitor, Download, ArrowUpRight
} from 'lucide-react';
import { AddInSettings, AIModel, Message, ChatSession } from './types';
import { isRunningInWord } from './utils/office-utils';
import SettingsView from './components/SettingsView';
import DocumentTools from './components/DocumentTools';
import TaskPaneChat from './components/TaskPaneChat';
import WordSimulator from './components/WordSimulator';
import ManifestDownloader from './components/ManifestDownloader';

// Default configuration settings
const DEFAULT_SETTINGS: AddInSettings = {
  deepseekApiKey: '',
  model: 'deepseek-chat',
  temperature: 0.7,
  systemPrompt: '你是一个专业的学术论文、公文、商务写作润色及校对专家。请提升行文流畅度与逻辑严密感。如无特定指示，请直接返回精炼后的修正文本本身，不做冗长说明。',
  autoInsertOnPolish: false,
  languagePair: {
    from: '中文',
    to: '英语'
  },
  customModelName: '',
  customBaseUrl: ''
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'tools' | 'chat' | 'settings' | 'manifest'>('tools');
  const [settings, setSettings] = useState<AddInSettings>(DEFAULT_SETTINGS);
  const [messages, setMessages] = useState<Message[]>([]);
  const [isLoadingChats, setIsLoadingChats] = useState(false);
  const [appUrl, setAppUrl] = useState('');
  
  // Environment detection state
  const [isWordHost, setIsWordHost] = useState(false);
  const [themeMode, setThemeMode] = useState<'light' | 'dark'>('light');

  // Backend server capabilities
  const [serverHealth, setServerHealth] = useState({
    hasGeminiKey: false,
    hasDeepseekKey: false
  });

  // Load configuration from local storage
  useEffect(() => {
    // 1. Check if running inside Word
    const detectionResult = isRunningInWord();
    setIsWordHost(detectionResult);
    
    // Fallback to active tab settings if in real Word (they might want to jump to settings directly if key is needed)
    if (detectionResult) {
      setActiveTab('tools');
    }

    // 2. Load Local Settings
    const saved = localStorage.getItem('deepseek_addin_settings');
    if (saved) {
      try {
        setSettings({ ...DEFAULT_SETTINGS, ...JSON.parse(saved) });
      } catch (e) {
        console.error("Failed to load local settings:", e);
      }
    }

    // 3. Load Local Chat History
    const savedChats = localStorage.getItem('deepseek_addin_chats_v1');
    if (savedChats) {
      try {
        setMessages(JSON.parse(savedChats));
      } catch (e) {
        console.error("Failed to load chats:", e);
      }
    }

    // 4. Query Server Health
    fetch('/api/health')
      .then(res => res.json())
      .then(data => {
        setServerHealth({
          hasGeminiKey: data.hasGeminiKey || false,
          hasDeepseekKey: data.hasDeepseekKey || false
        });
      })
      .catch(err => console.log("Health check skipped:", err));

    // Calculate simulated deployment host base URL
    setAppUrl(window.location.origin);

    // Bootstrap Office onReady hooks if running in Word
    if (typeof window !== 'undefined' && (window as any).Office) {
      (window as any).Office.onReady((info: any) => {
        if (info.host === (window as any).Office.HostType.Word) {
          setIsWordHost(true);
          console.log("Office.js is booted up inside Microsoft Word!");
        }
      });
    }
  }, []);

  // Save changes back to local storage
  const handleSaveSettings = (updated: AddInSettings) => {
    setSettings(updated);
    localStorage.setItem('deepseek_addin_settings', JSON.stringify(updated));
  };

  const persistMessages = (newMessages: Message[]) => {
    setMessages(newMessages);
    localStorage.setItem('deepseek_addin_chats_v1', JSON.stringify(newMessages));
  };

  const handleClearHistory = () => {
    persistMessages([]);
  };

  const performClientSideDirectChat = async (
    conversation: { role: string; content: string }[]
  ): Promise<{ content: string; reasoning?: string }> => {
    const isCustom = settings.model === 'custom-model';
    let targetUrl = 'https://api.deepseek.com/chat/completions';
    let modelName = 'deepseek-chat';

    if (isCustom) {
      modelName = settings.customModelName ? settings.customModelName.trim() : 'deepseek-chat';
      if (settings.customBaseUrl && settings.customBaseUrl.trim() !== '') {
        let cleanBase = settings.customBaseUrl.trim();
        if (cleanBase.endsWith('/')) {
          cleanBase = cleanBase.slice(0, -1);
        }
        targetUrl = cleanBase.includes('/chat/completions') ? cleanBase : `${cleanBase}/chat/completions`;
      }
    } else {
      modelName = settings.model === 'deepseek-reasoner' ? 'deepseek-reasoner' : 'deepseek-chat';
    }

    const apiKey = settings.deepseekApiKey;
    if (!apiKey || apiKey.trim() === '') {
      throw new Error('检测到属于独立静态部署（如 GitHub Pages），请先在“助手配置 (Settings)”中输入您的个人 API 密钥。');
    }

    const requestBody: any = {
      model: modelName,
      messages: conversation,
    };

    if (modelName !== 'deepseek-reasoner' && !modelName.includes('reasoner') && !modelName.includes('r1')) {
      requestBody.temperature = settings.temperature ?? 0.7;
    }

    const response = await fetch(targetUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey.trim()}`
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorText = await response.text();
      let errorData;
      try {
        errorData = JSON.parse(errorText);
      } catch (e) {
        errorData = { error: { message: errorText } };
      }
      throw new Error(errorData?.error?.message || errorData?.error || `直连 API 报错，状态码: ${response.status}`);
    }

    const data = await response.json();
    const choice = data.choices?.[0]?.message;
    if (!choice) {
      throw new Error('API 返回的数据格式不正确。');
    }

    return {
      content: choice.content || '',
      reasoning: choice.reasoning_content || choice.thinking || undefined
    };
  };

  // Helper function to invoke fullstack endpoint /api/chat
  const handleCallAI = async (
    systemDirective: string,
    prompt: string,
    onProgress?: (partial: string) => void
  ): Promise<string> => {
    setIsLoadingChats(true);
    try {
      const messagesPayload = [
        { role: 'system', content: systemDirective },
        { role: 'user', content: prompt }
      ];

      let useClientDirectFallback = false;
      let response;
      try {
        response = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            messages: messagesPayload,
            model: settings.model,
            temperature: settings.temperature,
            apiKey: settings.deepseekApiKey, // user's personal manual key override
            customModelName: settings.customModelName,
            customBaseUrl: settings.customBaseUrl
          })
        });
        if (response.status === 404) {
          useClientDirectFallback = true;
        }
      } catch (err) {
        useClientDirectFallback = true;
      }

      if (useClientDirectFallback) {
        const clientResult = await performClientSideDirectChat(messagesPayload);
        return clientResult.content;
      }

      if (!response || !response.ok) {
        const errJson = await response.json().catch(() => ({ error: '' }));
        throw new Error(errJson.error || `HTTP 错误码: ${response.status}`);
      }

      const data = await response.json();
      const choice = data.choices?.[0]?.message;
      if (!choice) {
        throw new Error("AI 接口返回的数据格式不正确。");
      }

      // Return both content and reasoning if exists
      return choice.content || '';
    } catch (err: any) {
      console.error(err);
      throw err;
    } finally {
      setIsLoadingChats(false);
    }
  };

  // Chat conversation runner, manages full stream session
  const handleSendChatMessage = async (text: string) => {
    const userMsg: Message = {
      id: Math.random().toString(36).substring(7),
      role: 'user',
      content: text,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    };

    const newMessages = [...messages, userMsg];
    persistMessages(newMessages);
    setIsLoadingChats(true);

    try {
      // Package conversation
      const conversation = [
        { role: 'system', content: settings.systemPrompt || "你是一个实用的 Word AI 智能助手。" },
        ...newMessages.map(m => ({ role: m.role, content: m.content }))
      ];

      let useClientDirectFallback = false;
      let response;
      try {
        response = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            messages: conversation,
            model: settings.model,
            temperature: settings.temperature,
            apiKey: settings.deepseekApiKey,
            customModelName: settings.customModelName,
            customBaseUrl: settings.customBaseUrl
          })
        });
        if (response.status === 404) {
          useClientDirectFallback = true;
        }
      } catch (err) {
        useClientDirectFallback = true;
      }

      if (useClientDirectFallback) {
        const clientResult = await performClientSideDirectChat(conversation);
        const aiMsg: Message = {
          id: Math.random().toString(36).substring(7),
          role: 'assistant',
          content: clientResult.content,
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          thinking: clientResult.reasoning
        };
        persistMessages([...newMessages, aiMsg]);
        return;
      }

      if (!response || !response.ok) {
        const errorJson = await response.json().catch(() => ({ error: '' }));
        throw new Error(errorJson.error || `API 错误, 状态码: ${response.status}`);
      }

      const data = await response.json();
      const answer = data.choices?.[0]?.message;
      if (!answer) {
        throw new Error("接口未成功返回文本内容。");
      }

      const aiMsg: Message = {
        id: Math.random().toString(36).substring(7),
        role: 'assistant',
        content: answer.content || '',
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        thinking: answer.reasoning_content || answer.thinking || undefined // Captures DeepSeek R1 reasoning content
      };

      persistMessages([...newMessages, aiMsg]);
    } catch (err: any) {
      const errorMsg: Message = {
        id: Math.random().toString(36).substring(7),
        role: 'assistant',
        content: `❌ 出错了: ${err.message || '网络连接失败，请检查设置。'}`,
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        isError: true
      };
      persistMessages([...newMessages, errorMsg]);
    } finally {
      setIsLoadingChats(false);
    }
  };

  const handleSideloadGuideClick = () => {
    setActiveTab('manifest');
  };

  // Render Sidebar Content based on active tab
  const renderTabContent = () => {
    switch (activeTab) {
      case 'tools':
        return <DocumentTools settings={settings} onCallAI={handleCallAI} />;
      case 'chat':
        return (
          <TaskPaneChat 
            messages={messages} 
            onSendMessage={handleSendChatMessage} 
            onClearHistory={handleClearHistory}
            isLoading={isLoadingChats}
            activeModel={settings.model}
          />
        );
      case 'settings':
        return (
          <SettingsView 
            settings={settings} 
            onSaveSettings={handleSaveSettings} 
            serverHealth={serverHealth}
          />
        );
      case 'manifest':
        return <ManifestDownloader appUrl={appUrl} />;
      default:
        return <DocumentTools settings={settings} onCallAI={handleCallAI} />;
    }
  };

  // Layout wrapper component
  // If running inside MS Word, show only the sidebar (TaskPane) with full-screen width.
  // If running standalone, show split viewport: Simulated Word Editor (left) + taskpane sidebar (right).
  return (
    <div className={`h-full flex flex-col font-sans select-none ${themeMode === 'dark' ? 'dark bg-slate-900' : 'bg-slate-50'}`}>
      
      {isWordHost ? (
        // Word Inside Sidebar Layout Mode
        <div class="flex-1 flex flex-col h-full overflow-hidden bg-white dark:bg-slate-950">
          {/* Compact Navbar for sidebar */}
          <div class="flex items-center justify-between px-3 py-2 border-b border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100">
            <div class="flex items-center gap-1.5 font-bold text-xs text-indigo-600 dark:text-indigo-400">
              <Sparkles class="h-4 w-4 animate-pulse text-amber-500" />
              <span>DeepSeek 写作助手</span>
            </div>
            
            {/* Dark Mode toggle icon for Word */}
            <button
              onClick={() => setThemeMode(themeMode === 'light' ? 'dark' : 'light')}
              class="p-1 rounded text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
              title="切换极色主题"
            >
              {themeMode === 'light' ? <Moon class="h-3.5 w-3.5" /> : <Sun class="h-3.5 w-3.5" />}
            </button>
          </div>

          {/* Sub Navigation controls tabs */}
          <div class="grid grid-cols-4 border-b border-slate-100 dark:border-slate-800 text-xs text-center select-none bg-slate-50/50 dark:bg-slate-900/50">
            <button
              type="button"
              onClick={() => setActiveTab('tools')}
              class={`py-2 px-1 font-semibold transition-colors border-b-2 flex flex-col items-center gap-0.5 cursor-pointer ${
                activeTab === 'tools'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'
              }`}
            >
              <span>⚡ 一键处理</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('chat')}
              class={`py-2 px-1 font-semibold transition-colors border-b-2 flex flex-col items-center gap-0.5 cursor-pointer ${
                activeTab === 'chat'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'
              }`}
            >
              <span>💬 自由提问</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('settings')}
              class={`py-2 px-1 font-semibold transition-colors border-b-2 flex flex-col items-center gap-0.5 cursor-pointer ${
                activeTab === 'settings'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'
              }`}
            >
              <span>⚙️ 助手配置</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('manifest')}
              class={`py-2 px-1 font-semibold transition-colors border-b-2 flex flex-col items-center gap-0.5 cursor-pointer ${
                activeTab === 'manifest'
                  ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                  : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300'
              }`}
            >
              <span>📖 本地部署</span>
            </button>
          </div>

          {/* Active Sub-panel page view */}
          <div class="flex-1 overflow-hidden bg-white dark:bg-slate-950">
            {renderTabContent()}
          </div>
        </div>
      ) : (
        // Web Showcase (Standalone Mock-editor split layout view)
        <div class="flex-1 flex flex-col h-full overflow-hidden">
          
          {/* Universal Header announcement bar */}
          <div class="bg-indigo-600 text-white py-2 px-4 flex items-center justify-between text-xs select-none shrink-0 font-sans shadow-sm">
            <div class="flex items-center gap-2">
              <span class="bg-indigo-700 text-indigo-100 font-bold px-1.5 py-0.5 rounded text-[10px]">网页演示环境 Web Showcase</span>
              <p class="truncate max-w-[400px] md:max-w-[600px]">
                🚀 您可以在本页左侧编辑草稿，选择文字体验 AI 润色替换。欲在您的真实 Microsoft Word 中加载本服务，请右键点击获取清单部署！
              </p>
            </div>
            <div class="flex items-center gap-3">
              <button 
                type="button"
                onClick={handleSideloadGuideClick}
                class="hover:underline flex items-center gap-1 font-semibold cursor-pointer"
              >
                导入真实 Word 教程 <ArrowUpRight class="h-3.5 w-3.5" />
              </button>
            </div>
          </div>

          {/* Split viewport */}
          <div class="flex-1 flex flex-row overflow-hidden bg-slate-50 dark:bg-slate-900">
            
            {/* Word Simulator (Left viewport) */}
            <WordSimulator onSideloadDownload={handleSideloadGuideClick} appUrl={appUrl} />

            {/* Simulated Word Panel Frame sidebar (Right viewport) */}
            <div class="w-[340px] md:w-[370px] shrink-0 bg-white dark:bg-slate-950 flex flex-col h-full border-l border-slate-200 dark:border-slate-800 shadow-xl overflow-hidden">
              
              {/* Header inside Panel representation */}
              <div class="p-3.5 border-b border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <div class="p-1 px-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded text-[10px] font-bold tracking-wider leading-none">
                    DEEPSEEK
                  </div>
                  <span class="text-xs font-bold text-slate-800 dark:text-slate-100 font-sans">
                    AI 写作助手
                  </span>
                </div>

                <div class="flex items-center gap-2">
                  {/* Theme toggling */}
                  <button
                    onClick={() => setThemeMode(themeMode === 'light' ? 'dark' : 'light')}
                    class="p-1.5 rounded-md hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400 hover:text-slate-600 transition-colors"
                  >
                    {themeMode === 'light' ? <Moon class="h-4 w-4" /> : <Sun class="h-4 w-4" />}
                  </button>
                </div>
              </div>

              {/* Sub-panels tabs inside frame */}
              <div class="grid grid-cols-4 border-b border-slate-100 dark:border-slate-800 text-xs text-center bg-slate-50/50 dark:bg-slate-900/50 select-none">
                <button
                  type="button"
                  onClick={() => setActiveTab('tools')}
                  class={`py-2 px-1 font-semibold transition-all border-b-2 cursor-pointer ${
                    activeTab === 'tools'
                      ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400 font-bold bg-white dark:bg-slate-950'
                      : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 hover:bg-slate-100/50'
                  }`}
                >
                  ⚡ 一键处理
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('chat')}
                  class={`py-2 px-1 font-semibold transition-all border-b-2 cursor-pointer ${
                    activeTab === 'chat'
                      ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400 font-bold bg-white dark:bg-slate-950'
                      : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 hover:bg-slate-100/50'
                  }`}
                >
                  💬 自由提问
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('settings')}
                  class={`py-2 px-1 font-semibold transition-all border-b-2 cursor-pointer ${
                    activeTab === 'settings'
                      ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400 font-bold bg-white dark:bg-slate-950'
                      : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 hover:bg-slate-100/50'
                  }`}
                >
                  ⚙️ 助手配置
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('manifest')}
                  class={`py-2 px-1 font-semibold transition-all border-b-2 cursor-pointer ${
                    activeTab === 'manifest'
                      ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400 font-bold bg-white dark:bg-slate-950'
                      : 'border-transparent text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 hover:bg-slate-100/50'
                  }`}
                >
                  📖 本地部署
                </button>
              </div>

              {/* Sidebar body page wrapper */}
              <div class="flex-1 overflow-hidden bg-white dark:bg-slate-955">
                {renderTabContent()}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
