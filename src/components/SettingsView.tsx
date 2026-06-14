import React, { useState } from 'react';
import { Eye, EyeOff, Save, ShieldAlert, Key, Settings, Sliders, FileText, CheckCircle2 } from 'lucide-react';
import { AddInSettings, AIModel } from '../types';

interface SettingsViewProps {
  settings: AddInSettings;
  onSaveSettings: (settings: AddInSettings) => void;
  serverHealth: {
    hasGeminiKey: boolean;
    hasDeepseekKey: boolean;
  };
}

export default function SettingsView({ settings, onSaveSettings, serverHealth }: SettingsViewProps) {
  const [apiKey, setApiKey] = useState(settings.deepseekApiKey);
  const [model, setModel] = useState<AIModel>(settings.model);
  const [temperature, setTemperature] = useState(settings.temperature);
  const [systemPrompt, setSystemPrompt] = useState(settings.systemPrompt);
  const [autoInsert, setAutoInsert] = useState(settings.autoInsertOnPolish);
  const [langFrom, setLangFrom] = useState(settings.languagePair.from);
  const [langTo, setLangTo] = useState(settings.languagePair.to);
  const [showKey, setShowKey] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // New states for custom models
  const [customModelName, setCustomModelName] = useState(settings.customModelName || '');
  const [customBaseUrl, setCustomBaseUrl] = useState(settings.customBaseUrl || '');

  const handleSave = () => {
    const updated: AddInSettings = {
      deepseekApiKey: apiKey,
      model,
      temperature,
      systemPrompt,
      autoInsertOnPolish: autoInsert,
      languagePair: {
        from: langFrom,
        to: langTo
      },
      customModelName,
      customBaseUrl
    };
    onSaveSettings(updated);
    setSaveSuccess(true);
    setTimeout(() => setSaveSuccess(false), 2500);
  };

  return (
    <div id="settings-panel" class="p-4 space-y-5 overflow-y-auto max-h-[85vh] animate-fade-in">
      {/* Title */}
      <div class="flex items-center gap-2 pb-2 border-b border-slate-100 dark:border-slate-800">
        <Settings class="h-5 w-5 text-indigo-500" />
        <h2 class="text-base font-semibold text-slate-800 dark:text-slate-100">AI 助理设置</h2>
      </div>

      {/* Model Selection */}
      <div class="space-y-2">
        <label class="block text-xs font-semibold text-slate-600 dark:text-slate-400">默认 AI 模型</label>
        <div class="grid grid-cols-1 gap-2">
          {/* DeepSeek V3 */}
          <button
            type="button"
            onClick={() => setModel('deepseek-chat')}
            class={`flex flex-col items-start p-3 rounded-lg border text-left transition-all ${
              model === 'deepseek-chat'
                ? 'border-indigo-500 bg-indigo-50/50 dark:bg-indigo-950/20 text-slate-800 dark:text-indigo-200'
                : 'border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-850'
            }`}
          >
            <div class="flex items-center justify-between w-full">
              <span class="text-sm font-medium">DeepSeek Chat (V3)</span>
              {model === 'deepseek-chat' && <span class="h-2 w-2 rounded-full bg-indigo-500" />}
            </div>
            <span class="text-xs text-slate-500 dark:text-slate-400 mt-1">适用于日常润色、翻译、行文创作，高效快捷。</span>
          </button>

          {/* DeepSeek R1 */}
          <button
            type="button"
            onClick={() => {
              setModel('deepseek-reasoner');
              setTemperature(0.7); // standard default
            }}
            class={`flex flex-col items-start p-3 rounded-lg border text-left transition-all ${
              model === 'deepseek-reasoner'
                ? 'border-indigo-500 bg-indigo-50/50 dark:bg-indigo-950/20 text-slate-800 dark:text-indigo-200'
                : 'border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-850'
            }`}
          >
            <div class="flex items-center justify-between w-full">
              <span class="text-sm font-medium">DeepSeek Reasoner (R1)</span>
              {model === 'deepseek-reasoner' && <span class="h-2 w-2 rounded-full bg-indigo-500" />}
            </div>
            <span class="text-xs text-slate-500 dark:text-slate-400 mt-1">深度推理模型。具备独立思考链路，可展示思考气泡，极佳的科研和论文撰写选择。</span>
          </button>

          {/* Gemini Flash */}
          <button
            type="button"
            onClick={() => setModel('gemini-3.5-flash')}
            class={`flex flex-col items-start p-3 rounded-lg border text-left transition-all ${
              model === 'gemini-3.5-flash'
                ? 'border-indigo-500 bg-indigo-50/50 dark:bg-indigo-950/20 text-slate-800 dark:text-indigo-200'
                : 'border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-850'
            }`}
          >
            <div class="flex items-center justify-between w-full">
              <span class="text-sm font-medium">Gemini 3.5 Flash</span>
              {model === 'gemini-3.5-flash' && <span class="h-2 w-2 rounded-full bg-indigo-500" />}
            </div>
            <span class="text-xs text-slate-500 dark:text-slate-400 mt-1">原生开箱即用支持（无需额外秘钥），速度极快，生成稳定。</span>
          </button>

          {/* Custom Model */}
          <button
            type="button"
            onClick={() => setModel('custom-model')}
            class={`flex flex-col items-start p-3 rounded-lg border text-left transition-all ${
              model === 'custom-model'
                ? 'border-indigo-500 bg-indigo-50/50 dark:bg-indigo-950/20 text-slate-800 dark:text-indigo-200'
                : 'border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-850'
            }`}
          >
            <div class="flex items-center justify-between w-full">
              <span class="text-sm font-medium">自定义模型 / 其它 API (Custom)</span>
              {model === 'custom-model' && <span class="h-2 w-2 rounded-full bg-indigo-500" />}
            </div>
            <span class="text-xs text-slate-500 dark:text-slate-400 mt-1">自主配置任意 OpenAI 兼容的第三方服务（如 Qwen、Llama 等）。</span>
          </button>
        </div>

        {/* Custom API Model fields detail */}
        {model === 'custom-model' && (
          <div class="p-3 bg-indigo-50/30 dark:bg-indigo-950/10 rounded-lg border border-indigo-100/50 dark:border-indigo-900/30 space-y-3 mt-2 animate-fade-in text-xs">
            <div class="space-y-1">
              <label class="block text-[11px] font-semibold text-slate-600 dark:text-slate-400">
                API Base URL (API 基准地址)
              </label>
              <input
                type="text"
                value={customBaseUrl}
                onChange={(e) => setCustomBaseUrl(e.target.value)}
                placeholder="例如: https://api.siliconflow.cn/v1"
                class="w-full text-xs p-2 border border-slate-200 dark:border-slate-800 rounded bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 font-mono"
              />
              <span class="text-[10px] text-slate-400 dark:text-slate-500 block">
                留空则默认调用 DeepSeek 官方接口。可搭配硅基流动、OpenRouter、本地 Ollama 或是中转通道配置。
              </span>
            </div>

            <div class="space-y-1">
              <label class="block text-[11px] font-semibold text-slate-600 dark:text-slate-400">
                Model Name (模型名称/标识)
              </label>
              <input
                type="text"
                value={customModelName}
                onChange={(e) => setCustomModelName(e.target.value)}
                placeholder="例如: deepseek-ai/DeepSeek-V3"
                class="w-full text-xs p-2 border border-slate-200 dark:border-slate-800 rounded bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 font-mono"
              />
              <span class="text-[10px] text-slate-400 dark:text-slate-500 block">
                您所使用的提供商平台专属模型 ID 标识。
              </span>
            </div>
          </div>
        )}
      </div>

      {/* API Key Configuration */}
      <div class="space-y-2 border-t border-slate-100 dark:border-slate-800 pt-3">
        <label class="flex items-center gap-1.5 text-xs font-semibold text-slate-600 dark:text-slate-400">
          <Key class="h-3.5 w-3.5" />
          DeepSeek API 密钥
        </label>
        <div class="relative">
          <input
            type={showKey ? 'text' : 'password'}
            value={apiKey}
            onChange={(e) => setApiKey(e.target.value)}
            placeholder={
              serverHealth.hasDeepseekKey 
                ? '已读取到服务器部署的 Key，非开发可留空' 
                : '请输入您的 sk-xxxxx DeepSeek 开放平台密钥'
            }
            class="w-full text-xs p-2.5 pr-10 border border-slate-200 dark:border-slate-800 rounded-lg bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 font-mono transition-colors"
          />
          <button
            type="button"
            onClick={() => setShowKey(!showKey)}
            class="absolute right-2.5 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 transition-colors"
          >
            {showKey ? <EyeOff class="h-4 w-4" /> : <Eye class="h-4 w-4" />}
          </button>
        </div>
        
        {/* Helper Badge status */}
        <div class="flex flex-col gap-1.5 pt-1">
          {serverHealth.hasDeepseekKey && (
            <div class="flex items-center gap-1.5 text-[11px] text-emerald-600 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/20 p-2 rounded">
              <CheckCircle2 class="h-3.5 w-3.5 flex-shrink-0" />
              <span>本系统后端已配置全局 DeepSeek API Key，您可免填直接使用。如有个人额度需使用，亦可在上侧输入。</span>
            </div>
          )}
          {!serverHealth.hasDeepseekKey && !apiKey && (
            <div class="flex items-start gap-1.5 text-[11px] text-amber-600 dark:text-amber-400 bg-amber-50 dark:bg-amber-950/20 p-2 rounded">
              <ShieldAlert class="h-3.5 w-3.5 mt-0.5 flex-shrink-0" />
              <span>检测到云端服务器未内置 DeepSeek Key。如需调用 DeepSeek，请在上侧输入您在开放平台申请的 DeepSeek API Key 或使用 Gemini Flash。</span>
            </div>
          )}
        </div>
      </div>

      {/* Model Parameters */}
      <div class="space-y-3 border-t border-slate-100 dark:border-slate-800 pt-3">
        <label class="flex items-center gap-1.5 text-xs font-semibold text-slate-600 dark:text-slate-400">
          <Sliders class="h-3.5 w-3.5" />
          生成参数设定
        </label>

        {/* Temperature */}
        <div class="space-y-1">
          <div class="flex items-center justify-between text-xs text-slate-500">
            <span>温度 (Temperature): {temperature}</span>
            <span class="text-[10px] bg-slate-100 dark:bg-slate-800 px-1.5 py-0.5 rounded text-slate-600 dark:text-slate-400">
              {temperature <= 0.3 ? '严谨学术' : temperature <= 0.7 ? '平衡平衡' : '富有创意'}
            </span>
          </div>
          <input
            type="range"
            min="0"
            max="1.2"
            step="0.1"
            value={temperature}
            disabled={model === 'deepseek-reasoner'}
            onChange={(e) => setTemperature(parseFloat(e.target.value))}
            class="w-full h-1 bg-slate-200 dark:bg-slate-800 rounded-lg appearance-none cursor-pointer disabled:opacity-50"
          />
          {model === 'deepseek-reasoner' && (
            <span class="text-[10px] text-slate-400 block">注：DeepSeek R1 深度思考推理模型不支持调整温度值，系统会自动以官方默认的最优参数请求。</span>
          )}
        </div>

        {/* System Prompt */}
        <div class="space-y-1 pt-1">
          <label class="block text-xs text-slate-500">全局指令 (System Prompt Override)</label>
          <textarea
            value={systemPrompt}
            onChange={(e) => setSystemPrompt(e.target.value)}
            rows={2}
            placeholder="例如：你是一个专业的高校博导、资深公文撰写专家。"
            class="w-full text-xs p-2 border border-slate-200 dark:border-slate-800 rounded-lg bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500"
          />
        </div>
      </div>

      {/* Workflow Options */}
      <div class="space-y-3 border-t border-slate-100 dark:border-slate-800 pt-3">
        <label class="flex items-center gap-1.5 text-xs font-semibold text-slate-600 dark:text-slate-400">
          <FileText class="h-3.5 w-3.5" />
          工作流预设
        </label>

        {/* Auto Insert Toggle */}
        <div class="flex items-center justify-between p-1.5 rounded bg-slate-50 dark:bg-slate-850">
          <div class="flex flex-col">
            <span class="text-xs font-medium text-slate-700 dark:text-slate-300">润色后直接插入替换</span>
            <span class="text-[10px] text-slate-500">快捷操作时，无需手动点击“替换”按钮</span>
          </div>
          <label class="relative inline-flex items-center cursor-pointer">
            <input
              type="checkbox"
              checked={autoInsert}
              onChange={(e) => setAutoInsert(e.target.checked)}
              class="sr-only peer"
            />
            <div class="w-8 h-4 bg-slate-200 dark:bg-slate-800 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-3 after:w-3 after:transition-all peer-checked:bg-indigo-500" />
          </label>
        </div>

        {/* Translation Language Selection */}
        <div class="grid grid-cols-2 gap-2">
          <div class="space-y-1">
            <span class="text-[10px] text-slate-500 block">快捷翻译 - 源语言</span>
            <select
              value={langFrom}
              onChange={(e) => setLangFrom(e.target.value)}
              class="w-full text-xs p-1.5 border border-slate-200 dark:border-slate-800 rounded bg-white dark:bg-slate-950 focus:outline-none"
            >
              <option value="中文">中文 (Chinese)</option>
              <option value="英语">英语 (English)</option>
              <option value="日语">日语 (Japanese)</option>
              <option value="韩语">韩语 (Korean)</option>
              <option value="德语">德语 (German)</option>
            </select>
          </div>
          <div class="space-y-1">
            <span class="text-[10px] text-slate-500 block">快捷翻译 - 目标语言</span>
            <select
              value={langTo}
              onChange={(e) => setLangTo(e.target.value)}
              class="w-full text-xs p-1.5 border border-slate-200 dark:border-slate-800 rounded bg-white dark:bg-slate-950 focus:outline-none"
            >
              <option value="英语">英语 (English)</option>
              <option value="中文">中文 (Chinese)</option>
              <option value="日语">日语 (Japanese)</option>
              <option value="韩语">韩语 (Korean)</option>
              <option value="德语">德语 (German)</option>
            </select>
          </div>
        </div>
      </div>

      {/* Save Button */}
      <div class="pt-4 border-t border-slate-100 dark:border-slate-800">
        <button
          type="button"
          onClick={handleSave}
          class="w-full flex items-center justify-center gap-1.5 py-2.5 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white font-medium text-xs rounded-lg transition-colors shadow-sm cursor-pointer"
        >
          <Save class="h-4 w-4" />
          {saveSuccess ? '已成功保存配置！' : '保存设置 & 应用修改'}
        </button>
      </div>
    </div>
  );
}
