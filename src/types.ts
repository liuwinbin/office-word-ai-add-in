export type AIModel = 'deepseek-chat' | 'deepseek-reasoner' | 'gemini-3.5-flash' | 'custom-model';

export interface Message {
  id: string;
  role: 'system' | 'user' | 'assistant';
  content: string;
  timestamp: string;
  thinking?: string; // DeepSeek R1 thinking content
  isError?: boolean;
}

export interface ChatSession {
  id: string;
  title: string;
  messages: Message[];
  model: AIModel | string;
  createdAt: string;
}

export interface AddInSettings {
  deepseekApiKey: string;
  model: AIModel;
  temperature: number;
  systemPrompt: string;
  autoInsertOnPolish: boolean;
  languagePair: {
    from: string;
    to: string;
  };
  customModelName: string; // E.g., deepseek-coder, qwen-max, or deepseek-ai/DeepSeek-V3
  customBaseUrl: string;   // E.g., https://api.siliconflow.cn/v1 or local Ollama
}

export interface PresetAction {
  id: string;
  name: string;
  icon: string;
  prompt: string;
  description: string;
  category: 'polish' | 'translate' | 'summarize' | 'custom';
}
