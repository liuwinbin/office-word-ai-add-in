/* global Word, Office */

// Ambient global declarations for MS Office JS loaded via CDN
declare const Office: any;
declare const Word: any;

// Store simulated callback bindings for browser local testing
let mockGetSelectionCallback: () => { text: string; start: number; end: number } = () => ({ text: "", start: 0, end: 0 });
let mockInsertTextCallback: (text: string, location: 'replace' | 'after' | 'before') => void = () => {};

/**
 * Configure standard mock event bridges when running outside of MS Word (e.g. standard browser preview)
 */
export function registerMockDocumentBridge(
  getSelection: () => { text: string; start: number; end: number },
  insertText: (text: string, location: 'replace' | 'after' | 'before') => void
) {
  mockGetSelectionCallback = getSelection;
  mockInsertTextCallback = insertText;
}

/**
 * Checks whether the application is running inside MS Word Office Add-in host environment
 */
export function isRunningInWord(): boolean {
  try {
    // 1. Direct query parameter check (Office always attaches ?_host_Info=Word$...)
    if (typeof window !== "undefined" && window.location.search) {
      const searchLower = window.location.search.toLowerCase();
      if (
        searchLower.includes("host_info=word") || 
        searchLower.includes("_host_info=word") ||
        searchLower.includes("host=word")
      ) {
        return true;
      }
    }
    // 2. Standard Office JS object model state check
    if (typeof Office !== "undefined" && Office.context) {
      const host = Office.context.host;
      if (host === "Word" || (Office.HostType && host === Office.HostType.Word)) {
        return true;
      }
    }
    return false;
  } catch {
    return false;
  }
}

/**
 * Safely fetches selection from the Word document or fallback simulated preview
 */
export async function getSelectedText(): Promise<string> {
  if (isRunningInWord()) {
    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        return selection.text || "";
      });
    } catch (error) {
      console.error("Word getSelectedText failed:", error);
      throw new Error("无法读取 Word 选区。请确保光标已在文档中选中内容。");
    }
  } else {
    // Return simulator mock selection
    return mockGetSelectionCallback().text;
  }
}

/**
 * Inserts content at selection (replaces selection, appends, or inserts before/after)
 */
export async function insertText(
  text: string,
  location: 'replace' | 'after' | 'before' = 'replace'
): Promise<void> {
  if (isRunningInWord()) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        
        // Map simplified locations to MS Word insert modes
        let wordLocation = Word.InsertLocation.replace;
        if (location === 'after') {
          wordLocation = Word.InsertLocation.after;
        } else if (location === 'before') {
          wordLocation = Word.InsertLocation.before;
        }
        
        selection.insertText(text, wordLocation);
        await context.sync();
      });
    } catch (error) {
      console.error("Word insertText failed:", error);
      throw new Error("无法将内容插入到 Word。请检查选区状态。");
    }
  } else {
    // Call simulator mock insert
    mockInsertTextCallback(text, location);
  }
}

/**
 * Performs deep polish operation directly inside the document.
 * Checks for selected text, requests processing, and writes back the polished output.
 */
export async function polishSelectedText(
  processor: (text: string) => Promise<string>,
  onStatusChange?: (status: string) => void
): Promise<void> {
  try {
    onStatusChange?.("正在读取文档选中文字...");
    const selected = await getSelectedText();
    
    if (!selected || selected.trim() === "") {
      throw new Error("未检测到选中的文字！请先用鼠标选中需要修改润色的文字。");
    }

    onStatusChange?.("正在调用 DeepSeek AI 处理中...");
    const polished = await processor(selected);

    onStatusChange?.("正在写入处理后的内容...");
    await insertText(polished, 'replace');
    
    onStatusChange?.("");
  } catch (error: any) {
    onStatusChange?.("");
    throw error;
  }
}
