import React, { useState } from 'react';
import { 
  Download, HelpCircle, FileCheck, Check, Settings, 
  ChevronRight, ExternalLink, RefreshCw, Layers, Terminal
} from 'lucide-react';

interface ManifestDownloaderProps {
  appUrl: string;
}

export default function ManifestDownloader({ appUrl }: ManifestDownloaderProps) {
  const [downloaded, setDownloaded] = useState(false);
  const [customOrigin, setCustomOrigin] = useState('');

  // Auto clean and determine deployment origin
  const getCleanOrigin = () => {
    if (customOrigin.trim() !== '') {
      let url = customOrigin.trim();
      if (url.endsWith('/')) {
        url = url.slice(0, -1);
      }
      return url;
    }
    return appUrl || (typeof window !== 'undefined' ? window.location.origin : 'https://word.msq.pub');
  };

  const deploymentOrigin = getCleanOrigin();

  // XML Manifest template that can be dynamically populated with current deployment origin
  const generateManifestXml = () => {
    return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
          xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
          xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
          xsi:type="TaskPaneApp">

  <!-- Basic Settings -->
  <Id>6b20e038-5f8e-4e3b-8481-976c5a9c0bc3</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>DeepSeek AI Coordinator</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="DeepSeek AI 写作助手" />
  <Description DefaultValue="接入 DeepSeek V3、R1 与 Gemini AI，直接在 Microsoft Word 中进行文本润色、中英翻译、智能续写和深度论文研讨。"/>

  <!-- Icon for your add-in -->
  <IconUrl DefaultValue="${deploymentOrigin}/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${deploymentOrigin}/assets/icon-64.png"/>

  <SupportUrl DefaultValue="${deploymentOrigin}" />

  <!-- Allowed domains for navigation -->
  <AppDomains>
    <AppDomain>${deploymentOrigin}</AppDomain>
    <AppDomain>https://api.deepseek.com</AppDomain>
  </AppDomains>

  <!-- Hosts configuration -->
  <Hosts>
    <Host Name="Document" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${deploymentOrigin}/index.html" />
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <!-- Ribbon command bindings -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          
          <FunctionFile resid="Commands.Url" />

          <!-- Main Toolbar command item -->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16" />
                  <bt:Image size="32" resid="Icon.32x32" />
                  <bt:Image size="80" resid="Icon.80x80" />
                </Icon>

                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>

                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <!-- Asset resources mappings -->
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="${deploymentOrigin}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="${deploymentOrigin}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="${deploymentOrigin}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${deploymentOrigin}" />
        <bt:Url id="Commands.Url" DefaultValue="${deploymentOrigin}/index.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="${deploymentOrigin}/index.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="欢迎使用 DeepSeek AI 写作助手" />
        <bt:String id="CommandsGroup.Label" DefaultValue="AI 写作" />
        <bt:String id="TaskpaneButton.Label" DefaultValue="DeepSeek 助手" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="加载项配置成功！返回【开始】顶栏，点击【DeepSeek 助手】按钮，打开极速 AI 侧边栏。" />
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="加载 DeepSeek AI 侧边栏助手" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
  };

  const downloadManifest = () => {
    const xml = generateManifestXml();
    const blob = new Blob([xml], { type: 'text/xml;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'deepseek-word-add-in-manifest.xml';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    setDownloaded(true);
    setTimeout(() => setDownloaded(false), 3000);
  };

  return (
    <div id="manifest-downloader-panel" class="p-4 space-y-4 overflow-y-auto max-h-[85vh] animate-fade-in text-xs">
      
      {/* Title */}
      <div class="flex items-center gap-2 pb-2 border-b border-slate-100 dark:border-slate-800">
        <Layers class="h-5 w-5 text-indigo-500" />
        <h2 class="text-base font-semibold text-slate-800 dark:text-slate-100">导入真实 Word 文档指南</h2>
      </div>

      {/* FAQ Clarification Card */}
      <div class="p-3 bg-amber-50/60 dark:bg-amber-950/20 border border-amber-200/50 dark:border-amber-900/30 rounded-lg space-y-1.5">
        <span class="font-bold text-amber-800 dark:text-amber-300 block text-xs flex items-center gap-1.5">
          <HelpCircle class="h-4 w-4 text-amber-600 dark:text-amber-400" />
          <span>核心原理解析：为什么在 Word 里不能直接用/上传 GitHub 后仍跳转？</span>
        </span>
        <div class="text-slate-600 dark:text-slate-300 text-[11px] space-y-1 my-1 leading-relaxed">
          <p>
            <b>1. Office 加载项的本质是一个“专属网页浏览器”：</b> Word 本身不运行代码，而是通过读取 <code>manifest.xml</code> 文件中的 <code>&lt;SourceLocation&gt;</code> 网址，在侧边栏里的 iframe 容器中加载对应的网页。
          </p>
          <p>
            <b>2. 为什么下载原版 XML 进去会显示演示版？</b> 因为您直接下载的代码库根目录下的 <code>manifest.xml</code> 里，所有的访问链接都被硬编码指定为了本项目演示站点的 URL (<code>ais-dev-...asia-southeast1.run.app</code>)。所以即使您把它导入了 Word，Word 依然会去读取网页沙盒演示版（它带有左边的白板编辑器）。
          </p>
          <p>
            <b>3. 解决方案：</b> 首先将项目静态打包上传到您的 GitHub Pages（例如 <code>https://your-username.github.io/your-project</code>），然后在下方<b>“GitHub Pages 自动部署配置”</b>中输入相同的网址，点击最下方的<b>“下载专属 manifest.xml”</b>。此时系统会实时把 XML 中所有的地址替换成您专属的 GitHub 域名。<b>使用该新生成的 XML 再次导入 Word，Word 侧栏就会跑您自己的纯净功能页，不再会看到任何网页端侧板！</b>
          </p>
        </div>
      </div>

      <p class="text-slate-600 dark:text-slate-400 leading-relaxed">
        本程序是一个符合微软 Office 原生标准的加载项 (Add-in) 插件。当前网页是沙盒演示环境，通过以下几个极简步骤，您就可以<b>将其发布并部署到您自己的 GitHub 上并作为独立的真实 Word 加载项运行</b>：
      </p>

      {/* GitHub Pages/Private Host Customization Card */}
      <div class="p-3 bg-indigo-50/50 dark:bg-indigo-950/20 border border-indigo-100 dark:border-indigo-900/30 rounded-lg space-y-2 mt-1">
        <span class="font-bold text-indigo-950 dark:text-indigo-200 block text-xs flex items-center gap-1.5">
          <Settings class="h-3.5 w-3.5 animate-spin-slow text-indigo-500" />
          <span>GitHub Pages 自动部署配置（免代码修改）</span>
        </span>
        <p class="text-slate-500 dark:text-slate-400 text-[11px] leading-relaxed">
          如果您想打包项目上传到 GitHub 并使用 <b>GitHub Pages (纯静态)</b> 托管。只需在下方填入您的部署网址，系统将全自动编译烘焙出符合您专属域名的 XML 文件，<b>您无需亲自动手修改任何代码！</b>
        </p>
        <div class="space-y-1.5 pt-1">
          <label class="block text-[10px] font-semibold text-slate-600 dark:text-slate-400">
            自定义部署网址 (例如 GitHub Pages URL):
          </label>
          <input
            type="text"
            value={customOrigin}
            onChange={(e) => setCustomOrigin(e.target.value)}
            placeholder="例如: https://your-username.github.io/your-repo-name"
            class="w-full text-xs p-2 border border-slate-200 dark:border-slate-800 rounded bg-white dark:bg-slate-950 focus:outline-none focus:border-indigo-500 font-mono text-slate-800 dark:text-slate-200"
          />
        </div>
      </div>

      {/* Origin Info Box */}
      <div class="bg-slate-50 dark:bg-slate-950 p-2.5 rounded-lg border border-slate-100 dark:border-slate-800 space-y-1">
        <span class="text-[10px] text-slate-400 block font-mono">当前已烘焙的目标部署网址 (Origin):</span>
        <code class="text-[10.5px] font-mono text-indigo-600 dark:text-indigo-400 break-all">{deploymentOrigin}</code>
      </div>

      {/* Step 1: Download button */}
      <div class="space-y-2 pt-2">
        <span class="font-bold text-slate-700 dark:text-slate-300 block">第一步：获取专属 Manifest 清单配置文件</span>
        <button
          type="button"
          onClick={downloadManifest}
          class="w-full flex items-center justify-center gap-1.5 py-2.5 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white font-medium rounded-lg transition-colors shadow-sm cursor-pointer"
        >
          {downloaded ? (
            <>
              <FileCheck class="h-4 w-4" />
              <span>配置文件下载成功！</span>
            </>
          ) : (
            <>
              <Download class="h-4 w-4" />
              <span>下载专属 manifest.xml 清单清单</span>
            </>
          )}
        </button>
      </div>

      {/* Step 2: Install Guide */}
      <div class="space-y-3 pt-2">
        <span class="font-bold text-slate-700 dark:text-slate-300 block">第二步：将其 sideload 导入您的常用 Microsoft Word</span>
        
        {/* Sub-Methods Selection */}
        <div class="space-y-2 max-w-full">
          {/* Method A: Word Online (easiest, cross-platform) */}
          <div class="p-3 bg-white dark:bg-slate-950 border border-slate-100 dark:border-slate-805 rounded-lg space-y-1.5 shadow-sm">
            <span class="text-xs font-semibold text-slate-800 dark:text-slate-200 flex items-center justify-between">
              <span>方法一：在 Word 网页版（企业/个人 Office.com）中载入</span>
              <span class="text-[9px] bg-indigo-50 dark:bg-indigo-950/20 text-indigo-600 dark:text-indigo-400 px-1.5 rounded-sm">更简单/不限平台</span>
            </span>
            <ol class="list-decimal list-inside pl-1 text-[11px] text-slate-500 dark:text-slate-400 space-y-1 leading-relaxed">
              <li>使用浏览器进入网页登录：<a href="https://office.com" target="_blank" rel="noreferrer" class="text-indigo-500 hover:underline inline-flex items-center gap-0.5 font-semibold">Office.com<ExternalLink class="h-2.5 w-2.5" /></a>；</li>
              <li>新建或打开任意一份已保存的空 <b>Word 文档</b>；</li>
              <li>在顶部菜单栏，点击进入<b>【插入 (Insert)】</b>面板；</li>
              <li>选择<b>【Office 加载项 (Add-ins)】</b>或 <b>【获取加载项】</b>控制框；</li>
              <li>在弹出的对话框中，点击右上角的 <b>【上传我的加载项】</b>，选择刚才下载的 <code>deepseek-word-add-in-manifest.xml</code>；</li>
              <li>上传完成后，顶部导航栏的<b>【开始 (Home)】</b>面板右侧即会新增一个<b>【DeepSeek 助手】</b>绿色/蓝色按钮，点击即可开启您的满血 DeepSeek 侧栏！</li>
            </ol>
          </div>

          {/* Method B: Native Word client */}
          <div class="p-3 bg-white dark:bg-slate-950 border border-slate-100 dark:border-slate-805 rounded-lg space-y-1.5 shadow-sm">
            <span class="text-xs font-semibold text-slate-800 dark:text-slate-200 flex items-center justify-between">
              <span>方法二：在 Windows/Mac 本地 Word 客户端载入</span>
            </span>
            <ul class="list-disc list-inside pl-1 text-[11px] text-slate-500 dark:text-slate-400 space-y-1 leading-relaxed">
              <li><b>Windows:</b> 新建一个共享文件夹，放入该 XML 文件。在 Word 的“信任中心”-“受信任的加载项目录”中添加此文件夹的路径，即可在 Word 里加载该插件。</li>
              <li><b>macOS:</b> 将下载的 XML 文件移动到指定目录：
                <code class="block font-mono text-[9.5px]/relaxed bg-slate-50 dark:bg-slate-900 border border-slate-100 dark:border-slate-800 p-1 rounded my-1 select-all break-all">
                  ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
                </code>
                (若无 wef 文件夹可自行创建)，重启 Word 客户端即可在【我的加载项】中双击导入。
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
