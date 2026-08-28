import { renderLibrary } from './pages/library.js';
import { mountFloatingBall } from './tools/floatingBall.js';
import { renderWizard } from './pages/wizard.js';
import { renderSettings } from './pages/settings.js';
import { renderSaveTemplate } from './pages/saveTemplate.js';
import { renderDecks } from './pages/decks.js';
import { renderDeckWizard } from './pages/deckWizard.js';
import { Api } from './api.js';
import { showToast } from './ui.js';
import { bindTooltips } from './lib/tooltip.js';
import { perfRecord, flushPerf } from './lib/perf.js';
import { confirmRouteLeave } from './lib/navigation.js';
import { applyLanguageFromConfig, startAutoTranslate, translateDom } from './lib/i18n.js';

// 全局兜底：任何脚本错误或未处理的异步错误都变成可见提示，避免「点了没反应」
window.addEventListener('error', (e) => {
  const msg = (e.error as Error | undefined)?.message || e.message || '未知错误';
  showToast(`脚本错误：${msg}`, 5000);
});
window.addEventListener('unhandledrejection', (e) => {
  const reason = e.reason as { message?: string } | undefined;
  showToast(`未处理错误：${reason?.message || String(e.reason)}`, 5000);
});

Office.onReady(async () => {
  const bootT0 = performance.now();
  // 带圈问号悬停气泡：事件委托只需绑定一次，重绘后的 .info-tip 自动生效
  bindTooltips();
  // 应用界面字体大小配置（CSS 变量，全局随动）
  try {
    const cfg = await Api.getConfig();
    applyLanguageFromConfig(cfg?.ui?.language);
    const size = cfg?.ui?.fontSize;
    if (size && size >= 10 && size <= 24) {
      document.documentElement.style.setProperty('--ui-font', size + 'px');
    }
  } catch { /* 后端未就绪时使用默认字体 */ }
  startAutoTranslate();

  const nav = document.querySelector('.topnav') as HTMLElement;
  nav.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest('button[data-route]');
    if (btn) location.hash = btn.getAttribute('data-route') || '#library';
  });

  let activeHash = location.hash || '#library';
  let restoringHash = false;
  async function route() {
    const hash = location.hash || '#library';
    if (restoringHash) {
      restoringHash = false;
      return;
    }
    if (!(await confirmRouteLeave(activeHash, hash))) {
      restoringHash = true;
      location.hash = activeHash;
      return;
    }
    const page = document.getElementById('page')!;
    const routes: Record<string, (c: HTMLElement) => Promise<void>> = {
      '#library': renderLibrary,
      '#wizard': renderWizard,
      '#settings': renderSettings,
      '#save': renderSaveTemplate,
      '#decks': renderDecks,
      '#deck-wizard': renderDeckWizard
    };
    document.querySelectorAll('.topnav button').forEach((b) => {
      (b as HTMLElement).classList.toggle('active', b.getAttribute('data-route') === hash);
    });
    const rT0 = performance.now();
    await (routes[hash] || renderLibrary)(page);
    translateDom(document.body);
    activeHash = hash;
    perfRecord('frontendPage', performance.now() - rT0, { page: hash });
  }

  window.addEventListener('hashchange', route);
  await route();

  // 小工具悬浮球（常驻，不随路由变化）
  mountFloatingBall();
  perfRecord('frontendStartup', performance.now() - bootT0);
  flushPerf();

  // 启动自检：比对前端注入版本与后端运行版本，不一致时提示重新部署（解决「代码已更新但跑的是旧 exe」）
  try {
    const banner = document.createElement('div');
    banner.id = 'version-banner';
    banner.className = 'version-banner';
    (document.querySelector('.topnav') || document.body).after(banner);
    const v = await Api.getAppVersion();
    if (v.ok && v.appVersion) {
      if (v.appVersion !== __APP_VERSION__) {
        banner.classList.add('warn');
        banner.innerHTML = `⚠ 版本不一致：前端 ${__APP_VERSION__} ≠ 后端 ${v.appVersion}。请重新运行 <b>安装.bat</b> 部署并完全退出重开 PowerPoint。<button class="close" title="关闭">✕</button>`;
      } else {
        // 版本一致：不再显示「✓ 版本一致」提示横幅（用户指定，避免每次打开都打扰）
        banner.remove();
      }
    } else {
      banner.classList.add('warn');
      banner.innerHTML = `⚠ 后端状态异常（${v.ok ? '未知版本' : '不可达'}）。请确认已运行 <b>安装.bat</b> 或 <b>启动插件.bat</b>。<button class="close" title="关闭">✕</button>`;
    }
    banner.querySelector('.close')?.addEventListener('click', () => banner.remove());
  } catch {
    // 后端未启动：静默（各页面自有错误提示），不阻塞任务窗格
  }
});
