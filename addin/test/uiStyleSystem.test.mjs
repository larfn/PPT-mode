import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const addinRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const css = readFileSync(resolve(addinRoot, 'src/styles.css'), 'utf8');
const saveTemplateSource = readFileSync(resolve(addinRoot, 'src/pages/saveTemplate.ts'), 'utf8');
const settingsSource = readFileSync(resolve(addinRoot, 'src/pages/settings.ts'), 'utf8');
const highlightSource = readFileSync(resolve(addinRoot, 'src/office/highlight.ts'), 'utf8');
const taskpaneSource = readFileSync(resolve(addinRoot, 'src/taskpane.ts'), 'utf8');
const writeSlideSource = readFileSync(resolve(addinRoot, 'src/office/writeSlide.ts'), 'utf8');
const decksSource = readFileSync(resolve(addinRoot, 'src/pages/decks.ts'), 'utf8');
const deckWizardSource = readFileSync(resolve(addinRoot, 'src/pages/deckWizard.ts'), 'utf8');
const apiSource = readFileSync(resolve(addinRoot, 'src/api.ts'), 'utf8');
const i18nSource = readFileSync(resolve(addinRoot, 'src/lib/i18n.ts'), 'utf8');

function ruleBody(selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = css.match(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`));
  return match?.[1] || '';
}

function assertDecl(selector, declaration) {
  assert.match(
    ruleBody(selector),
    new RegExp(declaration.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
    `${selector} should include "${declaration}"`,
  );
}

test('secondary, ghost, and icon buttons use visible unified backgrounds', () => {
  assertDecl('button.secondary', 'background: var(--control-bg)');
  assertDecl('button.ghost', 'background: var(--control-bg-muted)');
  assertDecl('button.locate, button.del', 'background: var(--control-bg-muted)');
  assertDecl('.preview-close', 'background: var(--control-bg-muted)');
  assertDecl('.tool-menu-item', 'background: var(--control-bg-muted)');
});

test('surfaces share unified radius, border, spacing, and shadow tokens', () => {
  assertDecl('.module', 'border-radius: var(--radius-card)');
  assertDecl('.module', 'box-shadow: var(--shadow-card)');
  assertDecl('.card', 'border-radius: var(--radius-card)');
  assertDecl('.modal-box', 'border-radius: var(--radius-modal)');
  assertDecl('.modal-box', 'box-shadow: var(--shadow-modal)');
  assertDecl('.src-row', 'border-radius: var(--radius-control)');
});

test('recommendation chips use compact labels and stable width', () => {
  assert.match(saveTemplateSource, /建议：/);
  assert.match(saveTemplateSource, /class="ghost rec-apply"[^>]*>√<\/button>/);
  assert.match(saveTemplateSource, /class="ghost rec-ignore"[^>]*>×<\/button>/);
  assert.doesNotMatch(saveTemplateSource, /规则\}建议|规则建议|Math\.round\(rec\.confidence|rec-conf|infoTip\(rec\.reason\)|>应用<\/button>|>忽略<\/button>/);
  assertDecl('.rec-line', 'width: 230px');
  assertDecl('.rec-line', 'justify-content: center');
  assertDecl('.rec-actions', 'margin-left: auto');
  assertDecl('.rec-actions', 'flex: 0 0 auto');
});

test('character range sliders move in ten-character increments', () => {
  assert.match(saveTemplateSource, /data-field="minChars"[^>]*step="10"/);
  assert.match(saveTemplateSource, /data-field="maxChars"[^>]*step="10"/);
  assert.match(saveTemplateSource, /Math\.round\(value \/ 10\) \* 10/);
});

test('highlight duration uses compact 0 to 0.5 second slider', () => {
  assert.match(settingsSource, /id="hl-duration" type="range" min="0" max="0\.5" step="0\.1"/);
  assert.match(settingsSource, /id="hl-duration-val"/);
  assert.match(settingsSource, /Math\.max\(0, Math\.min\(500,/);
  assert.doesNotMatch(settingsSource, /id="hl-duration" type="number"/);
  assert.doesNotMatch(highlightSource, /h\.durationMs >= 200/);
  assert.match(highlightSource, /Math\.max\(0, Math\.min\(500,/);
});

test('system diagnostics only shows user-facing essentials', () => {
  assert.match(settingsSource, /pushRow\('版本状态'/);
  assert.match(settingsSource, /pushRow\('后端连接', '正常', 'ok'\)/);
  assert.match(settingsSource, /pushRow\('MCP'/);
  assert.doesNotMatch(settingsSource, /pushRow\('前端资源'/);
  assert.doesNotMatch(settingsSource, /pushRow\('后端 exe'/);
  assert.doesNotMatch(settingsSource, /pushRow\('API 版本'/);
  assert.doesNotMatch(settingsSource, /pushRow\('运行端口'/);
  assert.doesNotMatch(settingsSource, /pushRow\('Runtime Token'/);
});

test('settings page exposes Chinese/English language switch and taskpane applies it', () => {
  assert.match(settingsSource, /id="ui-language"/);
  assert.match(settingsSource, /value="zh"[\s\S]*中文/);
  assert.match(settingsSource, /value="en"[\s\S]*English/);
  assert.match(settingsSource, /language: \(container\.querySelector\('#ui-language'\)/);
  assert.match(taskpaneSource, /applyLanguageFromConfig\(cfg\?\.ui\?\.language\)/);
  assert.match(taskpaneSource, /translateDom\(document\.body\)/);
  assert.match(i18nSource, /'模板库': 'Template Library'/);
  assert.match(i18nSource, /'AI 配置': 'AI Settings'/);
});

test('save template page asks before leaving after loading template state', () => {
  assert.match(taskpaneSource, /confirmRouteLeave\(activeHash, hash\)/);
  assert.match(saveTemplateSource, /setRouteLeaveGuard\('#save'/);
  assert.match(saveTemplateSource, /确认离开当前界面/);
  assert.match(saveTemplateSource, /离开后，本次已读取\/载入的模板信息不会保留/);
});

test('editing a saved template creates and cleans a marked temporary preview slide', () => {
  assert.match(saveTemplateSource, /TEMP_PREVIEW_NOTICE = '临时预览页：仅用于定位查看，离开或保存时会自动删除。'/);
  assert.match(saveTemplateSource, /font: 'Microsoft YaHei'/);
  assert.match(saveTemplateSource, /eaFont: '微软雅黑'/);
  assert.match(saveTemplateSource, /size: 15/);
  assert.match(saveTemplateSource, /color: '#FF0000'/);
  assert.match(saveTemplateSource, /insertTemporarySlideBase64/);
  assert.match(saveTemplateSource, /cleanupTempPreviewSlide/);
  assert.match(saveTemplateSource, /await deleteSlideById\(id\)/);
  assert.match(writeSlideSource, /export async function insertTemporarySlideBase64/);
  assert.match(writeSlideSource, /export async function deleteSlideById/);
});

test('temporary preview uses concise plain placeholder labels', () => {
  assert.match(saveTemplateSource, /const previewLabelFor = \(s: TemplateShape, index: number\): string =>/);
  assert.match(saveTemplateSource, /if \(s\.semanticRole\) return semanticRoleLabel\(s\.semanticRole\)/);
  assert.match(saveTemplateSource, /textsOut\[s\.id\] = previewLabelFor\(s, i\)/);
  assert.match(saveTemplateSource, /varsOut\[s\.id\] = previewLabelFor\(s, i\)/);
  assert.match(saveTemplateSource, /content: previewImageLabelFor\(s, i\)/);
  assert.doesNotMatch(saveTemplateSource, /textsOut\[s\.id\] = `\{\{/);
  assert.doesNotMatch(saveTemplateSource, /varsOut\[s\.id\] = `\{\{/);
  assert.doesNotMatch(saveTemplateSource, /content: `\{\{图片位：/);
});

test('locate action refreshes highlight color from settings before drawing', () => {
  assert.match(saveTemplateSource, /const refreshHighlightConfig = async \(\): Promise<void> =>/);
  assert.match(saveTemplateSource, /await refreshHighlightConfig\(\);[\s\S]*highlightShapeOnSlide\(s\.bounds, highlightColor, highlightDuration\)/);
});

test('deck builder uses template picker modal instead of inline template controls', () => {
  assert.match(decksSource, /openTemplatePicker/);
  assert.match(decksSource, /deck-picker-modal/);
  assert.match(decksSource, /deck-picker-grid/);
  assert.match(decksSource, /确认/);
  assert.match(decksSource, /取消/);
  assert.doesNotMatch(decksSource, /class="row-tpl"/);
  assert.doesNotMatch(decksSource, /class="row-ver"/);
  assert.doesNotMatch(decksSource, /图片搜索提示（可选）/);
  assert.doesNotMatch(decksSource, /文字生成指令（可选）/);
});

test('deck page presents visual sequence cards and compact deck actions', () => {
  assert.match(decksSource, /<button class="primary" id="deck-new"/);
  assert.doesNotMatch(decksSource, /<h3 style="margin-top:0">套版/);
  assert.doesNotMatch(decksSource, /点「新建套版」/);
  assert.match(decksSource, /deck-page-grid/);
  assert.match(decksSource, /deck-stack-preview/);
  assert.match(decksSource, /deck-preview-expand/);
  assert.match(decksSource, />使用<\/button>/);
  assert.match(decksSource, /title="删除套版">🗑️<\/button>/);
  assert.match(decksSource, /title="展开浏览">＋<\/button>/);
  assert.match(decksSource, /class="secondary deck-rc-open"[^>]*>🗑️<\/button>/);
  assert.match(decksSource, /aria-label="关闭预览">×<\/button>/);
  assertDecl('.deck-page-grid', 'grid-template-columns: repeat(2, minmax(0, 1fr))');
  assertDecl('.deck-card-actions', 'justify-content: flex-end');
});

test('deck cards use landscape stack preview beside details and include recycle bin', () => {
  assert.match(decksSource, /showDeckRecycleModal/);
  assert.match(apiSource, /listDeckRecycleBin/);
  assert.match(apiSource, /restoreDeckRecycle/);
  assert.match(apiSource, /purgeDeckRecycle/);
  assert.match(decksSource, /deckStackHtml\(firstPreviewOfDeck\(d\), d\.name, d\.pageCount\)/);
  assert.match(decksSource, /class="deck-preview-wrap deck-preview-landscape"/);
  assert.match(decksSource, /class="secondary deck-rc-open"/);
  assert.doesNotMatch(decksSource, /deck-folder-new/);
  assertDecl('.deck-card-main', 'flex-direction: row');
  assertDecl('.deck-preview-wrap.deck-preview-landscape', 'width: 150px');
});

test('deck builder supports drag reorder with snapping grid cards', () => {
  assert.match(decksSource, /pageOrdinal\(i\)/);
  assert.match(decksSource, /class="deck-page-number"/);
  assert.match(decksSource, /draggable="true"/);
  assert.match(decksSource, /dragstart/);
  assert.match(decksSource, /dragover/);
  assert.match(decksSource, /drop/);
  assert.match(decksSource, /moveRow\(from, to\)/);
  assertDecl('.deck-page-card.drag-over', 'border-color: var(--brand)');
  assertDecl('.deck-page-number', 'position: absolute');
  assertDecl('.deck-page-number', 'background: rgba(255, 255, 255, .94)');
  assertDecl('.deck-page-number', 'border: 1px solid rgba(31, 56, 100, .18)');
});

test('deck wizard mirrors single-template wizard structure with grouped page fills', () => {
  assert.match(deckWizardSource, /<label>套版<\/label>/);
  assert.match(deckWizardSource, /<select id="wb-template" style="flex:1"/);
  assert.match(deckWizardSource, /id="dw-back" title="返回套版">返回<\/button>/);
  assert.match(deckWizardSource, /class="card wb-write-bar"/);
  assert.match(deckWizardSource, /class="wb-adv gp-adv"/);
  assert.match(deckWizardSource, /infoTip\(/);
  assert.match(deckWizardSource, /backgroundPanelHtml/);
  assert.match(deckWizardSource, /deck-page-fill-card/);
  assert.match(deckWizardSource, /deck-page-fill-card accent-/);
  assert.match(deckWizardSource, /global-theme/);
  assert.match(deckWizardSource, /wb-sec-title/);
  assert.match(deckWizardSource, /wb-text-slot/);
  assert.match(deckWizardSource, /wb-img-slot/);
  assert.match(deckWizardSource, /slot-prompt/);
  assert.match(deckWizardSource, /om-limit/);
  assert.match(deckWizardSource, /write-ppt/);
  assert.match(deckWizardSource, /img-pick/);
  assert.match(deckWizardSource, /img-file/);
  assert.match(deckWizardSource, /img-dropzone/);
  assert.match(deckWizardSource, /img-nav/);
  assert.match(deckWizardSource, /img-crop/);
  assert.match(deckWizardSource, /img-clear/);
  assert.match(deckWizardSource, /wb-preview-card/);
  assert.match(deckWizardSource, /wb-preview-head/);
  assert.match(deckWizardSource, /wb-tbl-slot/);
  assert.match(deckWizardSource, /tableSlotHtml/);
  assert.match(deckWizardSource, /tbl-tabs/);
  assert.match(deckWizardSource, /tbl-paste-ta/);
  assert.match(deckWizardSource, /tbl-ai-req/);
  assert.doesNotMatch(deckWizardSource, /dw-build|dw-gen|dw-img-search|dw-text|dw-var|dw-img-q|✍|✓|⚠|📅/);
  assert.doesNotMatch(deckWizardSource, /<h1 class="page-title">套版生成<\/h1>/);
  assertDecl('.deck-page-fill-card', 'border: 1px solid var(--border)');
  assertDecl('.deck-page-fill-card.accent-1', 'background: #F8FAFF');
});

test('deck wizard keeps AI service panels open and bounds inline preview height', () => {
  assert.match(deckWizardSource, /refreshOutputModeForPageSlot/);
  assert.doesNotMatch(deckWizardSource, /mode\.touched = true;\s*render\(\);/);
  assert.match(deckWizardSource, /class="card wb-preview-card deck-preview-card"/);
  assert.match(deckWizardSource, /class="deck-preview-body"/);
  assert.match(deckWizardSource, /class="deck-preview-inner"/);
  assertDecl('.deck-preview-body', 'max-height: min(360px, 62vh)');
  assertDecl('.deck-preview-body', 'overflow-y: auto');
});
