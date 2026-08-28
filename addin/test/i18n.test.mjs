// i18n 逻辑回归测试：用 esbuild 编译 src/lib/i18n.ts 后在 Node 中验证
// 覆盖：中英切换恢复（原文不被吞）、动态新内容识别、词典/句式/多行/括号翻译
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildSync } from 'esbuild';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const store = {};
globalThis.localStorage = {
  getItem: (k) => (k in store ? store[k] : null),
  setItem: (k, v) => { store[k] = String(v); },
  removeItem: (k) => { delete store[k]; }
};
globalThis.document = { documentElement: {}, body: null };

const out = buildSync({ entryPoints: ['src/lib/i18n.ts'], bundle: true, format: 'cjs', write: false });
const mod = { exports: {} };
new Function('module', 'exports', 'require', out.outputFiles[0].text)(mod, mod.exports, require);
const i18n = mod.exports;

test('中文→英文→中文 原文不被吞（问题 1 回归）', () => {
  // 英文态下 DOM 文本是英文变体，此时切回中文：原文必须保留，否则永远回不去
  assert.equal(i18n.updatedOriginal('模板库', 'Template Library'), '模板库');
  assert.equal(i18n.updatedOriginal('模板库', '模板库'), '模板库');
  assert.equal(i18n.updatedOriginal('保存配置', 'Save Settings'), '保存配置');
  assert.equal(i18n.updatedOriginal('加载中…', 'Loading...'), '加载中…');
  // 属性原文同样不被吞
  assert.equal(i18n.updatedOriginal('版本历史 / 恢复 / 设为当前', 'Version History / Restore / Set Current'), '版本历史 / 恢复 / 设为当前');
});

test('外部写入的新内容会被识别并更新原文', () => {
  assert.equal(i18n.updatedOriginal('生成中…', '已生成'), '已生成');
  assert.equal(i18n.updatedOriginal('旧内容', '新内容'), '新内容');
  // 未收录的中文原文在英文态显示原文（不翻译），不会被误判
  assert.equal(i18n.updatedOriginal('未收录中文', '未收录中文'), '未收录中文');
});

test('translateText 只在英文态翻译', () => {
  i18n.setLanguage('zh');
  assert.equal(i18n.translateText('模板库'), '模板库');
  i18n.setLanguage('en');
  assert.equal(i18n.translateText('模板库'), 'Template Library');
  assert.equal(i18n.translateText('AI 配置'), 'AI Settings');
  i18n.setLanguage('zh');
  assert.equal(i18n.translateText('Template Library'), 'Template Library');
  i18n.setLanguage('zh');
  assert.equal(i18n.translateText('模板库'), '模板库');
});

test('精确词典覆盖关键界面文本', () => {
  const cases = [
    ['AI 配置', 'AI Settings'], ['语言', 'Language'], ['界面字体大小', 'UI Font Size'],
    ['保存配置', 'Save Settings'], ['模板库', 'Template Library'], ['生成向导', 'Generate'],
    ['套版', 'Decks'], ['保存模板', 'Save Template'], ['回收站', 'Recycle Bin'],
    ['版本历史', 'Version History'], ['清空回收站', 'Empty Recycle Bin'], ['导出诊断文件', 'Export Diagnostic File'],
    ['生成质量检查发现异常', 'Quality Check Issues'], ['尚未配置文本生成服务（API Key），无法自动生成文字。请先前往配置。', 'Text generation is not configured with an API key. Go to settings first.'],
    ['同一行', '同一行']
  ];
  for (const [zh, en] of cases) {
    assert.equal(i18n.translateToEn(zh), en, '翻译失败: ' + zh);
  }
});

test('动态句式 pattern 覆盖', () => {
  const cases = [
    ['第 3 页', 'Page 3'],
    ['12 张', '12 images'],
    ['共 12 张', '12 images'],
    ['最多200字', 'Max 200 chars'],
    ['至少 5 字', 'Min 5 chars'],
    ['正在下载 45%', 'Downloading 45%'],
    ['已写入整份 PPT（5 页）', 'Inserted full deck (5 pages)'],
    ['文本位 2', 'Text Slot 2'],
    ['图片位 1', 'Image Slot 1'],
    ['表格 3', 'Table 3'],
    ['生成失败：网络错误', 'Generation failed: 网络错误'],
    ['删除失败：权限', 'Delete failed: 权限'],
    ['已恢复为 v2 ✓', 'Restored to v2 ✓'],
    ['已删除 v1', 'Deleted v1'],
    ['版本历史（当前 v3）', 'Version History (current v3)'],
    ['已识别 5 个元素', 'Recognized 5 elements'],
    ['共 8 个元素：页面 2 · 版式 3 · 母版 3', '8 elements total: slide 2 · layout 3 · master 3'],
    ['约需 4 行，框内仅能容纳约 3 行，会溢出', 'Needs about 4 lines, but only about 3 fit; will overflow'],
    ['有 2 个单元格文字超出，行会被撑高（可能超出页面）', '2 cells have overflowing text; rows will expand (may exceed the slide)']
  ];
  for (const [zh, en] of cases) {
    assert.equal(i18n.translateToEn(zh), en, '句式翻译失败: ' + zh);
  }
});

test('多行消息逐行翻译', () => {
  const out = i18n.translateToEn('读取模板失败：boom\n请稍后重试');
  assert.equal(out, 'Failed to read template: boom\n请稍后重试');
  const out2 = i18n.translateToEn('插入 PPT 超时（PowerPoint 长时间没有响应）。\n请稍等片刻观察演示文稿中是否已出现新页面；若没有，可重试，或完全退出并重新打开 PowerPoint。');
  assert.ok(!out2.includes('插入 PPT 超时'), '第一行应被翻译: ' + out2);
  assert.ok(!out2.includes('请稍等片刻观察'), '第二行应被翻译: ' + out2);
  assert.ok(out2.includes('Inserting into PPT timed out'));
});

test('括号结构与建议句式拆分翻译', () => {
  assert.equal(i18n.translateToEn('文字（主标题）'), 'Text (Title)');
  assert.equal(i18n.translateToEn('建议：文字（主标题）'), 'Suggestion: Text (Title)');
  assert.equal(i18n.translateToEn('主标题：占位符类型=title'), 'Title: placeholder type=title');
  assert.equal(i18n.translateToEn('建议：表格位（生成时填数据）'), 'Suggestion: Table Slot (data filled when generating)');
});
