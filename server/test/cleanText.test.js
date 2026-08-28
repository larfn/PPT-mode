'use strict';

const { test } = require('node:test');
const assert = require('node:assert');
const { cleanAiText } = require('../src/cleanText.js');

// —— 始终生效：Markdown 痕迹剥离 ——

test('strips heading markers', () => {
  assert.equal(cleanAiText('### 标题\n\n正文内容'), '标题\n\n正文内容');
});

test('strips bold / italic / strikethrough markers', () => {
  assert.equal(cleanAiText('**重要内容** 和 ~~旧内容~~'), '重要内容 和 旧内容');
  assert.equal(cleanAiText('*强调* 与 _斜体_'), '强调 与 斜体');
});

test('strips inline code and code fences', () => {
  assert.equal(cleanAiText('运行 `npm test` 即可'), '运行 npm test 即可');
  assert.equal(cleanAiText('```js\nconst a = 1;\n```'), 'const a = 1;');
});

test('strips links and images, keeps text', () => {
  assert.equal(cleanAiText('参考[文档](https://x.com) ![图](http://y.com/a.png)'), '参考文档');
});

test('strips blockquote and hr lines', () => {
  assert.equal(cleanAiText('> 引用\n\n---\n\n正文'), '引用\n\n正文');
});

test('collapses 3+ blank lines to 2 and trims edges', () => {
  assert.equal(cleanAiText('a\n\n\n\n\nb\n\n'), 'a\n\nb');
});

// —— plain 模式：剥离列表编号/符号 ——

test('plain mode strips numbered and bullet lists', () => {
  assert.equal(cleanAiText('1. 第一点\n2. 第二点\n3. 第三点', { plain: true }), '第一点\n第二点\n第三点');
  assert.equal(cleanAiText('- 甲\n- 乙\n- 丙', { plain: true }), '甲\n乙\n丙');
});

test('keeps list markers when plain not set (bullets mode)', () => {
  assert.equal(cleanAiText('1. 第一点\n2. 第二点'), '1. 第一点\n2. 第二点');
  assert.equal(cleanAiText('- 甲\n- 乙'), '- 甲\n- 乙');
});

// —— 误伤防护 ——

test('does not strip decimals like 3.14 or math 5 * 3', () => {
  assert.equal(cleanAiText('3.14 是圆周率，5 * 3 = 15', { plain: true }), '3.14 是圆周率，5 * 3 = 15');
});

test('does not mangle snake_case', () => {
  assert.equal(cleanAiText('file_name_v1 与 hello_world'), 'file_name_v1 与 hello_world');
});

test('table JSON survives cleaning untouched', () => {
  const json = '[["1. 加热","2. 冷却"],["步骤A","步骤B"]]';
  assert.equal(cleanAiText(json, { plain: true }), json);
});

// —— 长度约束 ——

test('truncates overlong text with ellipsis at maxChars', () => {
  const out = cleanAiText('这是一段很长很长的内容，超过十一个字符', { maxChars: 10 });
  assert.ok(out.length <= 11);
  assert.ok(out.endsWith('…'));
});

test('truncates lines at maxLines', () => {
  assert.equal(cleanAiText('a\nb\nc\nd\ne', { maxLines: 3 }), 'a\nb\nc');
});

test('0 / missing limits mean no truncation', () => {
  const long = 'x'.repeat(500);
  assert.equal(cleanAiText(long, { maxChars: 0 }), long);
  assert.equal(cleanAiText(long), long);
});

test('combined: heading + list + bold in plain mode', () => {
  assert.equal(cleanAiText('### 方案\n\n1. **先做A**\n2. 再做B\n\n> 备注', { plain: true }), '方案\n\n先做A\n再做B\n\n备注');
});
