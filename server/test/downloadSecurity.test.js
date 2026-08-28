// 下载安全单元测试：SSRF 防护 / 协议限制 / 魔数 / MIME
const { test } = require('node:test');
const assert = require('node:assert');
const { isPrivateIp, validateDownloadUrl, sniffImageMime, extFromMime, ALLOWED_IMAGE_MIME, _setDnsLookup } = require('../src/downloadSecurity.js');

test('isPrivateIp covers private/reserved ranges', () => {
  for (const ip of ['10.0.0.1', '127.0.0.1', '0.0.0.0', '169.254.1.1', '172.16.0.1', '172.31.255.255', '192.168.1.1', '100.64.0.1', '100.127.255.255', '224.0.0.1', '255.255.255.255', '::1', 'fe80::1', 'fc00::1', '::ffff:10.0.0.1', '::ffff:192.168.0.1']) {
    assert.equal(isPrivateIp(ip), true, ip + ' 应为私网/保留');
  }
  for (const ip of ['8.8.8.8', '93.184.216.34', '1.1.1.1', '172.32.0.1', '192.169.1.1', '::ffff:8.8.8.8']) {
    assert.equal(isPrivateIp(ip), false, ip + ' 应为公网');
  }
});

test('validateDownloadUrl rejects non-http(s) protocols and localhost', async () => {
  assert.equal((await validateDownloadUrl('file:///etc/passwd')).ok, false);
  assert.equal((await validateDownloadUrl('ftp://x.com/a.jpg')).ok, false);
  assert.equal((await validateDownloadUrl('javascript:alert(1)')).ok, false);
  assert.equal((await validateDownloadUrl('https://localhost/x.jpg')).ok, false);
  assert.equal((await validateDownloadUrl('https://127.0.0.1/x.jpg')).ok, false);
  assert.equal((await validateDownloadUrl('http://10.1.2.3/x.jpg')).ok, false);
  assert.equal((await validateDownloadUrl('https://[::1]/x.jpg')).ok, false);
  assert.equal((await validateDownloadUrl('not a url')).ok, false);
});

test('validateDownloadUrl checks DNS-resolved addresses (SSRF via domain)', async () => {
  _setDnsLookup(async () => [{ address: '10.0.0.5', family: 4 }]);
  const r = await validateDownloadUrl('https://internal.example.com/img.png');
  assert.equal(r.ok, false);
  assert.ok(r.error.includes('内网'), r.error);

  _setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);
  const ok = await validateDownloadUrl('https://public.example.com/img.png');
  assert.equal(ok.ok, true);

  _setDnsLookup(async () => []);
  assert.equal((await validateDownloadUrl('https://nodns.example.com/x.png')).ok, false);
});

test('sniffImageMime detects jpg/png/gif/webp and rejects unknown (svg)', () => {
  const png = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex');
  assert.equal(sniffImageMime(png), 'image/png');
  const jpg = Buffer.from('ffd8ffe000104a464946000101', 'hex');
  assert.equal(sniffImageMime(jpg), 'image/jpeg');
  const gif = Buffer.from('47494638396101000100800000', 'hex');
  assert.equal(sniffImageMime(gif), 'image/gif');
  const webp = Buffer.from('52494646100000005745425000', 'hex');
  assert.equal(sniffImageMime(webp), 'image/webp');
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>', 'utf8');
  assert.equal(sniffImageMime(svg), null, 'SVG 文本无魔数 → 拒绝');
  assert.equal(sniffImageMime(Buffer.alloc(4)), null, '过短 → null');
  assert.equal(sniffImageMime(null), null);
});

test('extFromMime whitelist only', () => {
  assert.equal(extFromMime('image/jpeg'), 'jpg');
  assert.equal(extFromMime('image/png'), 'png');
  assert.equal(extFromMime('image/webp'), 'webp');
  assert.equal(extFromMime('image/gif'), 'gif');
  assert.equal(extFromMime('image/svg+xml'), '');
  assert.equal(extFromMime('text/html'), '');
});

test('ALLOWED_IMAGE_MIME rejects svg and non-images', () => {
  assert.ok(ALLOWED_IMAGE_MIME.has('image/jpeg'));
  assert.ok(ALLOWED_IMAGE_MIME.has('image/webp'));
  assert.ok(!ALLOWED_IMAGE_MIME.has('image/svg+xml'));
  assert.ok(!ALLOWED_IMAGE_MIME.has('text/html'));
});
