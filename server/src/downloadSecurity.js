// 下载安全：SSRF 防护 / 协议限制 / MIME 白名单 / 图片魔数校验
'use strict';
const dns = require('node:dns').promises;
const net = require('node:net');

// 允许下载的图片 MIME（拒绝 SVG——可能携带脚本；拒绝其他非图片）
const ALLOWED_IMAGE_MIME = new Set(['image/jpeg', 'image/png', 'image/webp', 'image/gif']);

const BLOCKED_HOSTNAMES = new Set(['localhost', 'localhost.localdomain', '0.0.0.0', '127.0.0.1', '::1']);

// 私有/保留/组播地址一律禁止（SSRF 防护核心）
function isPrivateIp(ip) {
  if (!ip) return true;
  if (net.isIPv4(ip)) {
    const parts = ip.split('.').map(Number);
    const a = parts[0], b = parts[1];
    if (a === 0) return true;                            // 0.0.0.0/8
    if (a === 10) return true;                           // 10/8
    if (a === 127) return true;                          // 127/8
    if (a === 169 && b === 254) return true;             // 169.254/16（链路本地）
    if (a === 172 && b >= 16 && b <= 31) return true;    // 172.16/12
    if (a === 192 && b === 168) return true;             // 192.168/16
    if (a === 100 && b >= 64 && b <= 127) return true;   // 100.64/10（运营商 NAT）
    if (a >= 224) return true;                           // 组播/保留
    return false;
  }
  const lower = String(ip).toLowerCase();
  if (lower === '::1' || lower === '::' || lower === '::ffff:127.0.0.1') return true;
  if (lower.startsWith('fe80') || lower.startsWith('fc') || lower.startsWith('fd')) return true; // 链路本地/ULA
  if (lower.startsWith('::ffff:')) {
    const v4 = lower.split('::ffff:')[1];
    if (v4 && net.isIPv4(v4)) return isPrivateIp(v4);    // IPv4 映射地址
  }
  return false;
}

// 测试可注入的 DNS 解析器（默认真实解析；测试注入公网/内网地址）
let lookupFn = async (host) => {
  const r = await dns.lookup(host, { all: true });
  return r;
};
function _setDnsLookup(fn) { lookupFn = fn || (async (h) => dns.lookup(h, { all: true })); }

// 校验一个下载 URL：协议 + 主机 + DNS 全部解析地址均不得指向内网/本机
// 返回 { ok:true, host, ips } 或 { ok:false, error }
async function validateDownloadUrl(url) {
  let parsed;
  try { parsed = new URL(url); } catch { return { ok: false, error: '非法 URL' }; }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    return { ok: false, error: '仅支持 http/https 下载（收到 ' + parsed.protocol + '）' };
  }
  const host = parsed.hostname.toLowerCase();
  if (!host) return { ok: false, error: 'URL 缺少主机名' };
  if (BLOCKED_HOSTNAMES.has(host)) return { ok: false, error: '禁止访问本机地址（' + host + '）' };
  if (net.isIP(host)) {
    if (isPrivateIp(host)) return { ok: false, error: '禁止访问内网/保留地址（' + host + '）' };
    return { ok: true, host, ips: [host] };
  }
  try {
    const addrs = await lookupFn(host);
    if (!addrs || !addrs.length) return { ok: false, error: '域名解析失败（' + host + '）' };
    for (const a of addrs) {
      if (isPrivateIp(a.address)) {
        return { ok: false, error: '禁止访问内网/保留地址（' + host + ' → ' + a.address + '）' };
      }
    }
    return { ok: true, host, ips: addrs.map((a) => a.address) };
  } catch {
    return { ok: false, error: '域名解析失败（' + host + '）' };
  }
}

// 图片魔数识别：JPG/PNG/GIF/WebP；未知（含 SVG 文本）返回 null → 拒绝
function sniffImageMime(buf) {
  if (!buf || buf.length < 8) return null;
  if (buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) return 'image/jpeg';
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) return 'image/png';
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return 'image/gif';
  if (buf.length >= 12 && buf.toString('ascii', 0, 4) === 'RIFF' && buf.toString('ascii', 8, 12) === 'WEBP') return 'image/webp';
  return null;
}

// content-type → 扩展名（仅白名单内返回）
function extFromMime(mime) {
  switch (mime) {
    case 'image/jpeg': return 'jpg';
    case 'image/png': return 'png';
    case 'image/webp': return 'webp';
    case 'image/gif': return 'gif';
    default: return '';
  }
}

module.exports = {
  ALLOWED_IMAGE_MIME, BLOCKED_HOSTNAMES, isPrivateIp, validateDownloadUrl, sniffImageMime, extFromMime, _setDnsLookup
};
