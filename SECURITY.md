# 安全策略

## 支持范围

安全修复以当前 `master` 分支和最新 GitHub Release 为准。旧安装包不会单独维护，请先升级到最新版本再复现问题。

## 报告安全问题

请使用 GitHub 仓库的 **Security → Report a vulnerability** 私下报告，不要在公开 Issue 中粘贴 API Key、Token、配置文件、演示文稿内容或可直接利用的攻击样例。

报告中请包含受影响版本、复现步骤、预期与实际结果，以及已做脱敏处理的日志。确认问题后会在修复可用时发布说明。

## 本地安全边界

后端只监听 `127.0.0.1`，受运行时随机 Token 保护；浏览器跨域访问仅允许本机回环来源。API Key 在 Windows 上使用当前用户范围的 DPAPI 加密保存。下载链路会校验协议、DNS 解析地址、重定向、MIME、文件魔数、大小和超时，并拒绝 SVG、内网地址及伪装图片。

`/api/runtime` 是前端同源启动所需的引导端点，因此不要求 Token；跨域读取仍受本机来源限制。本机同一用户下的其他进程本来就能读取对应运行时文件，不把本机用户会话视为互不信任边界。

## 已知依赖告警

截至 2026-08-28，`npm audit` 会通过 `pptxgenjs` 的传递依赖报告 `image-size` 的两个拒绝服务告警（GHSA-w3rx-r6r6-pgpr、GHSA-5p2g-fcmc-qvqq），上游尚无已发布修复版本。当前 PptxGenJS Node 构建未调用该包，插件进入生成链路的图片也被限制为经过魔数验证的 JPEG、PNG、GIF 或 WebP；因此已知恶意 ICNS、JXL、HEIF 解析路径在本项目中不可达。发布时仍应保留该告警并持续跟踪上游，不要使用 `npm audit fix --force` 建议的 PptxGenJS 1.x 降级方案。
