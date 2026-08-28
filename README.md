# 模板助手

模板助手是一款面向 Windows 桌面版 PowerPoint 的本地加载项。它可以把当前幻灯片保存成可复用模板，按模板填充文字、图片和变量，也可以把多个模板编排为整份套版。

## 主要功能

* 从当前幻灯片保存模板，并保留元素位置、文字样式、背景和固定图片
* 自动识别标题、正文、图片、变量、表格等内容位，可手动调整
* 按模板生成单页，支持文字生成、联网搜图、变量填充和写回 PowerPoint
* 把多个模板组成套版，一次生成多页演示文稿
* 模板版本管理、回收站、套版回收站和本地迁移
* 自定义图片来源，以及供外部 AI 客户端调用的 MCP 工具
* 界面支持中文 / English 切换
* 本地服务与一次性 Token 鉴权；API Key 在 Windows 下使用 DPAPI 加密保存

## 快速安装

1. 双击 `安装中心.exe`；如果无法启动，改用 `安装中心.bat`。
2. 按向导完成环境检查和安装。
3. 完全退出并重新打开 PowerPoint，在「开始」选项卡点击「模板助手」。

已有发布产物时，目标电脑无需安装 Node.js。需要从源码重新构建，或使用 MCP 服务时，需要 Node.js 24+。

完整步骤、升级、卸载和故障排查见 [安装说明](安装说明.md)。

## 基本使用

1. 在 PowerPoint 中设计一页模板。
2. 打开「保存模板」，读取当前页并确认各元素用途。
3. 到「模板库」选择模板，填写内容并写入 PowerPoint。
4. 需要多页时，到「套版」把多个模板排序组合后统一生成。
5. 文本服务、自定义图源、界面语言和字号在「AI 配置」中管理。

## 文档

* [文档索引](docs/README.md)
* [安装说明](安装说明.md)
* [技术文档](docs/插件技术总结.md)
* [自定义图源导入说明](自定义图源导入说明.txt)
* [安全说明](SECURITY.md)

## 本地开发与验证

```powershell
cd server
npm ci
npm test
npm run e2e:file

cd ..\\addin
npm ci
npm test
node node\_modules/typescript/bin/tsc --noEmit
npm run build
```

完整发布构建在项目根目录运行：

```powershell
npm run release
```

发布脚本会统一版本、执行测试与黄金路径检查、构建前端和后端可执行文件，并更新 `release.json`。

## 运行环境

* Windows 10 / 11
* Microsoft 365 PowerPoint 或 Office 2016+
* Node.js 24+：仅源码构建和 MCP 使用需要

## License

[MIT](LICENSE)
