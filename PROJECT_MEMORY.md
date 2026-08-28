# 项目记忆（PROJECT MEMORY）

> 用途：记录 `ppt模塊化插件` 项目的背景、已确认决策、架构部署、Bug 修复历史、环境信息和待办问题，供后续会话、子代理接续上下文。
> \*\*维护规则：每做重大修复/优化后，务必及时更新本文件（新增/修订对应小节）\*\*，保持与代码、Git 历史同步。
> \*\*记录原则（2026-08-23 用户要求）：每次记录都像 2026-08-23 精简版这样精简记录\*\*——按「里程碑简表」格式：每条只写做了什么 + 关键决策/教训 + 特殊部署提示，用 1\~3 行；不记过程性细节（测试数字明细、排查过程、函数签名等）；条目过多或过长时及时压缩合并。

## 〇、交流偏好（用户要求，务必遵守）

* **每次任务报告完后，必须追加一段大白话解释**：说明这次到底改了什么、为什么、用户需要做什么（如部署步骤）。报告可以专业（文件/测试/字段），但解释必须通俗，避免术语堆砌，多用类比和口语。
* **对用户输出一律使用中文**（含最终汇报、任务解释、技能清单等）；代码注释/变量名等工程产物不受此限。
* **最小改动原则（2026-08-23 用户要求）：用户未提及的地方一律不准改**。只改用户明确要求的内容；涉及顺带可见的明显小问题（如死代码、错别字）也不得擅自动手，除非用户提及或明确授权。改完自查时若发现无关问题，先记入汇报/待办，不直接修改。本约束适用所有后续会话。
* 本偏好由用户 2026-08-20 明确要求，适用所有后续会话（已同步到本文件供子代理/新会话接续）。

## 〇·一、项目级 Skills（2026-08-20 安装）

按用户要求，公开技能**仅安装到本项目专用目录**（非全局）：`.dsh/skills/`，由 DSH 自动扫描加载。

|技能|来源|说明|
|-|-|-|
|powerpoint|igorwarzocha/opencode-workflows|PPT 创建/设计/分析，自带 OOXML（ISO/IEC 29500）XSD 参考 + html2pptx/inventory/thumbnail 脚本|
|pptx|huyansheng3/ppt-skills|PPTX 创建/编辑/分析（zip+XML 工作流），含 OOXML 参考|
|single-slide-ppt|huyansheng3/ppt-skills|基于 PptxGenJS 的单页 PPT 生成（与后端 slideBuilder 同栈）|
|images-search|brave/brave-search-skills|Brave 图片搜索 API（需 BRAVE\_SEARCH\_API\_KEY，未配置时不可用）|

安装方式：GitHub tarball 下载 + 解压到指定目录（非全局 skills CLI）。验证：DSH 会话技能目录已自动识别全部 4 个技能，frontmatter 均有效。

## 一、项目概况

在 PowerPoint 中开发一个**加载项（Office Add-in）**，让用户通过"模板 + AI"的方式快速生成 PPT 页面。

核心工作流：

1. 用户在 PPT 中自定义设计模板页（图片位置、文字位置、格式、字体等）；
2. 将模板页保存为模板（存本地专属文件夹，可复制到另一台机器使用）；
3. 需要生成页面时，从模板库选择模板；
4. 用户向"内嵌 AI"描述需要的图片，AI 在**网上搜索真实图片**供用户选择；
5. AI 依据所选图片 + 用户的文字需求生成文字内容，**段落数取决于模板中文本嵌入框的数量**；
6. 用户在插件中检查生成结果（图片 + 文字）；
7. 确认后，插件把该页**直接生成到当前 PPT 文档中**。

## 二、已确认决策

|决策点|结论|备注|
|-|-|-|
|插件形态|Office 加载项（Office.js 网页技术，任务窗格形式）|用户明确"在PPT内叫做加载項"|
|模板编辑方式|直接在 PPT 里设计模板页，插件读取形状信息存为模板|所见即所得|
|模板存储|本地专属文件夹，模板 = 一个 JSON + 预览图，可整文件夹复制到另一台机器|路径：文档/PPT模板库/|
|AI 配置|独立"AI 配置"页，**查图与文本分开配置**（供应商/API Key/地址/模型），Key 存本地配置，界面掩码显示|配置存 %APPDATA%/ppt-ai-addin/config.json|
|图片来源|百度图片搜索（页面解析，默认）+ 必应图片搜索（页面解析）双源可选；未实现 Bing API（Azure）|配置页可切换|
|文本服务|OpenAI 兼容接口，默认预填 DeepSeek（国内直连）|配置页可改|
|内容位标记|4 类：AI 图片位 / AI 文本位 / 手动变量位（如页码 05）/ 固定元素|保存模板时逐个标记|
|生成方式|后端用 pptxgenjs 生成单页 .pptx（base64），前端经 Presentation.insertSlidesFromBase64 插入当前选中页之后|PowerPointApi 1.2|
|项目开源|计划开源（MIT，LICENSE 已建）||

## 三、环境信息

|项目|版本/值|
|-|-|
|OS|Windows 11（10.0.26100）|
|Node.js / npm|v24.16.0 / 11.13.0|
|Office|Microsoft 365，16.0.20228.20190，x64|
|工作区|C:/Users/larfn/Documents/ChatGPT/ppt模塊化插件|
|模板目录|C:/Users/larfn/Documents/PPT模板库（封面、橫向 两个模板）|
|图片下载目录|C:/Users/larfn/Documents/PPT下载图库|
|安装目录|%LOCALAPPDATA%/PPT-AI-Addin/（exe + dist + manifest + 自启动脚本）|
|配置存储|%APPDATA%/ppt-ai-addin/config.json（API Key 为 dpapi: 加密）|
|本地服务|http://127.0.0.1:3788（被占用自动 +1），健康检查 GET /api/health，express.json 上限 300mb|
|历史故障|2026-08-19 晚原生进程曾集体启动失败 0xC0000142（疑杀毒锁定），重启可恢复；故障期间可用 run\_code 内嵌 Node 验证|

## 四、架构与部署

* 前端：`addin/`（TypeScript + Vite 任务窗格，`src/pages/` 下模板库/生成向导/保存模板/AI 配置/套版五页）。
* 后端：`server/`（Node.js Express + pptxgenjs + jszip；`src/slideBuilder.js` 生成单页 pptx；`src/pptContext.js` + `mcp/\*.ps1` 经 PowerShell COM 读写当前 PPT；`src/analyze.js` AI 分析增强）。
* 打包：前端 `cd addin; npm run build` → `addin/dist`；后端 `cd server; npm run build:exe`（pkg）→ `dist-exe/ppt-ai-addin.exe`。
* 部署：运行根目录 **`安装.bat`**（Build→Package→Install→Verify 流水线：自动跑 release.js 构建 + 复制到 `%LOCALAPPDATA%/PPT-AI-Addin/` + 注册 WEF + verify-install.ps1 自检）。**任何改动后必须重跑 安装.bat，用户需完全退出并重开 PowerPoint。**
* 验证：后端 `cd server; npm test`；前端 `cd addin; npx tsc --noEmit; npm run build`；E2E `cd server; node e2e/golden-path.js`。
* 版本：统一 `scripts/release.js` 生成 `YYYY.MM.DD.NN`，单一来源 `server/src/version.js` + `release.json`，vite define 注入前端 `\_\_APP\_VERSION\_\_`。
* 静态资源：express.static 已设 `Cache-Control: no-store`（防任务窗格缓存旧页面）。

## 五、开发进度（里程碑简表）

> \*\*最新状态（2026-08-28）\*\*：界面中英双语切换完成（AI 配置→界面显示→语言，即时生效+保存后重启沿用）+ 历史文档/临时文件清理；验证全绿（server 296 + addin 26 + tsc 0 + vite build）。**待用户重跑 安装.bat + 完全重启 PowerPoint 实测**（此前表格引擎里程碑同样待实测）。
> \*\*此前状态（2026-08-24）\*\*：表格自动排版引擎完成（见下条里程碑）——粘贴/AI 表格内容自动适配模板表格位（列宽/行高/全表统一字号 10~14pt/超出向下延伸+红色警告）、「相同字樣一鍵合并」、MCP 支持表格位（get_template tableInfo + generate_slide/deck tables + fit_table 工具）。**待用户重跑 安装.bat + 完全重启 PowerPoint 实测**（含此前「加载项无效」修复与 08-23 规则引擎批次）。
> \*\*通用部署提示\*\*：除特别注明外，每个里程碑完成后均需\*\*重跑 安装.bat（自动 bump 版本 + E2E 门禁 + 自检）→ 完全退出重开 PowerPoint\*\*；MCP 改动后需用 启动MCP.bat 重启。

* **界面中英双语切换 + 文档清理（2026-08-28，用户指定）**：AI 配置「界面显示」新增语言下拉（中文/English）——切换即时生效、保存写 config ui.language（默认 zh，config 路由 GET/PUT 已成对）、重开沿用；前端新增 lib/i18n.ts（精确词典 + 正则模式：张/页/失败前缀/最多N字等 + MutationObserver 自动翻译动态增改文本及 title/placeholder/aria-label，WeakMap 记原文防二次翻译漂移），taskpane 启动 applyLanguageFromConfig + startAutoTranslate；翻译只覆盖界面文本，模板内容/用户输入/AI 生成内容不翻。文档：README/安装说明/技术总结/CHANGELOG 同步语言说明，docs/README.md 重建索引。清理：删 docs/superpowers 历史计划与规格、根目录 _tmp_*.bat 向导临时副本、.dsh/.dsh-vision-toolkit 旧工具缓存；PROJECT_MEMORY.md 保留（本文件即项目记忆）。验证：server 296/296、addin 26/26（uiStyleSystem 含语言开关断言）、tsc 0、vite build。部署：重跑 安装.bat 重启 PPT。**实测修复（2026-08-28 用户反馈 2 项）**：①切英文后回不来中文——根因：原文判定用「当前语言」算应有显示，切回中文瞬间 DOM 还是英文，误判「外部改写」把英文吞成新原文；修复：拆出 translateToEn（不依赖当前语言）+ updatedOriginal 纯函数（current 是原文或原文英文变体都不算新内容），属性同理；settings 语言下拉取值改 localStorage（实际生效语言）防「切了没保存」显示不一致。②英文态部分中文残留——词典 exact 183→约 460 条 + patterns 约 200 条（覆盖动态句式：失败类/回读状态/质量检查/分析 reason/多行错误弹窗逐行/括号结构拆分 head（inner）/分号并列段拆分）；translateToEn 增强多行+括号+分号兜底；扫描脚本（esbuild 编译 i18n 后遍历 src 全中文串实测）未覆盖 595→163 且剩余均为拼接断点词素/HTML 属性碎片/infoTip 替换残渣（非独立 text node）；AI 提示词/文件名（安装.bat 等）故意不翻。新增 addin/test/i18n.test.mjs（esbuild 编译 + localStorage/document stub 的 zh→en→zh 恢复/新内容识别/词典/句式/多行/括号回归）。验证：server 296/296、addin 27 项含 i18n 回归全绿、tsc 0、vite build。部署：重跑 安装.bat 重启 PPT 实测。
* **表格自动排版引擎（2026-08-24，用户指定）**：表格位=槽位（bounds+样式），粘贴/AI 内容自动适配列宽/行高/字号。引擎 addin/src/lib/tableModel.ts（fitTableLayout：min/max 宽→shrink-to-fit→富余按贪心分配→整表等比缩字 下限10/基准14→行高按折行拉极限→放不下锁槽宽向下延伸+overflow；colspan 迭代归一化；mergeSameTextCells 一键合并 trim 相同连续≥2 横/纵/auto）+ tableClipboard.ts（HTML 粘贴解析含 rowspan 感知列定位、CSV/TSV 降级、样式全丢）。后端 slideBuilder.js 接受 tables:{shapeId:FitResult} 精确照写 colW/rowH/fontSize，applyTableMergesXml 注入 gridSpan/rowSpan/hMerge/vMerge（须在 normalize 前；hMerge 占位经 normalize 清为 PPT 原生紧凑格式、vMerge 保留）；向导 UI 粘贴优先剪贴板 HTML、一键合并按钮+方向+撤销、fit 信息行（字号/缩字%/超出红字）、预览按 fit 渲染、提交 tables；MCP get_template 暴露 tableInfo + generate_slide/generate_deck 支持 tables + 新增 fit_table 工具。验证：tableFit 18 + table 15 + mcp 9 + 定向 86 + E2E table-fit 24/24 + golden-path 6/6 + tsc 0。**自查修复：parseTableHtml rowspan 后行 td 列定位 bug（错位与合并格重叠致内容丢失）**。**打包教训：slideBuilder.js 里的 expandGrid 必须内联（不能 require 跨目录的 addin .ts——pkg 只打包入口依赖图，exe 运行时会 MODULE_NOT_FOUND；已实测 exe 冒烟生成合併表通过）**。部署：重跑 安装.bat 重启 PPT；MCP 用 启动MCP.bat。
* **MCP 补全（P1-D 收官，2026-08-24，用户指定）**：MCP 增加图片位生成 + 套版工具 + 协议升级。①后端 /api/ai/generate 支持 images 参数（{图片位id: dataURL 或 http(s) URL}，URL 经 downloadStore 安全下载转 dataURL 后嵌入；非法值 400 明确报错）；②新增 POST /api/ai/generate-deck（pages 逐页加载模板支持 templateVersion 固定版本 → buildDeckBase64 多页 → 待写队列，write_slide 一次整份写入；逐页失败隔离）；③MCP 工具 10→13：generate_slide 加 images 透传 + 新增 list_decks / get_deck / generate_deck；④协议升级：capabilities 声明 tools/prompts/logging，实现 prompts/list + prompts/get（预设指令模板 ppt-gen-page / ppt-gen-deck，把「选模板→生成→写入」流程固化为提示词）；⑤文档：README MCP 章节 + 安装说明.md 新增「MCP 接入」章节。验证：mcp.test.js 6/6（新增图片位嵌入 png 断言/套版 pending 多页 base64/prompts）、ai.test.js 8/8（含 images 直通+非法值+generate-deck 多页+单页失败隔离）、回归 54 项全过、tsc 0 错。**注意：MCP 改动需用 启动MCP.bat 重启；前端零改动。**
* **字体/图片样式完整方案（2026-08-20）**：字体 font=latin + eaFont=ea 双字段（回读含占位符/母版继承链，生成端改写 a:ea）；图片样式保存 p:pic 完整 spPr（除 xfrm）+ blip 效果 + srcRect，生成时整体回写。⚠️ 旧模板 JSON 是快照，需重新保存才带新字段。
* **模板语义层（Semantic Template Layer，2026-08-20）**：TemplateShape 可选语义字段 semanticRole（title/subtitle/body/bullet/caption/label/data/conclusion/page\_number/tag/other）+ contentType/required/maxChars/maxLines/minChars/preferredLength/generationInstruction；后端 semantic.js 归一化 + 生成时把约束拼进 AI 提示词（/chat/completions 协议不变）；保存页语义角色下拉（按字号启发）+ ⚙ 约束面板；MCP 透传。schemaVersion 仍为 1（additive 可选字段，旧模板零影响）。
* **PPT Context（2026-08-20）**：MCP 4 工具（get\_presentation\_context / current\_slide / slide / inspect）经后端 /api/context → pptContext.js → context.ps1（COM 连正在运行的 PowerPoint，不依赖任务窗格）；稳定 ID 用 SlideID；绝不返回 PPTX 二进制。**坑：PS 5.1 对无 BOM 的 UTF-8 脚本按 GBK 解码，.ps1 必须带 UTF-8 BOM。**
* **模板版本控制（Template Versioning，2026-08-20）**：versions/ 子目录 v1/v2…，template.json 始终是「当前版本」镜像（旧代码零改动）；保存同名默认新建版本、绝不静默覆盖；恢复/删除版本 API + 模板库 vN 徽标弹窗；原子写（tmp+rename）+ O\_EXCL 防双写；旧模板读取时惰性迁移 v1。**教训：本模块必须保持同步 API（曾有 withLock 异步化 break 全部调用方）。**
* **图片搜索/下载重构（Image Provider，2026-08-20）**：imageService → providers/registry → Baidu/Bing 页面解析隔离，统一 ImageResult；下载安全（downloadSecurity.js）：SSRF 防护、MIME 白名单（拒 SVG）、魔数校验、20MB 上限、60s 超时、原子写、本地缓存。修复了原下载接口的 SSRF 漏洞。
* **AI 自动模板分析与自动标注（2026-08-20，规则引擎已被 08-23 版本重写取代）**：读取页面后自动推荐角色；规则分类器（analyze.ts）+ AI 增强（严格 JSON，失败完全回退规则）；config.analyze.enabled 总开关（默认关）；UI 推荐徽章 + 一键接受高置信建议（阈值后改 0.8）；用户手动改过的不覆盖。详见「元素规则分析优化」。
* **模板图片位人工裁剪（2026-08-20）**：向导选图后自动弹裁剪器（cropEditor.ts + cropMath.ts）；比例固定为图片位宽高比、实时预览、输出长边 cap 2048；裁剪后清该位旧 srcRect（圆角/阴影保留），后端零改动。
* **套版 Deck（P0-A，2026-08-20）**：整份/多页生成——套版 = 有序模板引用序列（固定版本，不内嵌快照）；deckStore + routes/decks + buildDeckBase64（单实例逐页 addSlide，跨页 slideSize 必须一致）；前端「套版」页 + deckWizard 逐页填参 → 一次插入整份；build 失败逐页隔离。
* **表格支持（P0-B/P0-C，2026-08-20）**：表格作为「表格位」保存（readStyles.js 从 zip XML 解析行列/合并/样式），生成用 pptxgenjs addTable 重建；生成向导「表格数据」步骤（手动逐格 / 粘贴 CSV / AI 生成三方式，上限 40×20）；样式保真：colW 归一化、逐格边框 a:lnT/R/B/L、tcMar、tblPr（表头/隔行变色）生成后回写；**Office.js 无表格 API，走 XML 回读路线；pptxgenjs 合并格会膨胀列数 → MVP 不传合并**。修复表格重复显示（按 bounds 命中原位转换）。
* **界面小字收纳 tooltip（2026-08-20）**：lib/tooltip.ts 带圈问号悬停气泡（事件委托挂一次即可）；规则：完整说明收进气泡、原解释段删除；动态状态/错误提示保持可见。
* **修复：AI 分析配置勾选保存无效（2026-08-20）**：config.js 路由 PUT/GET 漏配 analyze 键。**教训：改配置路由时所有键必须成对处理（GET 返回 + PUT 合并）。**
* **自动翻译副标题（2026-08-20）**：TemplateShape 加 translate / translateSource（'theme' = 全局主题）；语义角色选「副标题」自动勾选；生成顺序先普通文本位再翻译位；原文为空跳过。
* **MCP BUG 清单修复（2026-08-20，10 项）**：pkg 打包 ps1 必须 `bin: src/index.js` + `pkg .`（CLI 模式 pkg.assets 不生效）；启动MCP.bat 路径修复；getTemplate folder 参数（未指定时全局扫描兜底，防重名歧义）；PS 输出加 UTF8；待写队列惰性清理 + DELETE /api/ai/pending；新增 save\_template 工具（from-slide 保存）；ping 必须响应 result:{}；stdin 消息 msgQueue 串行防乱序；translateSource 改存形状索引再映射 shpN。
* **P2-E 安全加固（2026-08-21）**：API Key 用 Windows DPAPI 加密落盘（security.js 复用 powershell ProtectedData，非 Windows 回退明文）；后端随机端口（3788 被占 +1，写 runtime.json）+ 每次启动换一次性 token（X-Auth-Token 鉴权，缩略图白名单放行）；CORS 收紧为本机回环。⚠️ 首次部署后配置页需重新保存一次 API Key（明文自动迁移 dpapi: 加密）。
* **P2-F 模板回收站（2026-08-21）**：删除移入 文档/PPT模板库/.回收站/（保留分类层级与版本），可恢复 / 彻底删除 / 清空；全局扫描跳过回收站；entryId 防穿越。
* **实测回归修复（2026-08-21，真实使用发现 5 项）**：**部署不一致根因 = 改后端后没先 `npm run build:exe` 就运行 安装.bat**（教训：install.bat 不负责 build）；insert.ps1 补 BOM；**InsertFromFile 的 Index 是「之后插入」语义、合法 0..Count，末尾追加传 Count**；**pkg 环境判定必须用 typeof process.pkg !== 'undefined'（existsSync 对 VFS 资产也返回 true，exe 模式此前 context 从未真正可用）**；待写队列残留加清空入口。
* **保存/生成体验优化（P1-C，2026-08-21）**：保存页读页后自动应用高置信规则推荐（≥0.7，后改 0.8）；MCP save\_template 规则预标（不再一刀切 fixed）；固定图片位随模板保存（readStyles 从 zip 提取图片本体 dataURL）；重存同名模板自动继承角色标记（位置重叠 + 尺寸相近 + 名称相同贪心匹配，用户手改不覆盖）。
* **保存/使用流程简化（2026-08-21，用户指定）**：保存页 4 角色→2 角色（文字/图片，表格不变）；固定改勾选制（几何/无文本默认固定）；手动变量位角色移除（并入文字步骤）；向导步骤重排为 文字→图片(→表格)→预览写入；多图片位每个分别选图（本地/搜图/拖拽）。
* **部署一致性：版本统一 + 发布流水线（2026-08-22）**：scripts/release.js 自动版本 YYYY.MM.DD.NN → 写 version.js/release.json → 测试 → tsc/build → pkg exe → manifest Version → 产物清单；vite define 注入 `\_\_APP\_VERSION\_\_`；/api/version 富诊断（前后端版本/大小/MCP 心跳/端口）；MCP serverInfo 版本 + 30s 心跳文件；**安装.bat 升级为 Build→Package→Install→Verify**（verify-install.ps1 自检：运行版本 == release 版本，关键项 FAIL 才失败）；设置页「关于/系统诊断」卡片 + 启动自检横幅。本机 node --test 极慢（完整套件 10 分钟+），支持 --skip-tests 应急。
* **安装.bat 部署流水线 BUG 修复（2026-08-22）**：**规则：.bat 必须 CRLF 行尾**（LF 中文批处理会错乱、窗口一闪就关）；bat 变量拼接命令参数必须显式加空格；重写 bat 不要全局 replace 改关键词（会破坏 --no-pause）；verify 误报修复（MCP 未运行改 Info，仅关键项失败才 exit 1）；默认 --skip-tests；timeout 换 ping 延迟；安装前删除旧 runtime.json。
* **黄金路径 E2E（2026-08-23，发布门禁）**：server/e2e/golden-path.js 6 条黄金路径（单页模板 / 多图片 / 表格 / 套版 / 旧模板兼容 / 性能冒烟），文件级断言 + COM 打开检查（open-check.ps1，无 COM 自动 SKIP，--no-com 强制跳过）；接入 release.js 默认必跑（--skip-e2e 逃生舱）。
* **性能架构（2026-08-23，针对大 PPT）**：核心收益 = 生成后处理链合并（4 个 XML 纯函数 1 次 JSZip 往返，实测提速 53\~57%）；perf 指标系统（BUDGETS 14 项 + 环形缓冲 300 条 + /api/perf/stats 需 token，只写日志不展示界面）；读大文档跳过全文档扫描（indexValid）；模板库 mtime 缓存；搜图服务端缓存（深拷贝防污染）。
* **UI/UX 视觉与信息架构重构（2026-08-23，纯前端）**：Design Tokens（唯一主色深蓝 #264478，蓝只用于「值得注意」）；轻量顶部导航（2px 主色指示条）；按钮三体系（primary/secondary/ghost）；保存模板页三模块层级 + 元素卡片 + 折叠约束；AI 推荐轻提示条；全局去 emoji；**关键约束：JS 绑定选择器原样保留、未改 API/数据结构/后端逻辑**。
* **元素规则分析优化（2026-08-23，用户指定）**：先调研成熟方案（Office 官方占位符类型 ST\_PlaceholderType / python-pptx / PPTAgent 等）→ 决策表过目 → 用户 5 项决策：采纳 phType 官方语义优先 / 兜底改「信号不足」/ 自动应用阈值 0.7→0.8 / AI 增强一并做 / 总体照此实施。readSlide.ts ShapeInfo 加 `phType?`（Office.js placeholderFormat.type 归一化，最高优先级信号）；analyze.ts 重写为**分层规则引擎**（L0 phType 官方语义 → L1 版式/母版 → L2 几何 → L3 视觉 → L4 文本 → L5 重复 → L6 图注 → L7 兜底 R-FALLBACK 0.45「信号不足」，24 条规则带 ruleId + 信号 reason）；标题判定需「页面上部 30% + 视觉突出」；页眉页脚/日期/公司/网址 → 固定防 AI 改写；AI 摘要补 phType/align/color/lines；server 提示词加「占位符类型优先」。UI 零改动。技术备忘：run\_code 写大文件用 **pwsh here-string `@'...'@`（零转义）**；esbuild 可把前端 TS 单文件编译成 CJS 做规则引擎单测。

* **小工具·悬浮球（2026-08-23，用户指定，v3）**：任务窗格内可拖拽球（floatingBall.ts，pointer events 拖拽 + 点击展开菜单随球定位）。功能：一键去空格/段首加2空格/删除空行空段/分隔每段（文本4个，**保留每段字体字号写回**：整体 text 替换后 getSubstring 逐段恢复 font，段落数变化按索引映射）/选中标题/选中正文（**setSelected 高亮语义**，locateTitleRange/locateBodyRange 定位偏移，用实际文本计算避免 \r\n 错位）/复制格式（formatClipboard.ts：复制=读逐段 font+align 存内存、粘贴=循环映射应用到目标、清除=清空；**会话级内存不落盘**）/表格最佳适配/行列均分（表格组默认隐藏，菜单打开时 getSelectedTarget 检测选中表格才显示；PowerPoint DocumentSelectionChanged 不可靠 issue #5390 → 用打开时检测）。表格列宽行高用 TableColumn.width/TableRow.height（PowerPointApi 1.9 preview，不可用降级整体缩放）。新增 addin/src/tools/（textOps/tableOps/formatClipboard 纯函数 + selection 交互层）；单测 server/test/tools.test.js（Node 直 require .ts，纯函数模块不能有运行时相对 import）。
* **五界面 UI 精简（2026-08-23，用户指定，纯前端）**：①生成向导——质量检查只显示警告/异常类别（通过的✓不渲染，全过时整面板隐藏）；删「必填位为空」警告；**生成按钮移到主题下方 + position:sticky top:48px 常驻**（导航下吸附，滚动编辑文字/图片始终可见，方案由我定）；按钮去 ⚡；预览提示语收进「实时预览」旁圈问号。②模板库——删「共N个模板M个分类」；工具栏 🗑️ 左 + 1/2/4列右（space-between）；卡片「版本/🗑️」右下角（1列竖排右下、2/4列脚部右下，.lib-card-actions）。③保存模板——版本说明收进分类下的 details 折叠；删「回读状态」行（#enrich-status 移除，setStatus 变空操作）；保存/读取/文件夹选择改名；操作行 sticky 常驻。④套版——标题去「（整份生成）」、按钮改「新建」。⑤AI配置——去「（OpenAI 兼容）」、删 Bing API Key 输入+小字+未实现的 bing_api 选项（save 逻辑同步删）。验证：tsc 0 错、vite build 通过、tools 20/20。
* **AI 输出净化与一键输出模式（2026-08-23，用户指定，双层）**：①后端 `server/src/cleanText.js` 净化器（textService.generateText 统一出口，所有 /api/text/generate 返回自动生效）：始终剥离 Markdown 痕迹（###/**/`/代码围栏/引用/分隔线/链接图片）、规整换行（3+空行压2）；plain=true 额外剥列表编号（1. / - / ·）；maxChars/maxLines 截断（取模式限字数与模板语义约束的较严者，0=不限；向导路径自此也受模板 maxChars 约束）；表格 JSON 不受影响（调用方不传 plain/maxChars 即只做语法清洗）。②前端 `addin/src/lib/outputMode.ts`（localStorage `pptai.outputMode.v1` 记住，默认整段，touched 标记区分是否手动切换）：生成向导常驻贴条内一行 chips [整段][分点][精简]+[限字数▾]（互斥组合、点一下、不污染 UI）；模式同时变成给 AI 的提示词指令（buildOutputInstruction）与后端清洗规则（clean 字段，api.generateText 第4参）；未手动切换时语义角色=bullet 的文本位按分点（模板设计意图优先）；套版「生成本页文字」静默同享模式（每击 loadOutputMode 新鲜读）。验证：tsc 0 错、vite build 通过、后端 49+43 单测通过（含新增 cleanText.test.js 15 项）。微调：模式 chips 文案「纯文本」→「整段」；限字数下拉与 chips 同一行（select 全局 width:100% 会挤换行 → .om-limit width:auto + 同高圆角）；生成按钮去掉 inline padding 变扁（用全局 .primary 8px）；质量检查去掉「🔎 生成质量检查 / N 项警告」标题行只留警告内容。
* **UI 微调四连（2026-08-23，用户指定）**：①「✓ 版本一致」启动横幅不再显示（taskpane.ts 版本一致分支改 banner.remove()，仅版本不一致/后端异常才提示）；②输出模式 chips 从常驻贴条移到「文字」区块标题右侧（排版：文字 ▏整段/分点/精简/不限字数，去「输出」标签），贴条只留生成按钮；实时预览默认收起（previewCollapsed=true，门禁照常）；③保存模板 sticky 失效修复（根因：page-actions 包在 section.module 内，sticky 被限制在父卡片高度内 → 移到 .page 直接子级即相对整页吸附）；删 sticky-save 及读取按钮上下两条 border 横线；④删「页面」里「背景」行的「（已回读 N 个元素的精确样式）」后缀（enrichDone 逻辑保留，样式回读照常入库）。验证：tsc 0 错、vite build 通过。
* **UI 微调五连（2026-08-23，用户指定）**：①图片区——「把本地图片拖到这里直接上传…」提示收进「图片」题头问号（.img-dropzone 去文字保留拖拽，加虚线样式）；搜图结果只显示缩略图（删每张的来源/尺寸/标题文字行），每张右上角「＋」预览整张原图（showPreviewModal，不下载；点击图片本身仍下载选中）。②生成向导——「主题」改「全局提示词」，移到文字区块与生成按钮之间，details 可收起默认收起（.gp-adv summary 标题样式）；渲染顺序：模板→生成按钮(sticky常驻)→全局提示词(默认收起)→文字→图片→表格→预览。③保存模板——操作行去负 margin + 加 .module 同款 border/radius 矩形 + margin 上下等距（宽度与上下 UI 一致、读取后不贴下面区块）；删「先选中要保存的页面再读取」提示；背景勾选框 title 彻底删 + bgHint 默认说明清空；「图片」「跟随文档」字号内联统一 font-aux（label 必须 display 才让 margin 生效）；.pi-hint:empty 隐藏（删小字后 UI 自适应不留白）。④AI 配置——文本生成卡片 label 统一 display:block + margin-top:12px（inline 元素 margin-top 无效——此前"没做"根因）。验证：tsc 0 错、vite build 通过。
* **模板编辑 + 重名保护（2026-08-23，用户指定）**：保存重名不再静默新建版本——普通保存遇同名 → markInputError 红框红字报错（与未填名字同款显示，提示去模板库点「编辑」）；覆盖能力收敛为「编辑」功能：模板库每卡片 3 按钮同一行（编辑/版本/🗑️，.lib-row 原竖排改横排 + flex-wrap + info min-width:120px 防窄窗格挤压），点「编辑」→ sessionStorage 带 editTemplateId/editTemplateFolder 跳 #save；保存页载入模板完整信息（名称/分类/当前版本说明 changeNote/元素 TemplateShape→ShapeInfo 还原/role/fixed/提示词/形状类型/语义层/自选背景图还原，未调 runAnalysis 防规则覆盖已存角色）；编辑保存弹「是否确认覆盖该模板？确认/取消」（showModal，取消 return 留在界面继续编辑），确认 → saveTemplate updateCurrent:true 覆盖当前版本不建新版本；isSelf 按 existingTemplate.template.id===editId 判定（防编辑改名撞上另一同名模板误覆盖；极旧无 id 模板回退 name+folder）；编辑模式跳过同名标记继承（已完整载入角色，防覆盖用户改动）；editId 进入即消费（sessionStorage.removeItem）。api.ts TemplateDoc 补 id? 字段。验证：tsc 0 错、vite build 通过、后端 templateStore/versioning/recycle 31/31。
* **修复：加载项「无效」需手动重选（2026-08-24，用户指定）**：根因——release.js 每次部署改写 manifest.xml <Version>（YYYY.MM.DD.NN 递增），Office 在 HKCU\...\PowerPoint\Web Extension User MRU 记录已信任版本（规范化去前导零如 2026.8.24.2），清单版本一变即判「无效」需重新勾选启用。修复：**manifest <Version> 固定 2026.08.24.02**（与 MRU 已信任版本一致 → 无缝衔接零风险，比删除信任记录更稳：删记录反而可能触发重新信任）；release.js 删除 patchManifestVersion 调用与函数、更新头注释；MANIFEST_XML 常量仍用于产物 sha256 记录。版本号仍经 release.json + version.js 正常递增，前端/后端/设置页版本不受影响；verify-install 只比对 release.json vs 后端版本，不涉及 manifest。验证：跑 release.js --skip-tests --skip-e2e --skip-build 后 manifest.xml sha256 前后一致、Version 不变。若个别机器仍显示一次无效，手动重选一次后（MRU 更新为固定版本）永不再复发。