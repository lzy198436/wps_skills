# WPS Skills - AI驱动的WPS Office智能助手

# WPS Skills - AI-Powered WPS Office Smart Assistant

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。
> Any feature or architecture updates must be reflected in documentation. This is part of our contract.

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Skills-Framework-purple?style=flat-square" alt="Skills Framework">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square" alt="Windows | macOS">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

---

## 项目简介 / Overview

WPS Skills 是一个让 AI 编程助手通过自然语言操控 WPS Office（Excel、Word、PPT）的开源工具集。安装后，你可以用中文或英文直接描述需求，AI 自动调用对应工具执行操作。

WPS Skills is an open-source toolkit that enables AI coding assistants to control WPS Office (Excel, Word, PPT) via natural language. After installation, describe your needs in natural language, and AI will automatically invoke the corresponding tools.

**一句话概括 / In a nutshell**：对你的 AI 助手说"帮我做个 VLOOKUP 查价格"，它就真的能在 WPS 里帮你做。Tell your AI assistant "Create a VLOOKUP to find prices", and it actually does it in WPS.

**架构 / Architecture**：Skills 层（自然语言指导）+ MCP Server 层（145 个工具）+ 执行层（macOS 224 / Windows 232 action）

**支持 / Supports**：Claude Code / Cursor / Augment / 其他 MCP 兼容 IDE

---

## 快速开始 / Quick Start

### 前提条件 / Prerequisites

- 已安装 [WPS Office](https://www.wps.cn/)（[English site](https://www.wps.com/)）
- Node.js >= 18.0.0

> **AI 助手请注意 / AI Assistant Notice**：请读取 [INSTALL.md](./INSTALL.md) 并按其中步骤执行。

### 安装 / Installation

#### macOS 一键安装 / macOS One-Click Install

```bash
cd /path/to/WPS_Skills
bash scripts/auto-install-mac.sh
```

#### Windows 一键安装 / Windows One-Click Install

```powershell
cd \path\to\WPS_Skills
powershell -ExecutionPolicy Bypass -File scripts\auto-install.ps1
```

#### 手动安装（通用）/ Manual Installation

1. **编译 MCP Server / Build MCP Server**：
   ```bash
   cd wps-office-mcp && npm install && npm run build
   ```

2. **注册 MCP Server / Register MCP Server**：
   ```bash
   claude mcp add wps-office node /path/to/WPS_Skills/wps-office-mcp/dist/index.js
   ```

3. **安装 Skills（仅 Claude Code）/ Install Skills (Claude Code only)**：
   ```bash
   PROJECT_DIR=/path/to/WPS_Skills
   mkdir -p ~/.claude/skills
   ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
   ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
   ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
   ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
   ```

4. **重启 Claude Code / IDE / Restart Claude Code / IDE**

### 配置 / Configuration

#### Claude Code（推荐 / Recommended）

安装后自动获得 Skills + MCP 双层能力，无需额外配置。
After installation, Skills + MCP dual-layer capabilities are loaded automatically. No extra configuration needed.

#### Cursor / Augment / 其他 IDE / Other IDEs

在对应 IDE 的 MCP 设置中添加 / Add to your IDE's MCP settings：

```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/path/to/WPS_Skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

> 这些 IDE 通过 MCP Server 获得 145 个工具能力，但不包含 Skills 层的自然语言工作流指导。
> These IDEs gain access to 145 MCP tools via the MCP Server, but without the Skills layer's natural language workflow guidance.

---

## 支持的 AI 工具 / Supported AI Tools

| AI 工具 / Tool | 集成方式 / Integration | 说明 / Description |
|---------|---------|------|
| **Claude Code** | Skills + MCP Server（原生支持 / native） | 启动时自动加载 `/wps-excel` `/wps-word` `/wps-ppt` `/wps-office` 四个 Skills，体验最完整 / Auto-loads 4 Skills on startup for the most complete experience |
| **Cursor** | MCP Server | 在 Cursor Settings > MCP 中添加 / Add in Cursor Settings > MCP |
| **Augment** | MCP Server | 通过 MCP 协议接入 / Connect via MCP protocol |
| **其他 MCP 兼容 IDE / Other MCP-compatible IDEs** | MCP Server | 任何支持 MCP 协议的 AI 工具均可接入 / Any MCP-compatible AI tool can connect |

---

## 功能概览 / Feature Overview

### 能力数量一览 / Capability Summary

| 层级 / Layer | 数量 / Count | 说明 / Description |
|------|------|------|
| **MCP 注册工具 / MCP Tools** | 145 个 | 133 个专业工具 + 12 个内置工具（含万能工具 `wps_execute_method`）/ 133 specialized + 12 built-in (incl. `wps_execute_method` universal tool) |
| **macOS 底层 action** | 224 个 | 通过 WPS JS 加载项实现 / Via WPS JS Add-in |
| **Windows 底层 action** | 232 个 | 通过 PowerShell COM 桥接实现（覆盖 macOS 全部 + 扩展）/ Via PowerShell COM bridge (covers all macOS + extensions) |

> **注意 / Note**：MCP 工具和底层 action 是不同概念。145 个 MCP 工具是 AI 直接可调用的接口；底层 action 是 WPS 加载项执行器支持的原子操作。通过 `wps_execute_method` 万能工具，AI 可间接访问全部底层 action。
> MCP tools and underlying actions are different concepts. 145 MCP tools are directly callable interfaces; actions are atomic operations supported by the WPS add-in executor. Through the `wps_execute_method` universal tool, AI can indirectly access all actions.

### Excel（65 个 MCP 工具 / 86 个 action）

| 分类 / Category | MCP工具数 / Tools | 能力 / Capabilities |
|------|----------|------|
| 公式功能 / Formulas | 6 | 设置公式/数组公式/诊断/重算/自动求和 / Set formula, array formula, diagnose, recalculate, auto-sum |
| 数据处理 / Data Processing | 12 | 读写/清洗/去重/排序/查找替换/批注/保护 / Read-write, clean, dedup, sort, find-replace, comments, protect |
| 图表与透视表 / Charts & Pivots | 4 | 创建/更新图表和透视表 / Create and update charts and pivot tables |
| 工作表操作 / Sheet Operations | 16 | 创建/删除/重命名/复制/冻结/命名区域/缩放 / Create, delete, rename, copy, freeze, named ranges, zoom |
| 格式美化 / Formatting | 10 | 样式/边框/数字格式/合并/自动调整 / Style, border, number format, merge, auto-fit |
| 工作簿管理 / Workbook Management | 10 | 打开/创建/切换/关闭/单元格读写 / Open, create, switch, close, cell read-write |
| 高级数据 / Advanced Data | 7 | 筛选/复制粘贴/填充序列/转置/分列/分类汇总 / Filter, copy-paste, fill series, transpose, text-to-columns, subtotal |

### Word（24 个 MCP 工具 / 25 个 action）

| 分类 / Category | MCP工具数 / Tools | 能力 / Capabilities |
|------|----------|------|
| 格式设置 / Formatting | 5 | 样式/字体/目录/书签/页面设置 / Style, font, TOC, bookmark, page setup |
| 内容操作 / Content Operations | 10 | 文本/查找替换/表格/段落/图片/批注/超链接/分页 / Text, find-replace, table, paragraph, image, comment, hyperlink, page break |
| 文档管理 / Document Management | 9 | 打开/创建/切换/获取全文/页眉/页脚/目录 / Open, create, switch, get full text, header, footer, TOC |

### PPT（42 个 MCP 工具 / 85 个 action）

| 分类 / Category | MCP工具数 / Tools | 能力 / Capabilities |
|------|----------|------|
| 幻灯片基础 / Slide Basics | 5 | 添加/美化/统一字体/字体颜色/对齐 / Add, beautify, unify font, font color, align |
| 幻灯片操作 / Slide Operations | 22 | 删除/复制/移动/布局/备注/形状/文本框/图片/动画/背景/切换 / Delete, duplicate, move, layout, notes, shapes, textbox, image, animation, background, transition |
| 演示文稿管理 / Presentation Management | 8 | 创建/打开/关闭/切换/主题/复制幻灯片 / Create, open, close, switch, theme, copy slide |
| 文本框操作 / Textbox Operations | 7 | 文本框增删改/标题/副标题/内容 / Add, delete, modify textboxes, title, subtitle, content |

### 通用工具 / Common Tools（2 个 MCP 工具 + 12 个内置工具）

| 分类 / Category | 工具数 / Tools | 能力 / Capabilities |
|------|--------|------|
| 格式转换 / Format Conversion | 2 | PDF 转换 / 格式互转 / PDF conversion, format interchange |
| 连接管理 / Connection Management | 1 | 检查 WPS 连接状态 / Check WPS connection status |
| 万能执行 / Universal Executor | 1 | `wps_execute_method` 可调用任意底层 action / Invokes any underlying action |
| 上下文获取 / Context Retrieval | 3 | 获取当前活动的文档/工作簿/演示文稿 / Get current active document, workbook, presentation |
| 数据缓存 / Data Cache | 4 | 缓存读写/列表/清除（用于大数据传输）/ Cache read, write, list, clear (for large data transfer) |
| 快捷操作 / Quick Operations | 1 | 快速插入文本 / Quick text insertion |

---

## 架构说明 / Architecture

采用 **Anthropic 官方标准的 MCP + Skills 双层架构** / Built on **Anthropic's official MCP + Skills dual-layer architecture**：

```
用户自然语言请求 / User natural language request
（如"帮我写个VLOOKUP公式查价格" / e.g. "Write a VLOOKUP to find prices"）
        |
+-- Skills 层 / Skills Layer（自然语言指令包 / NL Instruction Packages）--+
|  skills/wps-excel/SKILL.md   Excel 任务指导 / Excel task guidance       |
|  skills/wps-word/SKILL.md    Word 任务指导 / Word task guidance         |
|  skills/wps-ppt/SKILL.md     PPT 任务指导 / PPT task guidance           |
|  skills/wps-office/SKILL.md  跨应用协调 / Cross-app coordination        |
+----------------------------------------------------------------------+
        |
+-- MCP 层 / MCP Layer（工具能力 / Tool Capabilities）------------------+
|  wps-office-mcp/                                                     |
|  145 个 MCP 工具（133专业 + 12内置）/ 145 MCP tools (133 + 12 built-in)|
|  关键工具 / Key tool: wps_execute_method（万能调用 / universal invoke） |
+----------------------------------------------------------------------+
        |
+-- WPS 执行层 / WPS Execution Layer（加载项 / Add-in）-----------------+
|  macOS: HTTP轮询(端口58891) -> WPS加载项(JS API) / HTTP polling → JS  |
|         224 个 action                                                |
|  Windows: PowerShell COM 桥接 -> WPS Office / PS COM bridge          |
|         232 个 action                                                |
+----------------------------------------------------------------------+
```

### Skills 与 MCP 的关系 / Relationship Between Skills and MCP

| 层级 / Layer | 职责 / Role | 内容 / Content |
|------|------|------|
| **Skills** | 教 AI "怎么做"：工作流程、最佳实践、参数组合 / Teaches AI "how to do it": workflows, best practices, parameter combinations | 4 个 SKILL.md 文件，Claude Code 启动时自动加载 / 4 SKILL.md files, auto-loaded by Claude Code |
| **MCP** | 告诉 AI "能做什么"：可调用的工具和参数 / Tells AI "what can be done": callable tools and parameters | 145 个 MCP 工具，通过 MCP Server 暴露 / 145 MCP tools exposed via MCP Server |

---

## 项目目录结构 / Project Structure

```
WPS_Skills/
+-- wps-office-mcp/              # MCP Server（核心服务 / Core Service）
|   +-- src/                     # TypeScript 源码 / Source code
|   |   +-- tools/               # MCP 工具定义 / Tool definitions（133个，按 excel/word/ppt/common 分组）
|   |   +-- server/mcp-server.ts # MCP Server + 12个内置工具 / + 12 built-in tools
|   |   +-- client/wps-client.ts # 跨平台 WPS 通信客户端 / Cross-platform WPS client
|   |   +-- index.ts             # MCP Server 入口 / Entry point
|   +-- scripts/wps-com.ps1      # Windows COM 桥接脚本 / COM bridge script（232 action）
|   +-- dist/                    # 编译输出 / Compiled output（npm run build）
|   +-- package.json
+-- wps-claude-assistant/        # WPS 加载项 / Add-in（macOS）
|   +-- main.js                  # HTTP 轮询 + 224 action dispatch / HTTP polling + dispatch
|   +-- handlers/                # Excel/Word/PPT handler 实现 / implementations
|   +-- manifest.xml             # 加载项清单 / Add-in manifest
|   +-- ribbon.xml               # WPS 功能区配置 / Ribbon config
+-- wps-claude-addon/            # WPS 加载项 / Add-in（Windows）
|   +-- js/main.js               # 加载项逻辑 / Add-in logic
|   +-- manifest.xml             # 加载项清单 / Add-in manifest
|   +-- ribbon.xml               # 功能区配置 / Ribbon config
+-- skills/                      # Claude Skills 定义 / Definitions
|   +-- wps-excel/SKILL.md       # Excel 技能 / Excel skill
|   +-- wps-word/SKILL.md        # Word 技能 / Word skill
|   +-- wps-ppt/SKILL.md         # PPT 技能 / PPT skill
|   +-- wps-office/SKILL.md      # 跨应用协调技能 / Cross-app skill
+-- scripts/
|   +-- auto-install-mac.sh      # macOS 一键安装脚本 / One-click install
|   +-- auto-install.ps1         # Windows 一键安装脚本 / One-click install
+-- INSTALL.md                   # AI 自动安装指南 / AI auto-install guide
+-- README.md                    # 本文件 / This file
```

---

## 跨平台差异 / Cross-Platform Differences

| 项目 / Item | macOS | Windows |
|------|-------|---------|
| Action 数量 / Count | 224 个 | 232 个（含 8 个 Windows 扩展 / incl. 8 Windows extensions） |
| 执行方式 / Execution | HTTP 轮询 + WPS JS 加载项 / HTTP polling + WPS JS add-in | PowerShell COM 桥接 / PowerShell COM bridge |
| MCP 工具 / Tools | 145 个（共享 / shared） | 145 个（共享 / shared） |

---

## 系统要求 / System Requirements

| 项目 / Item | Windows | macOS |
|------|---------|-------|
| 操作系统 / OS | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 或更高 / 2019+ | Mac 版最新版 / Latest Mac version |
| Node.js | >= 18.0.0 | >= 18.0.0 |
| AI 工具 / AI Tool | Claude Code / Cursor / Augment 等 | Claude Code / Cursor / Augment 等 |

---

## 开发指南 / Development Guide

### 如何添加新工具 / How to Add New Tools

1. 在 `wps-office-mcp/src/tools/` 下对应分类目录中添加工具定义（TypeScript）
   Add tool definition in the corresponding category under `wps-office-mcp/src/tools/`

2. 在 `wps-office-mcp/src/tools/index.ts` 中注册新工具
   Register the new tool in `wps-office-mcp/src/tools/index.ts`

3. 更新对应的 `skills/wps-*/SKILL.md` 文档，添加新工具的使用说明
   Update the corresponding `skills/wps-*/SKILL.md` with usage instructions

4. 编译并测试：`cd wps-office-mcp && npm run build`
   Build and test

### 如何贡献代码 / How to Contribute

1. Fork 本仓库 / Fork this repo
2. 创建功能分支 / Create a feature branch：`git checkout -b feat/your-feature`
3. 提交变更 / Commit changes：遵循 Conventional Commits 规范
4. 提交 Pull Request

---

## 故障排除 / Troubleshooting

### Claude 助手选项卡未出现 / Claude Assistant Tab Not Showing

1. 确认加载项文件夹名称以 `_` 结尾 / Confirm add-in folder name ends with `_`
2. 确认 `publish.xml` 已正确配置，包含加载项注册条目 / Confirm `publish.xml` has add-in registration entries
3. 强制退出并重启 WPS Office / Force quit and restart WPS Office

### Skills 未加载（仅 Claude Code）/ Skills Not Loaded (Claude Code Only)

检查软链接是否存在 / Check if symlinks exist：

```bash
ls ~/.claude/skills/
```

如果为空，手动创建 / If empty, create manually：

```bash
PROJECT_DIR=/path/to/WPS_Skills
mkdir -p ~/.claude/skills
ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
```

然后重启 Claude Code（必须重启才能加载 Skills）。
Then restart Claude Code (restart is required to load Skills).

### MCP Server 连接失败 / MCP Server Connection Failed

1. 确认已执行编译 / Confirm build was executed：
   ```bash
   cd wps-office-mcp && npm install && npm run build
   ```
2. 运行 `claude mcp list` 检查 `wps-office` 是否已注册 / Check if `wps-office` is registered
3. 重启 Claude Code / IDE / Restart Claude Code / IDE

### WPS Office 未响应 / WPS Office Not Responding

**macOS**：
1. 确认 WPS Office 已启动并打开了文档 / Confirm WPS Office is running with a document open
2. 检查端口 58891 是否被占用 / Check if port 58891 is in use：`lsof -i :58891`
3. 重启 WPS Office 并重新打开文档 / Restart WPS Office and reopen document

**Windows**：
1. 以管理员权限运行 PowerShell / Run PowerShell as administrator
2. 确认 WPS COM 组件可用 / Confirm WPS COM component is available
3. 检查 PowerShell 执行策略 / Check execution policy：`Get-ExecutionPolicy`

### 工具调用报错 "method not found" / Tool Call Error "method not found"

1. 确认 WPS 加载项已正确安装 / Confirm WPS add-in is installed correctly
2. 确认 WPS 中已启用 Claude 加载项 / Confirm Claude add-in is enabled in WPS
3. 尝试使用 `wps_execute_method` 万能工具作为替代 / Try using `wps_execute_method` as alternative

---

## 许可证 / License

MIT License

## 开发者 / Developer

**lc2panda** - [GitHub](https://github.com/lc2panda)

---

> **AI读者须知 / AI Reader Notice**：你正在阅读 WPS Skills 项目。这是一个让 AI 编程助手（Claude Code、Cursor、Augment 等）能通过自然语言操控 WPS Office（Excel、Word、PPT）的工具集。安装后，你将获得 145 个 MCP 工具和 macOS 224 / Windows 232 个底层 action 能力。通过 `wps_execute_method` 万能工具可访问全部底层 action。
> You are reading the WPS Skills project. This toolkit enables AI assistants (Claude Code, Cursor, Augment, etc.) to control WPS Office via natural language. After installation, you gain 145 MCP tools and macOS 224 / Windows 232 underlying action capabilities. The `wps_execute_method` universal tool provides access to all actions.
