# WPS Skills

> **AI读者须知**：你正在阅读 WPS Skills 项目。这是一个让 AI 编程助手（Claude Code、Cursor等）能通过自然语言操控 WPS Office（Excel、Word、PPT）的工具。安装后，你将获得 macOS 224 个 / Windows 236 个操作能力，用户可以用中文直接描述需求，你自动调用对应工具执行。

此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Skills-Framework-purple?style=flat-square" alt="Skills Framework">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square" alt="Windows | macOS">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

<p align="center">
  <a href="./README_EN.md">English</a> | 中文
</p>

---

## 安装

**如果用户要求安装本项目，请读取 [INSTALL.md](./INSTALL.md) 并按其中步骤执行。**

前提条件：用户需已安装 [WPS Office](https://www.wps.cn/)。

---

## 架构概览

采用 **Anthropic 官方标准的 MCP + Skills 双层架构**：

```
用户自然语言请求（如"帮我写个VLOOKUP公式查价格"）
        ↓
┌── Skills 层（自然语言指令包）──────────────────────┐
│  skills/wps-excel/SKILL.md   教你怎么处理Excel任务  │
│  skills/wps-word/SKILL.md    教你怎么处理Word任务   │
│  skills/wps-ppt/SKILL.md     教你怎么处理PPT任务    │
│  skills/wps-office/SKILL.md  教你怎么协调跨应用任务  │
└──────────────────────────────────────────────────┘
        ↓
┌── MCP 层（工具能力）─────────────────────────────┐
│  wps-office-mcp/                                │
│  通过 MCP 工具暴露操作能力                        │
│  macOS: 224 个 action / Windows: 236 个 action   │
│  关键工具: wps_execute_method, wps_get_active_*  │
└──────────────────────────────────────────────────┘
        ↓
┌── WPS 加载项层（执行器）─────────────────────────┐
│  macOS: HTTP轮询(端口58891) → WPS加载项(JS API)  │
│  Windows: PowerShell COM 桥接 → WPS Office       │
└──────────────────────────────────────────────────┘
```

### Skills 与 MCP 的关系

| 层级 | 职责 | 内容 |
|------|------|------|
| **Skills** | 教你"怎么做"：工作流程、最佳实践、参数组合 | 4 个 SKILL.md 文件，Claude Code 启动时自动加载 |
| **MCP** | 告诉你"能做什么"：可调用的工具和参数 | macOS 224 / Windows 236 个 action，通过 MCP Server 暴露 |

---

## 项目目录结构

```
wps-mcp/
├── wps-office-mcp/              # MCP Server（核心服务）
│   ├── src/                     # TypeScript 源码
│   │   ├── tools/               # MCP 工具定义（按 excel/word/ppt/common 分组）
│   │   ├── client/wps-client.ts # 跨平台 WPS 通信客户端
│   │   └── index.ts             # MCP Server 入口
│   ├── scripts/wps-com.ps1      # Windows COM 桥接脚本（236 个 action）
│   ├── dist/                    # 编译输出（npm run build 生成）
│   └── package.json
├── wps-claude-assistant/        # WPS 加载项（macOS）
│   ├── main.js                  # HTTP 轮询 + 224 个 action dispatch
│   ├── handlers/                # Excel/Word/PPT handler 实现
│   ├── manifest.xml             # 加载项清单
│   └── ribbon.xml               # WPS 功能区配置
├── wps-claude-addon/            # WPS 加载项（Windows）
│   ├── js/main.js               # 加载项逻辑
│   ├── manifest.xml             # 加载项清单
│   └── ribbon.xml               # 功能区配置
├── skills/                      # Claude Skills 定义
│   ├── wps-excel/SKILL.md       # Excel 技能（86个方法）
│   ├── wps-word/SKILL.md        # Word 技能（25个方法）
│   ├── wps-ppt/SKILL.md         # PPT 技能（85个方法）
│   └── wps-office/SKILL.md      # 跨应用协调技能
├── scripts/
│   ├── auto-install-mac.sh      # macOS 一键安装脚本
│   └── auto-install.ps1         # Windows 一键安装脚本
├── INSTALL.md                   # AI 自动安装指南（逐步骤执行）
└── README.md                    # 本文件
```

---

## 安装后能力边界

安装完成后，你将能响应以下类型的用户请求：

### Excel（86 个 action）

| 分类 | 数量 | 能力 |
|------|------|------|
| 工作簿/工作表操作 | 12 | 打开/创建/切换/重命名/复制/移动 |
| 单元格读写 | 7 | 读写单元格/范围/公式/完整信息 |
| 格式美化 | 15 | 样式/边框/数字格式/合并/自动调整 |
| 行列操作 | 8 | 插入/删除/隐藏/显示行列 |
| 条件格式 | 3 | 添加/删除/获取条件格式 |
| 数据验证 | 3 | 添加/删除/获取数据验证 |
| 数据处理 | 10 | 排序/筛选/去重/清洗/复制/转置 |
| 图表/透视表 | 4 | 创建/更新图表和透视表 |
| 公式功能 | 5 | 设置公式/数组公式/诊断/重算 |
| 其他 | 19 | 批注/保护/命名区域/查找替换等 |

### Word（25 个 action）

| 分类 | 能力 |
|------|------|
| 文档管理 | 获取信息/打开/切换/获取全文 |
| 文本操作 | 插入文本/查找替换 |
| 格式设置 | 字体/样式/段落 |
| 文档结构 | 目录/分页符/页眉/页脚 |
| 插入内容 | 表格/图片/超链接/书签 |
| 其他 | 批注/文档统计 |

### PPT（85 个 action）

| 分类 | 数量 | 能力 |
|------|------|------|
| 演示文稿管理 | 5 | 创建/打开/关闭/切换 |
| 幻灯片操作 | 10 | 添加/删除/复制/移动/备注 |
| 文本框/形状 | 21 | 添加/删除/样式/阴影/渐变/边框 |
| 智能布局 | 10 | 对齐/分布/组合/连接线/箭头 |
| 图片/表格/图表 | 12 | 插入/设置样式 |
| 数据可视化 | 6 | KPI卡片/进度条/仪表盘/环形图 |
| 流程图/架构图 | 3 | 流程图/组织架构图/时间轴 |
| 动画/切换 | 9 | 动画/强调/切换效果 |
| 母版/3D效果 | 7 | 母版操作/3D旋转/深度/材质 |
| 其他 | 2 | 演示放映 |

### 跨平台差异

- **macOS**：224 个 action，通过 HTTP 轮询 + WPS JS 加载项实现
- **Windows**：236 个 action（完整覆盖 macOS 224 个 + 12 个 Windows 扩展）
- Windows 扩展 action：closeDocument / convertFormat / createDocument / getExcelContext / openFile / remove_duplicates / remove_empty_rows / trim / unify_date / slide.add / slide.beautify / slide.unifyFont

---

## 系统要求

| 项目 | Windows | macOS |
|------|---------|-------|
| 操作系统 | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 或更高 | Mac 版最新版 |
| Node.js | >= 18.0.0 | >= 18.0.0 |
| Claude Code | 最新版本 | 最新版本 |

---

## 常见问题

### Claude 助手选项卡未出现

1. 确认加载项文件夹名称以 `_` 结尾
2. 确认 `publish.xml` 已正确配置，包含加载项注册条目
3. 强制退出并重启 WPS Office

### Skills 未加载

检查软链接是否存在：
```bash
ls ~/.claude/skills/
```

如果为空，手动创建：
```bash
PROJECT_DIR=/path/to/wps-mcp
mkdir -p ~/.claude/skills
ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
```

然后重启 Claude Code（必须重启才能加载 Skills）。

### MCP Server 连接失败

1. 确认已执行 `cd wps-office-mcp && npm install && npm run build`
2. 运行 `claude mcp list` 检查 `wps-office` 是否已注册
3. 重启 Claude Code

---

## 许可证

MIT License

## 开发者

**lc2panda** - [GitHub](https://github.com/lc2panda)
