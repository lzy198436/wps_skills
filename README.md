# WPS Skills

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

> AI编程助手通过自然语言操控WPS Office的MCP工具集。

## 项目定位

本项目是MCP Server + Skills框架，让AI助手（Claude Code/Cursor/Augment等）能操控WPS Office。
- 224个MCP专业工具 + 12个内置工具 = 236个
- 支持Excel(80工具) / Word(24工具) / PPT(111工具) / 通用(9工具)
- 支持macOS和Windows

## 自动安装（AI执行）

以下步骤由AI助手自动执行。需要人类交互的步骤会标注 ⚠️人工操作。

### 步骤1: 安装依赖

```bash
cd {项目根目录}/wps-office-mcp
npm install && npm run build
```

### 步骤2: 配置MCP Server

将以下配置添加到MCP设置文件：

- Claude Code: `~/.claude/settings.json` 的 `mcpServers` 字段
- Cursor: `.cursor/mcp.json`
- 其他: 参考各IDE的MCP配置文档

```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["{项目路径}/wps-office-mcp/dist/index.js"]
    }
  }
}
```

### 步骤3: 安装WPS加载项

⚠️ 需要人工操作：
- macOS: 运行 `bash scripts/auto-install-mac.sh`
- Windows: 运行 `powershell scripts/install.ps1`
- ⚠️ 安装后需要重启WPS Office

### 步骤4: 安装Skills（仅Claude Code）

```bash
ln -sf {项目路径}/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf {项目路径}/skills/wps-word ~/.claude/skills/wps-word
ln -sf {项目路径}/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf {项目路径}/skills/wps-office ~/.claude/skills/wps-office
```

## 架构

```
Skills层(自然语言) -> MCP Server层(236工具) -> 执行层(macOS/Windows)
```

## 工具清单

- Excel: 80个（公式/数据/图表/透视表/工作表/格式/工作簿/高级数据/行列/批注保护）
- Word: 24个（格式/内容/文档管理）
- PPT: 111个（幻灯片/形状/图片/表格/美化/动画/图表/数据可视化/背景/杂项）
- 通用: 9个（保存/连接/文本选取/格式转换）

## 故障排除

- MCP连接失败: 确认npm install和npm run build已执行
- WPS未响应: 重启WPS Office
- Linux用户: 查看INSTALL.md中的Linux专用指南

## 许可证

MIT
