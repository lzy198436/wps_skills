# Windows 功能扩充开发计划（对齐 macOS 196 方法）

## 目标与范围
- **目标**：Windows 通过 COM 实现与 macOS `main.js` 的 **196 个 action** 全量对齐。
- **覆盖**：Excel / Word / PPT / Common 全部能力。
- **文档**：按要求同步更新能力说明与 API 文档。

## 现状与关键位置（参考）
- **macOS action 分发**：`wps-claude-assistant/main.js:148`（`switch (cmd.action)` 共 196 个 `case`）
- **Windows action 分发**：`wps-office-mcp/scripts/wps-com.ps1:36`（`switch ($Action)`，约 25 个）
- **跨平台调用**：`wps-office-mcp/src/client/wps-client.ts:41-123`（`execMacPoll`/`execPowerShell`）
- **旧 API 映射**：`wps-office-mcp/src/client/wps-client.ts:170-178`（`actionMap`）
- **Tool 总入口**：`wps-office-mcp/src/tools/index.ts:49-54`
- **macOS 参考实现**：`wps-claude-assistant/handlers/*.js`（excel/word/ppt）

## 实施步骤
### 1. 生成 action 对照表（Mac vs Windows）
- 从 `main.js` 提取 196 个 action，按 **Common/Excel/Word/PPT** 分类。
- 从 `wps-com.ps1` 提取 Windows 已实现 action。
- 形成 **缺口清单**（action 级别），作为后续实现的唯一来源。

### 2. 扩展 Windows PowerShell COM 实现（核心）
在 `wps-office-mcp/scripts/wps-com.ps1` 中按分类补齐 action，保持与 macOS 返回结构一致（`{ success, data, error }`）。

**Excel（示例缺口方向）**
- 上下文：`getExcelContext`（参考 `handlers/excel-handler.js:119`）
- 图表：`createChart`/`updateChart` 扩展配置（类型、标题、图例、数据标签）
- 透视表：`createPivotTable`/`updatePivotTable`
- 数据清洗：`cleanData`、`removeDuplicates`
- 其它：排序/筛选、范围读写增强等

**Word**
- 文档管理：`getOpenDocuments`/`openDocument`/`switchDocument`
- 内容与格式：`applyStyle`、`generateTOC`、`insertTable` 等
- 与 macOS 行为对齐：文本长度限制、结果结构

**PPT（最大增量）**
- 幻灯片：增删改、复制、移动、布局、备注
- 文本框/标题：获取、设置、样式
- 形状/图片/表格/图表/动画：与 macOS case 集合对齐
- 美化与高级组件：配色、自动排版、KPI 卡片、流程图等

**Common**
- `convert_to_pdf`、`convert_format`（基于 COM 的 `SaveAs`/`ExportAsFixedFormat`）

> 实现时以 macOS handler 为语义对齐基准，Windows COM 负责具体落地。

### 3. 补齐 WpsClient 与类型
- `actionMap` 扩展：如存在旧 API 名称，需要映射到新增 action。
- `src/types/wps.ts`（若有）补齐响应结构类型，保持 TS 编译通过。

### 4. Tool 层对齐与扩展
在 `wps-office-mcp/src/tools/**` 中补齐工具定义与 handler：
- **Excel**：`excel/formula.ts`、`excel/data.ts`、`excel/chart.ts`、`excel/pivot.ts`
- **Word**：`word/format.ts`、`word/content.ts`
- **PPT**：`ppt/slide.ts` 及新增需要的工具文件
- **Common**：`common/convert.ts`

要求：工具名继续保持 `wps_{app}_{action}` 命名，并调用 `wpsClient.invokeAction` 对齐新增 action。

### 5. 文档同步更新（按要求）
- `README.md`：Windows 功能与方法数对齐声明，移除“仅基础功能”描述。
- `docs/API参考.md`：补齐新增工具或标注 Windows 已对齐。
- `skills/*/SKILL.md`：同步工具列表与能力描述。
- `docs/开发规划.md`：更新 Windows 相关完成项。
- （可选）增加 **Windows/Mac capability 对照表**（放 README 或 docs）。

## 验证方案（Windows）
- 运行 MCP Server，分别在 Excel/Word/PPT 打开示例文档。
- 逐类抽样验证：
  - **Excel**：读写范围、公式、图表、透视表
  - **Word**：插入文本、样式、目录、表格
  - **PPT**：新增幻灯片、文本框、形状、图表/动画、美化
  - **Common**：PDF/格式转换
- 检查 PowerShell 输出 JSON 可解析（避免 `Invalid JSON output`）。
- 回归：Mac 端 `ping`/基础调用不受影响。

## 备注
- 目标是 **全量对齐**，如发现 COM 不可达能力，需在实现前确认替代方案或显式标注限制，并同步文档说明。
