# WPS Skills 研发 TODO

> 香草上校维护 | 最后更新：2026-03-14 11:12 +08:00
> PM职责：盯进度、验收、决策推进。不插手具体实现。

---

## ✅ 已完成

### 基础修补波（T01-T11）
- [x] T01 Skills审计修复（Excel）
- [x] T02 Skills审计修复（Word）  
- [x] T03 Skills审计修复（PPT）
- [x] T04 MCP工具层审计（Excel）
- [x] T05 INSTALL.md + README（已跳过商业化，面向AI读者优化完成）
- [x] T06 MCP工具层审计（Word/PPT）+ 编译验证
- [x] T07 错误处理统一化
- [x] T08 跨平台一致性核查 + INSTALL.md重写
- [x] T09 文档全面更新（wps-office SKILL.md）
- [x] T10 main.js params校验修复
- [x] T11 wps-client.ts 6项修复

### 文档优化
- [x] README.md 面向AI读者重写
- [x] INSTALL.md 面向AI读者重写 + macOS沙盒权限已知问题
- [x] README_EN.md 同步中文版

### MCP工具扩展（T12-T14，共6轮）
- [x] 当前工具数：**99个**（初始55个，+44个）
- [x] 覆盖率：99/224 = **44.2%**
- [x] 编译：✅ tsc 零错误
- [x] 修复 formula.ts + content.ts 注册位置错误

---

## 🔴 进行中

### 精确统计与验收
- [ ] 精确统计各模块当前工具数
- [ ] 更新 MCP覆盖率总报告至99工具/44.2%

---

## 🟡 待推进（按优先级）

### T15：第七轮MCP扩展（目标突破110工具/49%）
- [ ] Excel: hide_column, set_print_settings, auto_sum, protect_workbook
- [ ] Word: insert_comment, accept_all_changes, set_text_color, insert_list
- [ ] PPT: set_shape_fill, group_objects, set_slide_size, add_speaker_notes

### T16：Skills文件同步更新
- [ ] wps-excel/SKILL.md 更新工具列表（反映99个MCP工具）
- [ ] wps-word/SKILL.md 更新工具列表
- [ ] wps-ppt/SKILL.md 更新工具列表
- [ ] wps-office/SKILL.md 更新跨应用工具列表

### T17：Windows COM层对齐验证
- [ ] 检查新增MCP工具对应的 wps-com.ps1 action是否存在
- [ ] 补充缺失的Windows action

### T18：最终验收 + push
- [ ] npm run build 最终验证
- [ ] 安装流程完整跑通验证
- [ ] git push 到 GitHub

---

## 📋 作战规则

1. 每个Agent任务开头加 /compact 压缩上下文
2. 单次任务控制2-3个工具，防止格式错误
3. 编码Agent完成后必须跟编译验证Agent
4. 发现重复工具跳过，不强行追加
5. 阶段性完成后主动汇报指挥官

---

## 🔧 技术参数

- 项目路径：`/Users/panda/Downloads/download/WPS_Skills`
- MCP Server：`wps-office-mcp/dist/index.js`
- main.js actions：224个（macOS）/ 236个（Windows）
- Skills目录：`~/.claude/skills/`（软链接）
- WPS加载项：`~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/`
