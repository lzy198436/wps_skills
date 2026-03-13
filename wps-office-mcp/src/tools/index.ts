/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - MCP工具汇总注册模块
 * 整合Excel、Word、PPT、Common的所有Tools
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 */

import { RegisteredTool } from '../types/tools';
import { excelTools } from './excel';
import { wordTools } from './word';
import { pptTools } from './ppt';
import { commonTools } from './common';

/**
 * 所有MCP Tools集合（共55个）
 *
 * Excel (27个):
 *   公式(3): set_formula, generate_formula, diagnose_formula
 *   数据(4): read_range, write_range, clean_data, remove_duplicates
 *   图表(2): create_chart, update_chart
 *   透视表(2): create_pivot_table, update_pivot_table
 *   工作表(8): create_sheet, delete_sheet, rename_sheet, copy_sheet, get_sheet_list, switch_sheet, move_sheet, get_selection
 *   格式化(8): set_cell_format, set_cell_style, set_border, set_number_format, merge_cells, unmerge_cells, set_column_width, set_row_height
 *
 * Word (9个):
 *   格式化(3): apply_style, set_font, generate_toc
 *   内容(2): insert_text, find_replace
 *   文档管理(4): get_open_documents, switch_document, open_document, get_document_text
 *
 * PPT (17个):
 *   幻灯片(3): add_slide, beautify, unify_font
 *   幻灯片操作(9): delete_slide, duplicate_slide, move_slide, get_slide_count, get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes
 *   演示文稿管理(5): create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation
 *
 * Common (2个):
 *   转换(2): convert_to_pdf, convert_format
 */
export const allTools: RegisteredTool[] = [
  ...excelTools,
  ...wordTools,
  ...pptTools,
  ...commonTools,
];

// 按应用类型分别导出
export { excelTools } from './excel';
export { wordTools } from './word';
export { pptTools } from './ppt';
export { commonTools } from './common';

// Excel相关导出
export {
  formulaTools,
  dataTools,
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
} from './excel';

// Word相关导出
export {
  formatTools,
  contentTools,
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
} from './word';

// PPT相关导出
export {
  slideTools,
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './ppt';

// Common相关导出
export {
  convertTools,
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './common';

/**
 * 获取所有Tool的数量
 */
export const getToolCount = (): number => allTools.length;

/**
 * 获取按应用分类的Tool数量
 */
export const getToolCountByApp = (): { excel: number; word: number; ppt: number; common: number } => ({
  excel: excelTools.length,
  word: wordTools.length,
  ppt: pptTools.length,
  common: commonTools.length,
});

export default allTools;
