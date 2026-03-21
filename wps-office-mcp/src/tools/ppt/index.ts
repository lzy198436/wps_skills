/**
 * Input: PPT 工具定义
 * Output: PPT 工具注册数组
 * Pos: PPT Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT Tools入口 - PPT工具汇总模块
 *
 * 整合所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font
 * - 幻灯片操作Tools: delete_slide, duplicate_slide, move_slide, get_slide_count,
 *   get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes,
 *   add_shape, set_shape_style, add_textbox, set_slide_title, insert_image, set_shape_text,
 *   set_animation, set_background, set_slide_size
 * - 演示文稿管理Tools: create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation, set_slide_theme, copy_slide, insert_slide_image
 */

import { RegisteredTool } from '../../types/tools';
import { slideTools } from './slide';
import { slideOpsTools } from './slide-ops';
import { presentationTools } from './presentation';
import { textboxTools } from './textbox';

/**
 * 所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font
 * - 幻灯片操作Tools: delete_slide, duplicate_slide, move_slide, get_slide_count,
 *   get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes
 * - 演示文稿管理Tools: create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation
 */
export const pptTools: RegisteredTool[] = [
  ...slideTools,
  ...slideOpsTools,
  ...presentationTools,
  ...textboxTools,
];

// 分别导出，方便按需使用
export { slideTools } from './slide';
export { slideOpsTools } from './slide-ops';
export { presentationTools } from './presentation';
export { textboxTools } from './textbox';

// 导出单独的定义和处理器，方便测试
export {
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './slide';

export {
  deleteSlideDefinition,
  deleteSlideHandler,
  duplicateSlideDefinition,
  duplicateSlideHandler,
  moveSlideDefinition,
  moveSlideHandler,
  getSlideCountDefinition,
  getSlideCountHandler,
  getSlideInfoDefinition,
  getSlideInfoHandler,
  switchSlideDefinition,
  switchSlideHandler,
  setSlideLayoutDefinition,
  setSlideLayoutHandler,
  getSlideNotesDefinition,
  getSlideNotesHandler,
  setSlideNotesDefinition,
  setSlideNotesHandler,
  addShapeDefinition,
  addShapeHandler,
  setShapeStyleDefinition,
  setShapeStyleHandler,
  addTextboxDefinition,
  addTextboxHandler,
  setSlideTitleDefinition,
  setSlideTitleHandler,
  insertImageDefinition,
  insertImageHandler,
  setShapeTextDefinition,
  setShapeTextHandler,
  setAnimationDefinition,
  setAnimationHandler,
  setBackgroundDefinition,
  setBackgroundHandler,
  setSlideSizeDefinition,
  setSlideSizeHandler,
} from './slide-ops';

export {
  createPresentationDefinition,
  createPresentationHandler,
  openPresentationDefinition,
  openPresentationHandler,
  closePresentationDefinition,
  closePresentationHandler,
  getOpenPresentationsDefinition,
  getOpenPresentationsHandler,
  switchPresentationDefinition,
  switchPresentationHandler,
  setSlideThemeDefinition,
  setSlideThemeHandler,
  copySlideDefinition,
  copySlideHandler,
  insertSlideImageDefinition,
  insertSlideImageHandler,
} from './presentation';

export default pptTools;
