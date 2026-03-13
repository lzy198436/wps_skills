/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 幻灯片操作Tools - 删除/复制/移动/查询/切换/布局/备注
 *
 * 包含：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_duplicate_slide: 复制幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_notes: 设置幻灯片备注
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

// ============================================================
// 1. wps_ppt_delete_slide - 删除指定幻灯片
// ============================================================

export const deleteSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_slide',
  description: `删除指定的幻灯片。

使用场景：
- "删除第3页幻灯片"
- "把最后一页删掉"
- "移除多余的页面"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '要删除的幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const deleteSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'deleteSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片删除成功！\n已删除: 第 ${slideIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 2. wps_ppt_duplicate_slide - 复制幻灯片
// ============================================================

export const duplicateSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_duplicate_slide',
  description: `复制指定的幻灯片，在其后插入副本。

使用场景：
- "复制第2页幻灯片"
- "把这页再复制一份"
- "克隆当前幻灯片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '要复制的幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const duplicateSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      newSlideIndex: number;
    }>(
      'duplicateSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片复制成功！\n源幻灯片: 第 ${slideIndex} 页\n副本位置: 第 ${response.data.newSlideIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `复制幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 3. wps_ppt_move_slide - 移动幻灯片到指定位置
// ============================================================

export const moveSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_move_slide',
  description: `移动幻灯片到指定位置。

使用场景：
- "把第5页移到第2页"
- "把最后一页移到开头"
- "调整幻灯片顺序"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      fromIndex: {
        type: 'number',
        description: '原位置索引（从1开始）',
      },
      toIndex: {
        type: 'number',
        description: '目标位置索引（从1开始）',
      },
    },
    required: ['fromIndex', 'toIndex'],
  },
};

export const moveSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { fromIndex, toIndex } = args as {
    fromIndex: number;
    toIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'moveSlide',
      { fromIndex, toIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片移动成功！\n从第 ${fromIndex} 页 → 移到第 ${toIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移动幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移动幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 4. wps_ppt_get_slide_count - 获取幻灯片总数
// ============================================================

export const getSlideCountDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_count',
  description: `获取演示文稿中的幻灯片总数。

使用场景：
- "一共有多少页幻灯片"
- "PPT有几页"
- "查看幻灯片数量"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSlideCountHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count: number;
    }>(
      'getSlideCount',
      {},
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `当前演示文稿共有 ${response.data.count} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片总数失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片总数出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 5. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
// ============================================================

export const getSlideInfoDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_info',
  description: `获取指定幻灯片的详细信息，包括布局、元素列表等。

使用场景：
- "查看第3页的信息"
- "这页有什么内容"
- "获取幻灯片详情"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideInfoHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      layout: string;
      shapesCount: number;
      shapes: Array<{
        name: string;
        type: string;
        text?: string;
      }>;
    }>(
      'getSlideInfo',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const info = response.data;
      let output = `第 ${info.slideIndex} 页幻灯片信息：\n`;
      output += `布局: ${info.layout}\n`;
      output += `元素数量: ${info.shapesCount}\n`;

      if (info.shapes && info.shapes.length > 0) {
        output += `\n元素列表：\n`;
        info.shapes.forEach((shape, i) => {
          output += `  ${i + 1}. [${shape.type}] ${shape.name}`;
          if (shape.text) {
            output += ` - "${shape.text.substring(0, 50)}${shape.text.length > 50 ? '...' : ''}"`;
          }
          output += '\n';
        });
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片信息失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片信息出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 6. wps_ppt_switch_slide - 切换到指定幻灯片
// ============================================================

export const switchSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_switch_slide',
  description: `切换到指定的幻灯片页面。

使用场景：
- "切换到第5页"
- "跳到最后一页"
- "显示第1页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '目标幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const switchSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到第 ${slideIndex} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 7. wps_ppt_set_slide_layout - 设置幻灯片版式布局
// ============================================================

export const setSlideLayoutDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_layout',
  description: `设置幻灯片的版式布局。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局
- section_header: 节标题
- title_only: 仅标题

使用场景：
- "把这页改成空白布局"
- "设置为两栏内容"
- "改成标题页版式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      layout: {
        type: 'string',
        description: '布局名称',
      },
    },
    required: ['slideIndex', 'layout'],
  },
};

export const setSlideLayoutHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, layout } = args as {
    slideIndex: number;
    layout: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideLayout',
      { slideIndex, layout },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const layoutName: Record<string, string> = {
        title: '标题页',
        title_content: '标题+内容',
        blank: '空白页',
        two_column: '两栏内容',
        comparison: '对比布局',
        section_header: '节标题',
        title_only: '仅标题',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片布局设置成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${layoutName[layout] || layout}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片布局失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片布局出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 8. wps_ppt_get_slide_notes - 获取幻灯片备注内容
// ============================================================

export const getSlideNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_notes',
  description: `获取指定幻灯片的备注内容。

使用场景：
- "查看第3页的备注"
- "读取演讲备注"
- "这页有什么备注"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      notes: string;
    }>(
      'getSlideNotes',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const notes = response.data.notes;
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: notes
              ? `第 ${slideIndex} 页备注内容：\n${notes}`
              : `第 ${slideIndex} 页没有备注内容`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 9. wps_ppt_set_slide_notes - 设置幻灯片备注
// ============================================================

export const setSlideNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_notes',
  description: `设置幻灯片的备注内容，用于演讲提示。

使用场景：
- "给第1页添加备注"
- "写一些演讲提示"
- "修改备注内容"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      notes: {
        type: 'string',
        description: '备注内容',
      },
    },
    required: ['slideIndex', 'notes'],
  },
};

export const setSlideNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, notes } = args as {
    slideIndex: number;
    notes: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideNotes',
      { slideIndex, notes },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片备注设置成功！\n幻灯片: 第 ${slideIndex} 页\n备注内容: "${notes.substring(0, 80)}${notes.length > 80 ? '...' : ''}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 导出所有幻灯片操作相关的Tools
// ============================================================

export const slideOpsTools: RegisteredTool[] = [
  { definition: deleteSlideDefinition, handler: deleteSlideHandler },
  { definition: duplicateSlideDefinition, handler: duplicateSlideHandler },
  { definition: moveSlideDefinition, handler: moveSlideHandler },
  { definition: getSlideCountDefinition, handler: getSlideCountHandler },
  { definition: getSlideInfoDefinition, handler: getSlideInfoHandler },
  { definition: switchSlideDefinition, handler: switchSlideHandler },
  { definition: setSlideLayoutDefinition, handler: setSlideLayoutHandler },
  { definition: getSlideNotesDefinition, handler: getSlideNotesHandler },
  { definition: setSlideNotesDefinition, handler: setSlideNotesHandler },
];

export default slideOpsTools;
