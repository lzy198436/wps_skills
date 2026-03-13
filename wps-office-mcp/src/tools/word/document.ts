/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_get_document_text: 获取文档文本内容
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

/**
 * 获取所有已打开的文档列表
 */
export const getOpenDocumentsDefinition: ToolDefinition = {
  name: 'wps_word_get_open_documents',
  description: `获取当前WPS Writer中所有已打开的文档列表。

使用场景：
- "看看现在打开了哪些文档"
- "列出所有打开的Word文件"
- "查看当前文档列表"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getOpenDocumentsHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documents: Array<{ name: string; path: string; active: boolean }>;
    }>(
      'getOpenDocuments',
      {},
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const docs = response.data.documents;
      if (!docs || docs.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有打开任何文档' }],
        };
      }

      const docList = docs
        .map((doc, i) => `${i + 1}. ${doc.name}${doc.active ? ' (当前活动)' : ''}\n   路径: ${doc.path}`)
        .join('\n');

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已打开的文档列表 (共${docs.length}个):\n${docList}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定文档
 */
export const switchDocumentDefinition: ToolDefinition = {
  name: 'wps_word_switch_document',
  description: `切换到指定名称的文档。

使用场景：
- "切换到报告.docx"
- "打开另一个文档窗口"
- "切换到那个合同文档"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的文档名称',
      },
    },
    required: ['name'],
  },
};

export const switchDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  if (!name || name.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文档名称不能为空！' }],
      error: '文档名称为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchDocument',
      { name },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到文档: ${name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 打开指定路径的文档
 */
export const openDocumentDefinition: ToolDefinition = {
  name: 'wps_word_open_document',
  description: `打开指定路径的Word文档。

使用场景：
- "打开桌面上的报告.docx"
- "帮我打开这个文件路径的文档"
- "加载指定位置的Word文件"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '要打开的文档文件路径',
      },
    },
    required: ['filePath'],
  },
};

export const openDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath } = args as { filePath: string };

  if (!filePath || filePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文件路径不能为空！' }],
      error: '文件路径为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documentName: string;
    }>(
      'openDocument',
      { filePath },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档打开成功！\n文件: ${response.data.documentName || filePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `打开文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `打开文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取文档文本内容
 */
export const getDocumentTextDefinition: ToolDefinition = {
  name: 'wps_word_get_document_text',
  description: `获取当前Word文档的文本内容。

使用场景：
- "读取文档内容"
- "获取文档的前100个字符"
- "查看文档从第50到第200个字符的内容"

可指定起始和结束位置来获取部分文本。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      start: {
        type: 'number',
        description: '起始位置（字符索引），默认从头开始',
      },
      end: {
        type: 'number',
        description: '结束位置（字符索引），默认到文档末尾',
      },
    },
  },
};

export const getDocumentTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { start, end } = args as { start?: number; end?: number };

  try {
    const params: Record<string, unknown> = {};
    if (start !== undefined) params.start = start;
    if (end !== undefined) params.end = end;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      text: string;
      length: number;
    }>(
      'getDocumentText',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const rangeInfo = start !== undefined || end !== undefined
        ? `\n范围: ${start ?? 0} - ${end ?? '末尾'}`
        : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档文本内容 (${response.data.length}字符)${rangeInfo}:\n\n${response.data.text}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有文档管理相关的Tools
 */
export const documentTools: RegisteredTool[] = [
  { definition: getOpenDocumentsDefinition, handler: getOpenDocumentsHandler },
  { definition: switchDocumentDefinition, handler: switchDocumentHandler },
  { definition: openDocumentDefinition, handler: openDocumentHandler },
  { definition: getDocumentTextDefinition, handler: getDocumentTextHandler },
];

export default documentTools;
