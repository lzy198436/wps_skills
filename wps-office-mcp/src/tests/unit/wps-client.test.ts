/**
 * Input: Mock 的 WPS 调用（macPollServer / child_process）
 * Output: 客户端行为断言
 * Pos: WPS 客户端单元测试。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS客户端单元测试 - 基于macPollServer/PowerShell架构重写
 *
 * 测试 WpsClient 在 macOS（macPollServer轮询）架构下的行为
 * 不再检查 axios 调用，改为验证方法返回值与行为
 *
 * @author 测试重写agent
 * @date 2026-03-21
 */

// Mock mac-poll-server - macOS上WpsClient走轮询服务器
const mockExecuteCommand = jest.fn().mockImplementation(async (action: string, params: Record<string, unknown>) => {
  // 模拟WPS返回结果，根据action返回对应格式
  if (action === 'ping') return { success: true, data: { status: 'ok' } };
  if (action === 'getActiveDocument') return { success: true, data: { name: 'test.docx', path: '/tmp/test.docx', saved: true, readOnly: false } };
  if (action === 'getActiveWorkbook') return { success: true, data: { name: 'test.xlsx', sheets: ['Sheet1'] } };
  if (action === 'getActivePresentation') return { success: true, data: { name: 'test.pptx', slideCount: 10, currentSlideIndex: 1 } };
  if (action === 'insertText') return { success: true };
  if (action === 'getCellValue') return { success: true, data: { value: 'test_value' } };
  if (action === 'setCellValue') return { success: true };
  if (action === 'getRangeData') return { success: true, data: { data: [['A1', 'B1'], ['A2', 'B2']] } };
  if (action === 'setRangeData') return { success: true };
  if (action === 'setFormula') return { success: true };
  if (action === 'addSlide') return { success: true };
  if (action === 'openFile') return { success: true };
  if (action === 'save') return { success: true };
  if (action === 'saveAs') return { success: true };
  if (action === 'createDocument') return { success: true };
  if (action === 'createPresentation') return { success: true };
  if (action === 'getDocumentText') return { success: true, data: { text: 'Hello World' } };
  return { success: true, data: params };
});

jest.mock('../../client/mac-poll-server', () => ({
  macPollServer: {
    isRunning: true,
    start: jest.fn().mockResolvedValue(undefined),
    executeCommand: mockExecuteCommand,
  },
}));

// Mock child_process.spawn - Windows上走PowerShell，测试环境不需要
jest.mock('child_process', () => ({
  spawn: jest.fn(),
}));

// Mock axios - 虽然当前架构macOS不走axios，但构造函数中可能仍有import
jest.mock('axios', () => ({
  create: jest.fn(() => ({
    post: jest.fn(),
    get: jest.fn(),
    interceptors: {
      request: { use: jest.fn() },
      response: { use: jest.fn() },
    },
    defaults: {},
  })),
  isAxiosError: jest.fn(),
}));

import { WpsClient } from '../../client/wps-client';
import { WpsAppType, WpsApiRequest } from '../../types/wps';

// Mock logger - 不想看到一堆日志输出
jest.mock('../../utils/logger', () => ({
  log: {
    info: jest.fn(),
    debug: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
  logRequest: jest.fn(),
  logResponse: jest.fn(),
}));

// Mock错误工具
jest.mock('../../utils/error', () => ({
  WpsConnectionError: class WpsConnectionError extends Error {
    constructor(message: string, public details?: unknown) {
      super(message);
      this.name = 'WpsConnectionError';
    }
  },
  WpsApiError: class WpsApiError extends Error {
    constructor(message: string, public details?: unknown) {
      super(message);
      this.name = 'WpsApiError';
    }
  },
  TimeoutError: class TimeoutError extends Error {
    constructor(operation: string, timeout: number) {
      super(`Operation '${operation}' timed out after ${timeout}ms`);
      this.name = 'TimeoutError';
    }
  },
  errorUtils: {
    wrap: jest.fn((error: unknown, message: string) => {
      const err = error instanceof Error ? error : new Error(String(error));
      err.message = `${message}: ${err.message}`;
      return err;
    }),
  },
}));

describe('WpsClient', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    // 重置默认mock实现
    mockExecuteCommand.mockImplementation(async (action: string, params: Record<string, unknown>) => {
      if (action === 'ping') return { success: true, data: { status: 'ok' } };
      if (action === 'getActiveDocument') return { success: true, data: { name: 'test.docx', path: '/tmp/test.docx', saved: true, readOnly: false } };
      if (action === 'getActiveWorkbook') return { success: true, data: { name: 'test.xlsx', sheets: ['Sheet1'] } };
      if (action === 'getActivePresentation') return { success: true, data: { name: 'test.pptx', slideCount: 10, currentSlideIndex: 1 } };
      if (action === 'insertText') return { success: true };
      if (action === 'getCellValue') return { success: true, data: { value: 'test_value' } };
      if (action === 'setCellValue') return { success: true };
      if (action === 'getRangeData') return { success: true, data: { data: [['A1', 'B1'], ['A2', 'B2']] } };
      if (action === 'setRangeData') return { success: true };
      if (action === 'setFormula') return { success: true };
      if (action === 'addSlide') return { success: true };
      if (action === 'openFile') return { success: true };
      if (action === 'save') return { success: true };
      if (action === 'saveAs') return { success: true };
      if (action === 'createDocument') return { success: true };
      if (action === 'createPresentation') return { success: true };
      if (action === 'getDocumentText') return { success: true, data: { text: 'Hello World' } };
      return { success: true, data: params };
    });
  });

  // ==================== 构造函数 ====================
  describe('构造函数', () => {
    it('应该成功创建客户端（默认配置）', () => {
      const client = new WpsClient();
      expect(client).toBeDefined();
    });

    it('应该接受自定义配置', () => {
      const client = new WpsClient({ baseUrl: 'http://custom:8080', timeout: 5000 });
      expect(client).toBeDefined();
    });

    it('应该接受空配置', () => {
      const client = new WpsClient({});
      expect(client).toBeDefined();
    });

    it('应该接受仅baseUrl配置', () => {
      const client = new WpsClient({ baseUrl: 'http://localhost:9999' });
      expect(client).toBeDefined();
    });

    it('应该接受仅timeout配置', () => {
      const client = new WpsClient({ timeout: 60000 });
      expect(client).toBeDefined();
    });
  });

  // ==================== 客户端状态 ====================
  describe('getStatus - 获取客户端状态', () => {
    it('初始状态应该是未连接', () => {
      const client = new WpsClient();
      const status = client.getStatus();
      expect(status).toHaveProperty('connected');
      expect(status.connected).toBe(false);
    });

    it('返回的状态应该是副本，修改不影响原始状态', () => {
      const client = new WpsClient();
      const status1 = client.getStatus();
      status1.connected = true;
      const status2 = client.getStatus();
      expect(status2.connected).toBe(false);
    });

    it('成功调用后状态应更新为已连接', async () => {
      const client = new WpsClient();
      await client.checkConnection();
      const status = client.getStatus();
      expect(status.connected).toBe(true);
    });

    it('调用失败后状态应更新为未连接', async () => {
      const client = new WpsClient();
      // 先连接成功
      await client.checkConnection();
      expect(client.getStatus().connected).toBe(true);
      // 然后模拟失败
      mockExecuteCommand.mockRejectedValueOnce(new Error('Connection lost'));
      await client.checkConnection();
      expect(client.getStatus().connected).toBe(false);
    });
  });

  // ==================== 连接检查 ====================
  describe('checkConnection - 连接检查', () => {
    it('连接正常时应返回true', async () => {
      const client = new WpsClient();
      const result = await client.checkConnection();
      expect(result).toBe(true);
    });

    it('应该调用macPollServer.executeCommand发送ping', async () => {
      const client = new WpsClient();
      await client.checkConnection();
      expect(mockExecuteCommand).toHaveBeenCalledWith('ping', {});
    });

    it('连接失败时应返回false', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('Connection refused'));
      const client = new WpsClient();
      const result = await client.checkConnection();
      expect(result).toBe(false);
    });

    it('连接失败后状态应包含错误信息', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('Connection refused'));
      const client = new WpsClient();
      await client.checkConnection();
      const status = client.getStatus();
      expect(status.connected).toBe(false);
      expect(status.error).toBe('Connection check failed');
    });
  });

  // ==================== invokeAction ====================
  describe('invokeAction - 核心调用方法', () => {
    it('应该成功调用并返回响应', async () => {
      const client = new WpsClient();
      const result = await client.invokeAction('ping');
      expect(result).toEqual({ success: true, data: { status: 'ok' } });
    });

    it('应该将action和params传递给macPollServer', async () => {
      const client = new WpsClient();
      await client.invokeAction('setCellValue', { sheet: 'Sheet1', row: 1, col: 1, value: 'test' });
      expect(mockExecuteCommand).toHaveBeenCalledWith('setCellValue', { sheet: 'Sheet1', row: 1, col: 1, value: 'test' });
    });

    it('调用失败时应抛出包装后的错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('WPS not responding'));
      const client = new WpsClient();
      await expect(client.invokeAction('someAction')).rejects.toThrow();
    });

    it('成功调用后应更新lastHeartbeat', async () => {
      const client = new WpsClient();
      await client.invokeAction('ping');
      const status = client.getStatus();
      expect(status.lastHeartbeat).toBeDefined();
    });
  });

  // ==================== callApi - 兼容旧API ====================
  describe('callApi - 兼容旧API映射', () => {
    it('应该将旧方法名映射为新action名', async () => {
      const client = new WpsClient();
      const request: WpsApiRequest = { method: 'cell.getValue', params: { sheet: 'Sheet1', row: 1, col: 1 } };
      await client.callApi(request);
      expect(mockExecuteCommand).toHaveBeenCalledWith('getCellValue', { sheet: 'Sheet1', row: 1, col: 1 });
    });

    it('应该映射ping方法', async () => {
      const client = new WpsClient();
      await client.callApi({ method: 'ping' });
      expect(mockExecuteCommand).toHaveBeenCalledWith('ping', {});
    });

    it('应该映射workbook.getActive为getActiveWorkbook', async () => {
      const client = new WpsClient();
      await client.callApi({ method: 'workbook.getActive' });
      expect(mockExecuteCommand).toHaveBeenCalledWith('getActiveWorkbook', {});
    });

    it('应该映射document.insertText为insertText', async () => {
      const client = new WpsClient();
      await client.callApi({ method: 'document.insertText', params: { text: 'hi' } });
      expect(mockExecuteCommand).toHaveBeenCalledWith('insertText', { text: 'hi' });
    });

    it('应该映射file.open为openFile', async () => {
      const client = new WpsClient();
      await client.callApi({ method: 'file.open', params: { path: '/tmp/a.xlsx' } });
      expect(mockExecuteCommand).toHaveBeenCalledWith('openFile', { path: '/tmp/a.xlsx' });
    });

    it('未映射的方法应直接透传', async () => {
      const client = new WpsClient();
      await client.callApi({ method: 'custom.method', params: { arg: 1 } });
      expect(mockExecuteCommand).toHaveBeenCalledWith('custom.method', { arg: 1 });
    });

    it('请求失败时应抛出错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('fail'));
      const client = new WpsClient();
      await expect(client.callApi({ method: 'ping' })).rejects.toThrow();
    });
  });

  // ==================== Excel操作 ====================
  describe('表格操作 (WPS表格)', () => {
    it('getCellValue 应返回单元格值', async () => {
      const client = new WpsClient();
      const result = await client.getCellValue('Sheet1', 1, 1);
      expect(result).toBe('test_value');
    });

    it('getCellValue 应传递正确参数', async () => {
      const client = new WpsClient();
      await client.getCellValue('Sheet1', 2, 3);
      expect(mockExecuteCommand).toHaveBeenCalledWith('getCellValue', { sheet: 'Sheet1', row: 2, col: 3 });
    });

    it('getCellValue 支持数字索引的sheet', async () => {
      const client = new WpsClient();
      await client.getCellValue(0, 1, 1);
      expect(mockExecuteCommand).toHaveBeenCalledWith('getCellValue', { sheet: 0, row: 1, col: 1 });
    });

    it('setCellValue 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.setCellValue('Sheet1', 1, 1, 'Test Value');
      expect(result).toBe(true);
    });

    it('setCellValue 应传递正确参数', async () => {
      const client = new WpsClient();
      await client.setCellValue('Sheet1', 1, 1, 42);
      expect(mockExecuteCommand).toHaveBeenCalledWith('setCellValue', { sheet: 'Sheet1', row: 1, col: 1, value: 42 });
    });

    it('getRangeData 应返回二维数组', async () => {
      const client = new WpsClient();
      const result = await client.getRangeData('Sheet1', 'A1:B2');
      expect(result).toEqual([['A1', 'B1'], ['A2', 'B2']]);
    });

    it('getRangeData 应传递正确参数', async () => {
      const client = new WpsClient();
      await client.getRangeData('Sheet1', 'A1:C3');
      expect(mockExecuteCommand).toHaveBeenCalledWith('getRangeData', { sheet: 'Sheet1', range: 'A1:C3' });
    });

    it('getRangeData 数据为空时应返回空数组', async () => {
      mockExecuteCommand.mockResolvedValueOnce({ success: true, data: {} });
      const client = new WpsClient();
      const result = await client.getRangeData('Sheet1', 'A1:B2');
      expect(result).toEqual([]);
    });

    it('setRangeData 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.setRangeData('Sheet1', 'A1', [['v1', 'v2']]);
      expect(result).toBe(true);
    });

    it('setRangeData 应传递正确参数', async () => {
      const data = [['A', 'B'], ['C', 'D']];
      const client = new WpsClient();
      await client.setRangeData('Sheet1', 'A1:B2', data);
      expect(mockExecuteCommand).toHaveBeenCalledWith('setRangeData', { sheet: 'Sheet1', range: 'A1:B2', data });
    });

    it('getActiveWorkbook 应返回工作簿信息', async () => {
      const client = new WpsClient();
      const result = await client.getActiveWorkbook();
      expect(result).toEqual({ name: 'test.xlsx', sheets: ['Sheet1'] });
    });

    it('getActiveWorkbook 失败时应返回null', async () => {
      mockExecuteCommand.mockResolvedValueOnce({ success: false });
      const client = new WpsClient();
      const result = await client.getActiveWorkbook();
      expect(result).toBeNull();
    });

    it('setFormula 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.setFormula('Sheet1', 1, 1, '=SUM(A1:A10)');
      expect(result).toBe(true);
    });

    it('setFormula 应传递正确参数', async () => {
      const client = new WpsClient();
      await client.setFormula('Sheet1', 1, 1, '=SUM(A1:A10)');
      expect(mockExecuteCommand).toHaveBeenCalledWith('setFormula', { sheet: 'Sheet1', row: 1, col: 1, formula: '=SUM(A1:A10)' });
    });
  });

  // ==================== Word操作 ====================
  describe('文档操作 (WPS文字)', () => {
    it('getActiveDocument 应返回文档信息', async () => {
      const client = new WpsClient();
      const result = await client.getActiveDocument();
      expect(result).toEqual({ name: 'test.docx', path: '/tmp/test.docx', saved: true, readOnly: false });
    });

    it('getActiveDocument 失败时应返回null', async () => {
      mockExecuteCommand.mockResolvedValueOnce({ success: false });
      const client = new WpsClient();
      const result = await client.getActiveDocument();
      expect(result).toBeNull();
    });

    it('insertText 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.insertText('Hello World');
      expect(result).toBe(true);
    });

    it('insertText 应传递文本和位置参数', async () => {
      const client = new WpsClient();
      await client.insertText('Hello', 100);
      expect(mockExecuteCommand).toHaveBeenCalledWith('insertText', { text: 'Hello', position: 100 });
    });

    it('insertText 不带位置参数时position为undefined', async () => {
      const client = new WpsClient();
      await client.insertText('Hello');
      expect(mockExecuteCommand).toHaveBeenCalledWith('insertText', { text: 'Hello', position: undefined });
    });

    it('createDocument 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.createDocument();
      expect(result).toBe(true);
    });

    it('getDocumentText 应返回文本内容', async () => {
      const client = new WpsClient();
      const result = await client.getDocumentText();
      expect(result).toBe('Hello World');
    });

    it('getDocumentText 数据为空时应返回空字符串', async () => {
      mockExecuteCommand.mockResolvedValueOnce({ success: true, data: {} });
      const client = new WpsClient();
      const result = await client.getDocumentText();
      expect(result).toBe('');
    });
  });

  // ==================== PPT操作 ====================
  describe('演示操作 (WPS演示)', () => {
    it('getActivePresentation 应返回演示文稿信息', async () => {
      const client = new WpsClient();
      const result = await client.getActivePresentation();
      expect(result).toEqual({ name: 'test.pptx', slideCount: 10, currentSlideIndex: 1 });
    });

    it('getActivePresentation 失败时应返回null', async () => {
      mockExecuteCommand.mockResolvedValueOnce({ success: false });
      const client = new WpsClient();
      const result = await client.getActivePresentation();
      expect(result).toBeNull();
    });

    it('addSlide 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.addSlide();
      expect(result).toBe(true);
    });

    it('addSlide 应传递layout参数', async () => {
      const client = new WpsClient();
      await client.addSlide('title');
      expect(mockExecuteCommand).toHaveBeenCalledWith('addSlide', { layout: 'title' });
    });

    it('addSlide 不传layout时应传递undefined', async () => {
      const client = new WpsClient();
      await client.addSlide();
      expect(mockExecuteCommand).toHaveBeenCalledWith('addSlide', { layout: undefined });
    });

    it('createPresentation 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.createPresentation();
      expect(result).toBe(true);
    });
  });

  // ==================== 通用操作 ====================
  describe('通用操作', () => {
    it('executeMethod 应返回成功结果', async () => {
      const client = new WpsClient();
      const result = await client.executeMethod('custom.method', { arg: 1 });
      expect(result.success).toBe(true);
    });

    it('executeMethod 应传递方法名和参数', async () => {
      const client = new WpsClient();
      await client.executeMethod('custom.method', { arg: 1 });
      expect(mockExecuteCommand).toHaveBeenCalledWith('custom.method', { arg: 1 });
    });

    it('executeMethod 不传参数时应传递空对象（invokeAction默认值）', async () => {
      const client = new WpsClient();
      await client.executeMethod('custom.method');
      // executeMethod传undefined给invokeAction，invokeAction默认params={}
      expect(mockExecuteCommand).toHaveBeenCalledWith('custom.method', {});
    });

    it('openFile 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.openFile('/path/to/file.xlsx');
      expect(result).toBe(true);
    });

    it('openFile 应传递文件路径', async () => {
      const client = new WpsClient();
      await client.openFile('/path/to/file.xlsx', WpsAppType.SPREADSHEET);
      expect(mockExecuteCommand).toHaveBeenCalledWith('openFile', { path: '/path/to/file.xlsx' });
    });

    it('saveFile 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.saveFile();
      expect(result).toBe(true);
    });

    it('saveFile 应调用save action', async () => {
      const client = new WpsClient();
      await client.saveFile(WpsAppType.WRITER);
      expect(mockExecuteCommand).toHaveBeenCalledWith('save', {});
    });

    it('saveFileAs 应返回true', async () => {
      const client = new WpsClient();
      const result = await client.saveFileAs('/path/to/output.xlsx');
      expect(result).toBe(true);
    });

    it('saveFileAs 应传递文件路径', async () => {
      const client = new WpsClient();
      await client.saveFileAs('/path/to/output.xlsx', WpsAppType.SPREADSHEET);
      expect(mockExecuteCommand).toHaveBeenCalledWith('saveAs', { path: '/path/to/output.xlsx' });
    });
  });

  // ==================== 错误处理 ====================
  describe('错误处理', () => {
    it('macPollServer抛出错误时invokeAction应抛出包装后的错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('Unknown action'));
      const client = new WpsClient();
      await expect(client.invokeAction('nonexistent')).rejects.toThrow();
    });

    it('executeMethod调用失败时应抛出错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('WPS crashed'));
      const client = new WpsClient();
      await expect(client.executeMethod('bad.method')).rejects.toThrow();
    });

    it('getCellValue调用失败时应抛出错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('Sheet not found'));
      const client = new WpsClient();
      await expect(client.getCellValue('NonExist', 1, 1)).rejects.toThrow();
    });

    it('openFile调用失败时应抛出错误', async () => {
      mockExecuteCommand.mockRejectedValueOnce(new Error('File not found'));
      const client = new WpsClient();
      await expect(client.openFile('/nonexistent/file.xlsx')).rejects.toThrow();
    });

    it('连续失败后再次成功应恢复connected状态', async () => {
      const client = new WpsClient();
      // 失败
      mockExecuteCommand.mockRejectedValueOnce(new Error('fail'));
      await client.checkConnection();
      expect(client.getStatus().connected).toBe(false);
      // 恢复
      const result = await client.checkConnection();
      expect(result).toBe(true);
      expect(client.getStatus().connected).toBe(true);
    });
  });
});
