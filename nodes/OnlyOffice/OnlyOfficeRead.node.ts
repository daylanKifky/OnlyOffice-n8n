import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  NodeOperationError,
  NodeConnectionType,
  IBinaryData,
} from 'n8n-workflow';
import { parseApiKeyPermissions } from './OnlyOfficeWebhook.types';
import { validatePermissions, getApiKeyPermissions, getPermissionErrorMessage } from './OnlyOfficePermissions.helper';

export class OnlyOfficeRead implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'OnlyOffice Read',
    name: 'onlyOfficeRead',
    icon: 'file:onlyoffice_doc.svg',
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
    description: 'Read operations for OnlyOffice files and folders',
    defaults: {
      name: 'OnlyOffice Read',
    },
    inputs: [NodeConnectionType.Main],
    outputs: [NodeConnectionType.Main],
    credentials: [
      {
        name: 'onlyOfficeApi',
        required: true,
      },
    ],
    requestDefaults: {
      baseURL: '={{$credentials.baseUrl}}/api/2.0',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    },
    properties: [
      {
        displayName: 'Resource',
        name: 'resource',
        type: 'options',
        noDataExpression: true,
        options: [
          {
            name: 'Folder',
            value: 'folder',
          },
          {
            name: 'File',
            value: 'file',
          },
          {
            name: 'API Key',
            value: 'apiKey',
          },
        ],
        default: 'folder',
      },
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['folder', 'file'],
          },
        },
        options: [
          {
            name: 'List',
            value: 'list',
            action: 'List items',
          },
          {
            name: 'Get',
            value: 'get',
            action: 'Get file contents',
            displayOptions: {
              show: {
                resource: ['file'],
              },
            },
          },
        ],
        default: 'list',
      },
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['apiKey'],
          },
        },
        options: [
          {
            name: 'Get Permissions',
            value: 'getPermissions',
            action: 'Get API key permissions',
          },
        ],
        default: 'getPermissions',
      },
      {
        displayName: 'Folder ID',
        name: 'folderId',
        type: 'string',
        default: '@my',
        required: true,
        displayOptions: {
          show: {
            operation: ['list'],
          },
        },
        description: 'ID of the folder to list contents from. Use @my for My Documents, @common for Common Documents.',
      },
      {
        displayName: 'Item ID',
        name: 'itemId',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['get'],
          },
        },
        description: 'ID of the file to download',
      },
      {
        displayName: 'Output Format',
        name: 'outputFormat',
        type: 'options',
        options: [
          {
            name: 'Original Format',
            value: 'original',
          },
          {
            name: 'PDF',
            value: 'pdf',
          },
        ],
        default: 'original',
        displayOptions: {
          show: {
            resource: ['file'],
            operation: ['get'],
          },
        },
        description: 'Format to download the file in',
      },
      {
        displayName: 'Binary Property',
        name: 'binaryPropertyName',
        type: 'string',
        default: 'data',
        required: true,
        displayOptions: {
          show: {
            resource: ['file'],
            operation: ['get'],
          },
        },
        description: 'Name of the binary property to store the file data',
      },
    ],
  };

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const returnData: INodeExecutionData[] = [];

    // Validate permissions before executing (skip for API key operations)
    const resource = this.getNodeParameter('resource', 0) as string;
    if (resource !== 'apiKey') {
      try {
        await validatePermissions(
          this,
          resource === 'file' ? 'files' : 'folders',
          'read'
        );
      } catch (error) {
        const permissions = await getApiKeyPermissions(this).catch(() => ({}));
        const errorMessage = getPermissionErrorMessage(
          resource === 'file' ? 'files' : 'folders',
          'read',
          permissions
        );
        throw new NodeOperationError(this.getNode(), errorMessage);
      }
    }

    for (let i = 0; i < items.length; i++) {
      try {
        const resource = this.getNodeParameter('resource', i) as string;
        const operation = this.getNodeParameter('operation', i) as string;

        let responseData;
        let binaryData: { [key: string]: IBinaryData } | undefined;

        if (resource === 'folder') {
          responseData = await OnlyOfficeRead.executeFolderOperation(this, operation, i);
        } else if (resource === 'file') {
          const result = await OnlyOfficeRead.executeFileOperation(this, operation, i);
          if (operation === 'get' && result.binary) {
            responseData = result.data;
            binaryData = result.binary;
          } else {
            responseData = result;
          }
        } else if (resource === 'apiKey') {
          responseData = await OnlyOfficeRead.executeApiKeyOperation(this, operation, i);
        }

        const item: INodeExecutionData = { json: responseData };
        if (binaryData) {
          item.binary = binaryData;
        }

        if (Array.isArray(responseData)) {
          returnData.push(...responseData.map(item => ({ json: item })));
        } else {
          returnData.push(item);
        }
      } catch (error) {
        if (this.continueOnFail()) {
          returnData.push({
            json: {
              error: error instanceof Error ? error.message : String(error),
            },
          });
          continue;
        }
        throw error;
      }
    }

    return [returnData];
  }

  private static parseResponse(response: any): any {
    if (Array.isArray(response) && response.length > 0 && typeof response[0] === 'string') {
      try {
        response = JSON.parse(response[0]);
      } catch (e) {
        return { error: 'Failed to parse JSON response', rawResponse: response };
      }
    }
    
    if (typeof response === 'string') {
      try {
        response = JSON.parse(response);
      } catch (e) {
        return { error: 'Failed to parse JSON response', rawResponse: response };
      }
    }
    
    if (response && response.response) {
      return response.response;
    }
    
    return response;
  }

  private static async executeFolderOperation(context: IExecuteFunctions, operation: string, itemIndex: number): Promise<any> {
    const credentials = await context.getCredentials('onlyOfficeApi');
    const baseUrl = `${credentials.baseUrl}/api/2.0`;

    switch (operation) {
      case 'list':
        const folderId = context.getNodeParameter('folderId', itemIndex) as string;
        const response = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/files/${folderId}`,
          },
        );
        
        const parsedResponse = OnlyOfficeRead.parseResponse(response);
        const folders = parsedResponse.folders || [];
        const files = parsedResponse.files || [];
        return [...folders, ...files];

      default:
        throw new NodeOperationError(context.getNode(), `Unknown folder operation: ${operation}`);
    }
  }

  private static async executeFileOperation(context: IExecuteFunctions, operation: string, itemIndex: number): Promise<any> {
    const credentials = await context.getCredentials('onlyOfficeApi');
    const baseUrl = `${credentials.baseUrl}/api/2.0`;

    switch (operation) {
      case 'list':
        const folderId = context.getNodeParameter('folderId', itemIndex) as string;
        const response = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/files/${folderId}`,
          },
        );
        
        const parsedResponse = OnlyOfficeRead.parseResponse(response);
        const folders = parsedResponse.folders || [];
        const files = parsedResponse.files || [];
        return [...folders, ...files];

      case 'get':
        const fileId = context.getNodeParameter('itemId', itemIndex) as string;
        const outputFormat = context.getNodeParameter('outputFormat', itemIndex) as string;
        const binaryPropertyName = context.getNodeParameter('binaryPropertyName', itemIndex) as string;
        
        context.logger.debug('OnlyOffice Get - Starting file download', {
          fileId,
          outputFormat,
          binaryPropertyName,
          baseUrl,
        });
        
        const fileInfo = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/files/file/${fileId}`,
          },
        );
        
        context.logger.debug('OnlyOffice Get - File metadata response received', {
          rawResponse: typeof fileInfo === 'string' ? fileInfo.substring(0, 200) : fileInfo,
        });
        
        const fileData = OnlyOfficeRead.parseResponse(fileInfo);
        const fileName = fileData.title || `file_${fileId}`;
        const fileExtension = outputFormat === 'pdf' ? 'pdf' : fileName.split('.').pop() || 'bin';
        
        context.logger.debug('OnlyOffice Get - File metadata parsed', {
          fileName,
          fileExtension,
          fileData: {
            id: fileData.id,
            title: fileData.title,
            fileType: fileData.fileType,
          },
        });
        
        const presignedUrl = `${baseUrl}/files/file/${fileId}/presigned`;
        context.logger.debug('OnlyOffice Get - Requesting presigned URI', {
          url: presignedUrl,
        });
        
        const presignedResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: presignedUrl,
          },
        );
        
        let safeRawResponse: string;
        try {
          safeRawResponse = typeof presignedResponse === 'string' 
            ? presignedResponse.substring(0, 200) 
            : JSON.stringify(presignedResponse, null, 2).substring(0, 200);
        } catch (e) {
          safeRawResponse = '[Unable to serialize response]';
        }
        
        context.logger.debug('OnlyOffice Get - Presigned URI response received', {
          rawResponse: safeRawResponse,
        });
        
        const presignedData = OnlyOfficeRead.parseResponse(presignedResponse);
        
        let safePresignedData: any;
        try {
          if (typeof presignedData === 'object' && presignedData !== null) {
            safePresignedData = {
              filetype: presignedData.filetype,
              token: presignedData.token ? presignedData.token.substring(0, 50) + '...' : undefined,
              url: presignedData.url ? presignedData.url.substring(0, 100) + '...' : undefined,
              uri: presignedData.uri ? presignedData.uri.substring(0, 100) + '...' : undefined,
            };
          } else {
            safePresignedData = presignedData;
          }
        } catch (e) {
          safePresignedData = '[Unable to serialize]';
        }
        
        context.logger.debug('OnlyOffice Get - Presigned data parsed', {
          presignedData: safePresignedData,
          presignedDataType: typeof presignedData,
          isString: typeof presignedData === 'string',
          isObject: typeof presignedData === 'object',
          keys: typeof presignedData === 'object' && presignedData !== null ? Object.keys(presignedData) : null,
        });
        
        let downloadUrl: string;
        if (typeof presignedData === 'string') {
          downloadUrl = presignedData;
        } else if (presignedData && typeof presignedData === 'object') {
          if (presignedData.uri && typeof presignedData.uri === 'string') {
            downloadUrl = presignedData.uri;
          } else if (presignedData.url && typeof presignedData.url === 'string') {
            downloadUrl = presignedData.url;
          } else if (presignedData.downloadUrl && typeof presignedData.downloadUrl === 'string') {
            downloadUrl = presignedData.downloadUrl;
          } else if (presignedData.token && typeof presignedData.token === 'string') {
            const baseUrlWithoutApi = baseUrl.replace('/api/2.0', '');
            downloadUrl = `${baseUrlWithoutApi}/filehandler.ashx?action=stream&fileid=${fileId}&token=${presignedData.token}`;
          } else {
            throw new NodeOperationError(
              context.getNode(),
              `Could not extract download URL from presigned response. Available properties: ${Object.keys(presignedData).join(', ')}. Response: ${JSON.stringify(presignedData)}`
            );
          }
        } else {
          throw new NodeOperationError(
            context.getNode(),
            `Invalid presigned URI response format. Expected string or object with uri/url/token, got: ${typeof presignedData}`
          );
        }
        
        if (!downloadUrl || typeof downloadUrl !== 'string') {
          throw new NodeOperationError(
            context.getNode(),
            `Could not extract download URL from presigned response. Response: ${JSON.stringify(presignedData)}`
          );
        }
        
        context.logger.debug('OnlyOffice Get - Presigned URI extracted', {
          downloadUrlLength: downloadUrl.length,
          downloadUrlPreview: downloadUrl.substring(0, 100) + '...',
          hasToken: downloadUrl.includes('token'),
        });
        
        context.logger.debug('OnlyOffice Get - Downloading file from presigned URI', {
          downloadUrl: downloadUrl.substring(0, 100) + '...',
          hasStreamAuth: downloadUrl.includes('stream_auth'),
        });
        
        try {
          const downloadResponse = await context.helpers.httpRequest({
            method: 'GET',
            url: downloadUrl,
            headers: {
              'Authorization': `Bearer ${credentials.token}`,
              'Accept': '*/*',
            },
            encoding: 'arraybuffer',
          });
          
          context.logger.debug('OnlyOffice Get - File download completed', {
            responseType: typeof downloadResponse,
            isArrayBuffer: downloadResponse instanceof ArrayBuffer,
          });
          
          const buffer = Buffer.from(downloadResponse as ArrayBuffer);
          const base64Data = buffer.toString('base64');
          const fileSizeBytes = buffer.length;
          const fileSizeKB = (fileSizeBytes / 1024).toFixed(2);
          
          context.logger.info('OnlyOffice Get - File successfully downloaded and converted', {
            fileId,
            fileName,
            fileExtension,
            fileSizeBytes,
            fileSizeKB: `${fileSizeKB} KB`,
            binaryPropertyName,
          });
          
          return {
            data: {
              id: fileId,
              title: fileName,
              format: outputFormat === 'pdf' ? 'pdf' : fileExtension,
            },
            binary: {
              [binaryPropertyName]: {
                data: base64Data,
                mimeType: OnlyOfficeRead.getMimeType(fileExtension),
                fileName: outputFormat === 'pdf' ? fileName.replace(/\.[^.]+$/, '.pdf') : fileName,
              },
            },
          };
        } catch (downloadError: any) {
          context.logger.error('OnlyOffice Get - File download failed', {
            error: downloadError.message,
            statusCode: downloadError.statusCode,
            downloadUrl: downloadUrl.substring(0, 150) + '...',
          });
          
          if (downloadError.statusCode === 403) {
            context.logger.debug('OnlyOffice Get - Retrying download without Authorization header');
            try {
              const retryResponse = await context.helpers.httpRequest({
                method: 'GET',
                url: downloadUrl,
                encoding: 'arraybuffer',
              });
              
              const buffer = Buffer.from(retryResponse as ArrayBuffer);
              const base64Data = buffer.toString('base64');
              const fileSizeBytes = buffer.length;
              const fileSizeKB = (fileSizeBytes / 1024).toFixed(2);
              
              context.logger.info('OnlyOffice Get - File successfully downloaded (without auth)', {
                fileId,
                fileName,
                fileSizeKB: `${fileSizeKB} KB`,
              });
              
              return {
                data: {
                  id: fileId,
                  title: fileName,
                  format: outputFormat === 'pdf' ? 'pdf' : fileExtension,
                },
                binary: {
                  [binaryPropertyName]: {
                    data: base64Data,
                    mimeType: OnlyOfficeRead.getMimeType(fileExtension),
                    fileName: outputFormat === 'pdf' ? fileName.replace(/\.[^.]+$/, '.pdf') : fileName,
                  },
                },
              };
            } catch (retryError: any) {
              throw new NodeOperationError(
                context.getNode(),
                `Failed to download file. Status: ${retryError.statusCode || 'unknown'}. Error: ${retryError.message}`
              );
            }
          }
          
          throw new NodeOperationError(
            context.getNode(),
            `Failed to download file. Status: ${downloadError.statusCode || 'unknown'}. Error: ${downloadError.message}`
          );
        }

      default:
        throw new NodeOperationError(context.getNode(), `Unknown file operation: ${operation}`);
    }
  }

  private static getMimeType(extension: string): string {
    const mimeTypes: { [key: string]: string } = {
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'pdf': 'application/pdf',
      'doc': 'application/msword',
      'xls': 'application/vnd.ms-excel',
      'ppt': 'application/vnd.ms-powerpoint',
    };
    return mimeTypes[extension.toLowerCase()] || 'application/octet-stream';
  }

  private static async executeApiKeyOperation(context: IExecuteFunctions, operation: string, itemIndex: number): Promise<any> {
    const credentials = await context.getCredentials('onlyOfficeApi');
    const baseUrl = `${credentials.baseUrl}/api/2.0`;

    switch (operation) {
      case 'getPermissions':
        const response = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/keys/permissions`,
          },
        );
        
        let permissions: string[] = [];
        
        if (response && typeof response === 'object' && response.response) {
          permissions = Array.isArray(response.response) ? response.response : [];
        } else if (Array.isArray(response)) {
          if (response.length > 0 && typeof response[0] === 'string' && response[0].startsWith('{')) {
            try {
              const parsed = JSON.parse(response[0]);
              permissions = parsed.response || parsed.permissions || [];
            } catch (e) {
              permissions = response;
            }
          } else {
            permissions = response;
          }
        } else if (typeof response === 'string') {
          try {
            const parsed = JSON.parse(response);
            permissions = Array.isArray(parsed) ? parsed : (parsed.response || parsed.permissions || []);
          } catch (e) {
            permissions = [];
          }
        }
        
        const parsedPermissions = parseApiKeyPermissions(permissions);
        return parsedPermissions;

      default:
        throw new NodeOperationError(context.getNode(), `Unknown API key operation: ${operation}`);
    }
  }
}

