import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  NodeOperationError,
  NodeConnectionType,
  IBinaryData,
} from 'n8n-workflow';

export class OnlyOffice implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'OnlyOffice',
    name: 'onlyOffice',
    icon: 'file:onlyoffice_doc.svg',
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
    description: 'Interact with OnlyOffice files and folders',
    defaults: {
      name: 'OnlyOffice',
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
        ],
        default: 'folder',
      },

      // Folder Operations
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['folder'],
          },
        },
        options: [
          {
            name: 'List',
            value: 'list',
            action: 'List folders',
          },
          {
            name: 'Create',
            value: 'create',
            action: 'Create folder',
          },
          {
            name: 'Rename',
            value: 'rename',
            action: 'Rename folder',
          },
          {
            name: 'Move',
            value: 'move',
            action: 'Move folder',
          },
          {
            name: 'Copy',
            value: 'copy',
            action: 'Copy folder',
          },
          {
            name: 'Delete',
            value: 'delete',
            action: 'Delete folder',
          },
        ],
        default: 'list',
      },

      // File Operations
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['file'],
          },
        },
        options: [
          {
            name: 'List',
            value: 'list',
            action: 'List files',
          },
          {
            name: 'Get',
            value: 'get',
            action: 'Get file contents',
          },
          {
            name: 'Create',
            value: 'create',
            action: 'Create file',
          },
          {
            name: 'Rename',
            value: 'rename',
            action: 'Rename file',
          },
          {
            name: 'Move',
            value: 'move',
            action: 'Move file',
          },
          {
            name: 'Copy',
            value: 'copy',
            action: 'Copy file',
          },
          {
            name: 'Delete',
            value: 'delete',
            action: 'Delete file',
          },
        ],
        default: 'list',
      },

      // List Operations - Folder ID
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

      // Create Operations
      {
        displayName: 'Parent Folder ID',
        name: 'parentFolderId',
        type: 'string',
        default: '@my',
        required: true,
        displayOptions: {
          show: {
            operation: ['create'],
          },
        },
        description: 'ID of the parent folder where the new item will be created',
      },
      {
        displayName: 'Title',
        name: 'title',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['create'],
          },
        },
        description: 'Name of the new folder or file',
      },

      // File Type for Create File
      {
        displayName: 'File Type',
        name: 'fileType',
        type: 'options',
        options: [
          {
            name: 'Document (.docx)',
            value: 'docx',
          },
          {
            name: 'Spreadsheet (.xlsx)',
            value: 'xlsx',
          },
          {
            name: 'Presentation (.pptx)',
            value: 'pptx',
          },
        ],
        default: 'docx',
        required: true,
        displayOptions: {
          show: {
            resource: ['file'],
            operation: ['create'],
          },
        },
        description: 'Type of file to create',
      },

      // Operations requiring Item ID
      {
        displayName: 'Item ID',
        name: 'itemId',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['get', 'rename', 'move', 'copy', 'delete'],
          },
        },
        description: 'ID of the folder or file to operate on',
      },

      // Get Operation Options
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

      // Rename Operation
      {
        displayName: 'New Title',
        name: 'newTitle',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['rename'],
          },
        },
        description: 'New name for the folder or file',
      },

      // Move/Copy Operations
      {
        displayName: 'Destination Folder ID',
        name: 'destFolderId',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['move', 'copy'],
          },
        },
        description: 'ID of the destination folder',
      },
      {
        displayName: 'Conflict Resolution',
        name: 'conflictResolveType',
        type: 'options',
        options: [
          {
            name: 'Skip',
            value: 'Skip',
          },
          {
            name: 'Overwrite',
            value: 'Overwrite',
          },
          {
            name: 'Duplicate',
            value: 'Duplicate',
          },
        ],
        default: 'Skip',
        displayOptions: {
          show: {
            operation: ['move', 'copy'],
          },
        },
        description: 'How to handle conflicts when moving or copying',
      },

      // Delete Options
      {
        displayName: 'Delete Immediately',
        name: 'deleteImmediately',
        type: 'boolean',
        default: false,
        displayOptions: {
          show: {
            operation: ['delete'],
          },
        },
        description: 'Whether to delete immediately or move to trash',
      },
    ],
  };

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const returnData: INodeExecutionData[] = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const resource = this.getNodeParameter('resource', i) as string;
        const operation = this.getNodeParameter('operation', i) as string;

        let responseData;
        let binaryData: { [key: string]: IBinaryData } | undefined;

        if (resource === 'folder') {
          responseData = await OnlyOffice.executeFolderOperation(this, operation, i);
        } else if (resource === 'file') {
          const result = await OnlyOffice.executeFileOperation(this, operation, i);
          if (operation === 'get' && result.binary) {
            responseData = result.data;
            binaryData = result.binary;
          } else {
            responseData = result;
          }
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
    // Handle array with JSON string as first element: ["{ ... }"]
    if (Array.isArray(response) && response.length > 0 && typeof response[0] === 'string') {
      try {
        response = JSON.parse(response[0]);
      } catch (e) {
        return { error: 'Failed to parse JSON response', rawResponse: response };
      }
    }
    
    // Handle string response
    if (typeof response === 'string') {
      try {
        response = JSON.parse(response);
      } catch (e) {
        return { error: 'Failed to parse JSON response', rawResponse: response };
      }
    }
    
    // Extract data from OnlyOffice API response structure
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
        
        const parsedResponse = OnlyOffice.parseResponse(response);
        const folders = parsedResponse.folders || [];
        const files = parsedResponse.files || [];
        return [...folders, ...files];

      case 'create':
        const parentFolderId = context.getNodeParameter('parentFolderId', itemIndex) as string;
        const title = context.getNodeParameter('title', itemIndex) as string;
        const createResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'POST',
            url: `${baseUrl}/files/folder/${parentFolderId}`,
            body: { title },
          },
        );
        return OnlyOffice.parseResponse(createResponse);

      case 'rename':
        const itemId = context.getNodeParameter('itemId', itemIndex) as string;
        const newTitle = context.getNodeParameter('newTitle', itemIndex) as string;
        const renameResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/folder/${itemId}`,
            body: { title: newTitle },
          },
        );
        return OnlyOffice.parseResponse(renameResponse);

      case 'move':
      case 'copy':
        const moveItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const destFolderId = context.getNodeParameter('destFolderId', itemIndex) as string;
        const conflictResolveType = context.getNodeParameter('conflictResolveType', itemIndex) as string;
        const moveResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/fileops/${operation}`,
            body: {
              folderIds: [moveItemId],
              fileIds: [],
              destFolderId,
              conflictResolveType,
            },
          },
        );
        return OnlyOffice.parseResponse(moveResponse);

      case 'delete':
        const deleteItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const deleteImmediately = context.getNodeParameter('deleteImmediately', itemIndex) as boolean;
        const deleteResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'DELETE',
            url: `${baseUrl}/files/folder/${deleteItemId}`,
            body: deleteImmediately ? { deleteAfter: true } : {},
          },
        );
        return OnlyOffice.parseResponse(deleteResponse);

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
        
        const parsedResponse = OnlyOffice.parseResponse(response);
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
        
        // Get file metadata to determine file extension
        context.logger.debug('OnlyOffice Get - Fetching file metadata', {
          url: `${baseUrl}/files/file/${fileId}`,
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
        
        const fileData = OnlyOffice.parseResponse(fileInfo);
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
        
        // Get presigned URI for file download
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
        
        // Safely serialize response for logging (avoid circular references)
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
        
        const presignedData = OnlyOffice.parseResponse(presignedResponse);
        
        // Safely serialize presignedData for logging
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
        
        // Extract download URL from presigned response
        // The response might be a string URI, or an object with uri/url/token properties
        let downloadUrl: string;
        if (typeof presignedData === 'string') {
          downloadUrl = presignedData;
        } else if (presignedData && typeof presignedData === 'object') {
          // Try common property names for the URI first
          if (presignedData.uri && typeof presignedData.uri === 'string') {
            downloadUrl = presignedData.uri;
          } else if (presignedData.url && typeof presignedData.url === 'string') {
            downloadUrl = presignedData.url;
          } else if (presignedData.downloadUrl && typeof presignedData.downloadUrl === 'string') {
            downloadUrl = presignedData.downloadUrl;
          } else if (presignedData.token && typeof presignedData.token === 'string') {
            // Token needs to be used with filehandler endpoint
            // Construct URL using the base URL without /api/2.0
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
        
        // Download file using presigned URI
        // Some OnlyOffice instances require Authorization header even for presigned URLs
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
          
          // Convert arraybuffer to base64 for n8n binary data
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
                mimeType: OnlyOffice.getMimeType(fileExtension),
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
          
          // If 403, try without Authorization header (some presigned URLs don't need it)
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
                    mimeType: OnlyOffice.getMimeType(fileExtension),
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

      case 'create':
        const parentFolderId = context.getNodeParameter('parentFolderId', itemIndex) as string;
        const title = context.getNodeParameter('title', itemIndex) as string;
        const fileType = context.getNodeParameter('fileType', itemIndex) as string;
        const createResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'POST',
            url: `${baseUrl}/files/${parentFolderId}/file`,
            body: { 
              title: `${title}.${fileType}`,
              templateId: OnlyOffice.getTemplateId(fileType),
            },
          },
        );
        return OnlyOffice.parseResponse(createResponse);

      case 'rename':
        const itemId = context.getNodeParameter('itemId', itemIndex) as string;
        const newTitle = context.getNodeParameter('newTitle', itemIndex) as string;
        const renameResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/file/${itemId}`,
            body: { title: newTitle },
          },
        );
        return OnlyOffice.parseResponse(renameResponse);

      case 'move':
      case 'copy':
        const moveItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const destFolderId = context.getNodeParameter('destFolderId', itemIndex) as string;
        const conflictResolveType = context.getNodeParameter('conflictResolveType', itemIndex) as string;
        const moveResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/fileops/${operation}`,
            body: {
              folderIds: [],
              fileIds: [moveItemId],
              destFolderId,
              conflictResolveType,
            },
          },
        );
        return OnlyOffice.parseResponse(moveResponse);

      case 'delete':
        const deleteItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const deleteImmediately = context.getNodeParameter('deleteImmediately', itemIndex) as boolean;
        const deleteResponse = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'DELETE',
            url: `${baseUrl}/files/file/${deleteItemId}`,
            body: { deleteAfter: deleteImmediately },
          },
        );
        return OnlyOffice.parseResponse(deleteResponse);

      default:
        throw new NodeOperationError(context.getNode(), `Unknown file operation: ${operation}`);
    }
  }

  private static getTemplateId(fileType: string): number {
    // These are standard OnlyOffice template IDs
    switch (fileType) {
      case 'docx':
        return 1;
      case 'xlsx':
        return 2;
      case 'pptx':
        return 3;
      default:
        return 1;
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
}