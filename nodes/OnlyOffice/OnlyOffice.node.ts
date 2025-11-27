import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  NodeOperationError,
  NodeConnectionType,
} from 'n8n-workflow';

export class OnlyOffice implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'OnlyOffice',
    name: 'onlyOffice',
    icon: 'file:onlyoffice.svg',
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
            operation: ['rename', 'move', 'copy', 'delete'],
          },
        },
        description: 'ID of the folder or file to operate on',
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

        if (resource === 'folder') {
          responseData = await OnlyOffice.executeFolderOperation(this, operation, i);
        } else if (resource === 'file') {
          responseData = await OnlyOffice.executeFileOperation(this, operation, i);
        }

        if (Array.isArray(responseData)) {
          returnData.push(...responseData.map(item => ({ json: item })));
        } else {
          returnData.push({ json: responseData });
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
}