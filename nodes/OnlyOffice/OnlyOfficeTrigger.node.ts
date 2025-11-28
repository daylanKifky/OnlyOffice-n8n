import {
  IHookFunctions,
  IWebhookFunctions,
  IWebhookResponseData,
  INodeType,
  INodeTypeDescription,
  NodeConnectionType,
} from 'n8n-workflow';

export class OnlyOfficeTrigger implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'OnlyOffice Trigger',
    name: 'onlyOfficeTrigger',
    icon: 'file:onlyoffice.svg',
    group: ['trigger'],
    version: 1,
    description: 'Triggers when files or folders are created, updated, or modified in OnlyOffice',
    defaults: {
      name: 'OnlyOffice Trigger',
    },
    inputs: [],
    outputs: [NodeConnectionType.Main],
    credentials: [
      {
        name: 'onlyOfficeApi',
        required: true,
      },
    ],
    webhooks: [
      {
        name: 'default',
        httpMethod: 'POST',
        responseMode: 'onReceived',
        path: 'webhook',
      },
      {
        name: 'setup',
        httpMethod: 'HEAD',
        responseMode: 'onReceived',
        path: 'webhook',
      },
    ],
    properties: [
      {
        displayName: 'Events',
        name: 'events',
        type: 'multiOptions',
        options: [
          {
            name: 'File Created',
            value: 'FileCreated',
          },
          {
            name: 'File Updated',
            value: 'FileUpdated',
          },
          {
            name: 'File Renamed',
            value: 'FileRenamed',
          },
          {
            name: 'File Moved',
            value: 'FileMoved',
          },
          {
            name: 'File Deleted',
            value: 'FileDeleted',
          },
          {
            name: 'File Copied',
            value: 'FileCopied',
          },
          {
            name: 'Folder Created',
            value: 'FolderCreated',
          },
          {
            name: 'Folder Renamed',
            value: 'FolderRenamed',
          },
          {
            name: 'Folder Moved',
            value: 'FolderMoved',
          },
          {
            name: 'Folder Deleted',
            value: 'FolderDeleted',
          },
        ],
        default: ['FileCreated', 'FileUpdated'],
        required: true,
        description: 'The events to listen for',
      },
    ],
  };

  webhookMethods = {
    setup: {
      async checkExists(this: IHookFunctions): Promise<boolean> {
        const webhookData = this.getWorkflowStaticData('node');
        
        // If we have a webhook ID in static data, assume it exists
        if (webhookData.webhookId) {
          this.logger.debug('OnlyOffice Trigger - checkExists: Webhook ID found in static data', {
            webhookId: webhookData.webhookId,
          });
          return true;
        }
        
        // No webhook ID in static data, check if one exists on the server
        this.logger.debug('OnlyOffice Trigger - checkExists: No webhook ID in static data, checking server');
        
        try {
          const webhookUrl = this.getNodeWebhookUrl('default');
          const credentials = await this.getCredentials('onlyOfficeApi');
          const baseUrl = credentials.baseUrl as string;
          
          const response = await this.helpers.httpRequest({
            method: 'GET',
            url: `${baseUrl}/api/2.0/settings/webhooks`,
            headers: {
              'Authorization': `Bearer ${credentials.token}`,
              'Accept': 'application/json',
            },
          });
          
          const webhooks = response.response || [];
          const existing = webhooks.find((webhook: any) => webhook.uri === webhookUrl);
          
          if (existing) {
            // Found existing webhook, save its ID
            this.logger.debug('OnlyOffice Trigger - checkExists: Found existing webhook on server', {
              webhookId: existing.id,
            });
            webhookData.webhookId = existing.id;
            return true;
          }
          
          this.logger.debug('OnlyOffice Trigger - checkExists: No webhook found', {
            totalWebhooks: webhooks.length,
          });
          return false;
        } catch (error) {
          const err = error as any;
          this.logger.error('OnlyOffice Trigger - checkExists: Failed to check webhook existence', {
            error: err.message,
          });
          return false;
        }
      },
      async create(this: IHookFunctions): Promise<boolean> {
        try {
          const webhookUrl = this.getNodeWebhookUrl('default');
          const webhookData = this.getWorkflowStaticData('node');
          const credentials = await this.getCredentials('onlyOfficeApi');
          const baseUrl = credentials.baseUrl as string;
          const apiUrl = `${baseUrl}/api/2.0/settings/webhook`;
          
          const requestBody = {
            name: 'n8n Trigger Webhook',
            uri: webhookUrl,
            secretKey: 'n8nWebhook',
            enabled: true,
            ssl: false,
            triggers: 0, // 0 = All events
          };
          
          this.logger.info('OnlyOffice Trigger - create: Creating webhook', {
            webhookUrl,
            baseUrl,
            apiUrl,
            credentials: {
              baseUrl: credentials.baseUrl,
              hasToken: !!credentials.token,
            },
            requestBody,
          });
          
          const response = await this.helpers.httpRequest({
            method: 'POST',
            url: apiUrl,
            headers: {
              'Authorization': `Bearer ${credentials.token}`,
              'Accept': 'application/json',
              'Content-Type': 'application/json',
            },
            body: requestBody,
          });
          
          webhookData.webhookId = response.response.id;
          
          this.logger.info('OnlyOffice Trigger - create: Webhook created successfully', {
            webhookId: webhookData.webhookId,
            response: response.response,
          });
          
          return true;
        } catch (error) {
          const err = error as any;
          const credentials = await this.getCredentials('onlyOfficeApi');
          const baseUrl = credentials.baseUrl as string;
          
          // Extract response data safely to avoid circular references
          let responseData;
          try {
            responseData = typeof err.response === 'object' ? JSON.stringify(err.response) : err.response;
          } catch {
            responseData = 'Unable to serialize response';
          }
          
          this.logger.error('OnlyOffice Trigger - create: Failed to create webhook', {
            error: err.message,
            errorName: err.name,
            statusCode: err.statusCode,
            responseData,
            baseUrl: baseUrl,
            apiUrl: `${baseUrl}/api/2.0/settings/webhook`,
            credentialsBaseUrl: credentials.baseUrl,
          });
          throw error;
        }
      },
      async delete(this: IHookFunctions): Promise<boolean> {
        return true;
        const webhookData = this.getWorkflowStaticData('node');
        
        if (!webhookData.webhookId) {
          this.logger.debug('OnlyOffice Trigger - delete: No webhook ID found, skipping deletion');
          return true;
        }
        
        try {
          const credentials = await this.getCredentials('onlyOfficeApi');
          const baseUrl = credentials.baseUrl as string;
          
          this.logger.info('OnlyOffice Trigger - delete: Deleting webhook', {
            webhookId: webhookData.webhookId,
            baseUrl,
          });
          
          await this.helpers.httpRequest({
            method: 'DELETE',
            url: `${baseUrl}/api/2.0/settings/webhook/${webhookData.webhookId}`,
            headers: {
              'Authorization': `Bearer ${credentials.token}`,
              'Accept': 'application/json',
            },
          });
          
          this.logger.info('OnlyOffice Trigger - delete: Webhook deleted successfully', {
            webhookId: webhookData.webhookId,
          });
          
          delete webhookData.webhookId;
        } catch (error) {
          const err = error as any;
          this.logger.warn('OnlyOffice Trigger - delete: Failed to delete webhook (might already be deleted)', {
            error: err.message,
            statusCode: err.statusCode,
            webhookId: webhookData.webhookId,
          });
          // Don't throw error - webhook might already be deleted
          delete webhookData.webhookId;
        }
        
        return true;
      },
    },
    default: {
      async checkExists(this: IHookFunctions): Promise<boolean> {
        // Setup webhook is always available
        return true;
      },
      async create(this: IHookFunctions): Promise<boolean> {
        // Setup webhook doesn't need creation
        return true;
      },
      async delete(this: IHookFunctions): Promise<boolean> {
        // Setup webhook doesn't need deletion
        return true;
      },
    },
  };

  async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
    const req = this.getRequestObject();
    const webhookName = this.getWebhookName();
    
    // Log request info without circular references
    this.logger.info('OnlyOffice Trigger - Incoming request', {
      method: req.method,
      url: req.url,
      webhookName,
      headers: {
        'content-type': req.headers['content-type'],
        'user-agent': req.headers['user-agent'],
      },
    });

    // Handle HEAD request from OnlyOffice webhook verification
    // This should not trigger the workflow, just return 200 OK
    if (webhookName === 'setup' || req.method === 'HEAD') {
      this.logger.info('OnlyOffice Trigger - HEAD request received, responding with 200 OK');
      return {
        webhookResponse: {
          status: 200,
        },
      };
    }

    // Handle POST request with actual webhook data
    const body = this.getBodyData();
    
    this.logger.info('OnlyOffice Trigger - Processing webhook event', {
      method: req.method,
      bodyKeys: Object.keys(body || {}),
    });

    return {
      workflowData: [
        this.helpers.returnJsonArray([body]),
      ],
    };
  }
}

