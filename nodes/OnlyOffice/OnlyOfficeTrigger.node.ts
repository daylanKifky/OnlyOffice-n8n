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
    setup: {
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
    
    this.logger.info('OnlyOffice Trigger - Incoming request', {
      method: req.method,
      url: req.url,
      webhookName,
      headers: req.headers,
    });

    const requestInfo = {
      headers: req.headers,
      method: req.method,
      url: req.url,
      body: this.getBodyData(),
    };

    return {
      workflowData: [
        this.helpers.returnJsonArray(requestInfo),
      ],
    };

  }
}

