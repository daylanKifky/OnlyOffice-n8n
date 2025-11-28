import {
  IAuthenticateGeneric,
  ICredentialType,
  INodeProperties,
} from 'n8n-workflow';

export class OnlyOfficeApi implements ICredentialType {
  name = 'onlyOfficeApi';
  displayName = 'OnlyOffice API';
  documentationUrl = 'https://api.onlyoffice.com/';
  properties: INodeProperties[] = [
    {
      displayName: 'Base URL',
      name: 'baseUrl',
      type: 'string',
      default: '',
      placeholder: 'https://your-onlyoffice-instance.com',
      description: 'The base URL of your OnlyOffice instance',
      required: true,
    },
    {
      displayName: 'Access Token',
      name: 'token',
      type: 'string',
      typeOptions: {
        password: true,
      },
      default: '',
      description: 'The API token for authentication',
      required: true,
    },
  ];

  authenticate: IAuthenticateGeneric = {
    type: 'generic',
    properties: {
      headers: {
        Authorization: '=Bearer {{$credentials.token}}',
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    },
  };
}