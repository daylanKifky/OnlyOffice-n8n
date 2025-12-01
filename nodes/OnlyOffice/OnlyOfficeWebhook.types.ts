import { IDataObject } from 'n8n-workflow';

export interface OnlyOfficeWebhookEvent {
  id: number;
  createOn: string;
  createBy: string;
  trigger: string;
  triggerId: number;
}

export interface OnlyOfficeWebhookPayload {
  id: number;
  parentId: number;
  folderIdDisplay: number;
  rootId: number;
  title: string;
  createBy: string;
  modifiedBy: string;
  createOn: string;
  modifiedOn: string;
  rootFolderType: number;
  rootCreateBy: string;
  fileEntryType: number;
}

export interface OnlyOfficeWebhook {
  id: number;
  name: string;
  url: string;
  triggers: string[];
}

export interface OnlyOfficeWebhookBody {
  event: OnlyOfficeWebhookEvent;
  payload: OnlyOfficeWebhookPayload;
  webhook: OnlyOfficeWebhook;
}

/**
 * Type guard function to validate OnlyOffice webhook body structure
 */
export function isOnlyOfficeWebhookBody(body: unknown): body is OnlyOfficeWebhookBody {
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return false;
  }

  const obj = body as IDataObject;

  // Check event object
  if (!obj.event || typeof obj.event !== 'object' || Array.isArray(obj.event)) {
    return false;
  }
  const event = obj.event as IDataObject;
  if (
    typeof event.id !== 'number' ||
    typeof event.createOn !== 'string' ||
    typeof event.createBy !== 'string' ||
    typeof event.trigger !== 'string' ||
    typeof event.triggerId !== 'number'
  ) {
    return false;
  }

  // Check payload object
  if (!obj.payload || typeof obj.payload !== 'object' || Array.isArray(obj.payload)) {
    return false;
  }
  const payload = obj.payload as IDataObject;
  if (
    typeof payload.id !== 'number' ||
    typeof payload.parentId !== 'number' ||
    typeof payload.folderIdDisplay !== 'number' ||
    typeof payload.rootId !== 'number' ||
    typeof payload.title !== 'string' ||
    typeof payload.createBy !== 'string' ||
    typeof payload.modifiedBy !== 'string' ||
    typeof payload.createOn !== 'string' ||
    typeof payload.modifiedOn !== 'string' ||
    typeof payload.rootFolderType !== 'number' ||
    typeof payload.rootCreateBy !== 'string' ||
    typeof payload.fileEntryType !== 'number'
  ) {
    return false;
  }

  // Check webhook object
  if (!obj.webhook || typeof obj.webhook !== 'object' || Array.isArray(obj.webhook)) {
    return false;
  }
  const webhook = obj.webhook as IDataObject;
  if (
    typeof webhook.id !== 'number' ||
    typeof webhook.name !== 'string' ||
    typeof webhook.url !== 'string' ||
    !Array.isArray(webhook.triggers) ||
    !webhook.triggers.every((t) => typeof t === 'string')
  ) {
    return false;
  }

  return true;
}

/**
 * Safely extracts the trigger from a webhook body
 */
export function getWebhookTrigger(body: unknown): string | undefined {
  if (isOnlyOfficeWebhookBody(body)) {
    return body.event.trigger;
  }
  if (body && typeof body === 'object' && !Array.isArray(body)) {
    const obj = body as IDataObject;
    const event = obj.event;
    if (event && typeof event === 'object' && !Array.isArray(event)) {
      const eventObj = event as IDataObject;
      return typeof eventObj.trigger === 'string' ? eventObj.trigger : undefined;
    }
  }
  return undefined;
}

/**
 * Safely extracts the title from a webhook body
 */
export function getWebhookTitle(body: unknown): string | undefined {
  if (isOnlyOfficeWebhookBody(body)) {
    return body.payload.title;
  }
  if (body && typeof body === 'object' && !Array.isArray(body)) {
    const obj = body as IDataObject;
    const payload = obj.payload;
    if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
      const payloadObj = payload as IDataObject;
      return typeof payloadObj.title === 'string' ? payloadObj.title : undefined;
    }
  }
  return undefined;
}

/**
 * Generates a warning log object for unexpected webhook body shapes
 */
export function getWebhookBodyShapeWarning(body: unknown, method: string): Record<string, unknown> {
  return {
    method,
    bodyKeys: Object.keys(body || {}),
    bodyType: typeof body,
    hasEvent: !!(body && typeof body === 'object' && !Array.isArray(body) && (body as IDataObject).event),
    hasPayload: !!(body && typeof body === 'object' && !Array.isArray(body) && (body as IDataObject).payload),
    hasWebhook: !!(body && typeof body === 'object' && !Array.isArray(body) && (body as IDataObject).webhook),
  };
}

