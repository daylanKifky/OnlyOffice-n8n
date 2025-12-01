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

/**
 * Permission action flags
 */
export interface PermissionActions {
  read: boolean;
  write: boolean;
}

/**
 * Nested permission structure
 * Resources can have nested sub-resources (e.g., "account.self" becomes account -> self)
 * Using intersection type to allow both explicit properties and index signature
 */
export type PermissionResource = PermissionActions & {
  [subResource: string]: PermissionResource | boolean;
};

/**
 * API key permissions response with nested structure
 */
export interface OnlyOfficeApiKeyPermissionsResponse {
  permissions: Record<string, PermissionResource>;
  rawPermissions: string[];
  count: number;
}

/**
 * Parses an array of permission strings into nested structured format
 * Example: ["*:read", "account:read", "account.self:write"] becomes:
 * {
 *   "*": { read: true, write: false },
 *   "account": { read: true, write: false, "self": { read: false, write: true } }
 * }
 */
export function parseApiKeyPermissions(permissions: string[]): OnlyOfficeApiKeyPermissionsResponse {
  const result: Record<string, PermissionResource> = {};

  for (const permission of permissions) {
    // Handle wildcard permissions
    if (permission === '*') {
      if (!result['*']) {
        result['*'] = { read: false, write: false };
      }
      // '*' without action means both read and write
      result['*'].read = true;
      result['*'].write = true;
      continue;
    }

    // Parse permission format: "resource:action" or "resource.subresource:action"
    if (permission.includes(':')) {
      const [resourcePath, action] = permission.split(':', 2);
      const isRead = action === 'read';
      const isWrite = action === 'write';

      // Handle wildcard resource
      if (resourcePath === '*') {
        if (!result['*']) {
          result['*'] = { read: false, write: false };
        }
        if (isRead) result['*'].read = true;
        if (isWrite) result['*'].write = true;
        continue;
      }

      // Handle nested resources (e.g., "account.self")
      const resourceParts = resourcePath.split('.');
      let currentLevel: Record<string, PermissionResource> = result;

      // Navigate/create nested structure
      for (let i = 0; i < resourceParts.length; i++) {
        const part = resourceParts[i];
        const isLastPart = i === resourceParts.length - 1;

        if (!currentLevel[part]) {
          currentLevel[part] = { read: false, write: false };
        }

        if (isLastPart) {
          // Set the action flags on the final resource
          const resource = currentLevel[part];
          if (isRead) resource.read = true;
          if (isWrite) resource.write = true;
        } else {
          // Navigate deeper into nested structure
          // Ensure the nested resource exists and can hold sub-resources
          const existingResource = currentLevel[part];
          if (!existingResource || typeof existingResource === 'boolean') {
            currentLevel[part] = { read: false, write: false };
          }
          // Cast to Record to allow nested navigation
          currentLevel = currentLevel[part] as unknown as Record<string, PermissionResource>;
        }
      }
    } else {
      // Permission without action (e.g., just "resource")
      // Treat as both read and write
      if (!result[permission]) {
        result[permission] = { read: true, write: true };
      } else {
        result[permission].read = true;
        result[permission].write = true;
      }
    }
  }

  return {
    permissions: result,
    rawPermissions: permissions,
    count: permissions.length,
  };
}

