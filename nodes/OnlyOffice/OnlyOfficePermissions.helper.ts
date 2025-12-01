import { IExecuteFunctions, ILoadOptionsFunctions } from 'n8n-workflow';
import { parseApiKeyPermissions, PermissionResource } from './OnlyOfficeWebhook.types';

/**
 * Cache for permissions to avoid repeated API calls
 */
const permissionsCache = new Map<string, { permissions: Record<string, PermissionResource>; timestamp: number }>();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

/**
 * Checks if the API key has a specific permission
 */
export function hasPermission(
  permissions: Record<string, PermissionResource>,
  resource: string,
  action: 'read' | 'write' | null = null
): boolean {
  // Check for wildcard permissions
  if (permissions['*']) {
    const wildcard = permissions['*'];
    if (action === null) {
      return wildcard.read || wildcard.write;
    }
    if (action === 'read' && wildcard.read) return true;
    if (action === 'write' && wildcard.write) return true;
  }

  // Check for specific resource permission
  const resourceParts = resource.split('.');
  let currentLevel: Record<string, PermissionResource> | PermissionResource | undefined = permissions;

  // Navigate through nested structure
  for (let i = 0; i < resourceParts.length; i++) {
    const part = resourceParts[i];
    
    if (!currentLevel || typeof currentLevel !== 'object') {
      return false;
    }

    if (part in currentLevel) {
      const resourceObj = currentLevel[part];
      
      // If this is the last part, check the action flags
      if (i === resourceParts.length - 1) {
        if (action === null) {
          return resourceObj.read || resourceObj.write;
        }
        if (action === 'read' && resourceObj.read) return true;
        if (action === 'write' && resourceObj.write) return true;
        return false;
      }
      
      // Navigate deeper
      currentLevel = resourceObj as Record<string, PermissionResource>;
    } else {
      return false;
    }
  }

  return false;
}

/**
 * Fetches and caches API key permissions
 */
export async function getApiKeyPermissions(
  context: IExecuteFunctions | ILoadOptionsFunctions,
  credentialsKey: string = 'onlyOfficeApi'
): Promise<Record<string, PermissionResource>> {
  const credentials = await context.getCredentials(credentialsKey);
  const cacheKey = `${credentials.baseUrl}:${credentials.token}`;

  // Check cache
  const cached = permissionsCache.get(cacheKey);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.permissions;
  }

  // Fetch permissions
  const baseUrl = `${credentials.baseUrl}/api/2.0`;
  const response = await context.helpers.requestWithAuthentication.call(
    context,
    credentialsKey,
    {
      method: 'GET',
      url: `${baseUrl}/keys/permissions`,
    },
  );

  // Parse permissions
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

  // Cache the result
  permissionsCache.set(cacheKey, {
    permissions: parsedPermissions.permissions,
    timestamp: Date.now(),
  });

  return parsedPermissions.permissions;
}

/**
 * Validates that the API key has required permissions for an operation
 */
export async function validatePermissions(
  context: IExecuteFunctions | ILoadOptionsFunctions,
  resource: 'files' | 'folders' | 'accounts',
  action: 'read' | 'write',
  credentialsKey: string = 'onlyOfficeApi'
): Promise<void> {
  const permissions = await getApiKeyPermissions(context, credentialsKey);
  const hasRequiredPermission = hasPermission(permissions, resource, action);

  if (!hasRequiredPermission) {
    const requiredPermission = `${resource}:${action}`;
    const availablePermissions = Object.keys(permissions).join(', ') || 'none';
    
    throw new Error(
      `API key does not have required permission: ${requiredPermission}. ` +
      `Available permissions: ${availablePermissions}. ` +
      `Please update your API key permissions in OnlyOffice settings.`
    );
  }
}

/**
 * Gets a user-friendly error message for missing permissions
 */
export function getPermissionErrorMessage(
  resource: 'files' | 'folders' | 'accounts',
  action: 'read' | 'write',
  availablePermissions: Record<string, PermissionResource>
): string {
  const requiredPermission = `${resource}:${action}`;
  const permissionsList = Object.keys(availablePermissions).join(', ') || 'none';
  
  return (
    `This operation requires the "${requiredPermission}" permission, but your API key doesn't have it.\n\n` +
    `Required: ${requiredPermission}\n` +
    `Available: ${permissionsList}\n\n` +
    `To fix this:\n` +
    `1. Go to your OnlyOffice instance\n` +
    `2. Navigate to Settings → Integration → API\n` +
    `3. Edit your API key and add the "${requiredPermission}" permission\n` +
    `4. Save and try again`
  );
}

