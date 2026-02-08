import type { AccountInfo, Configuration } from '@azure/msal-node';
import { PublicClientApplication } from '@azure/msal-node';
import logger from './logger.js';
import fs, { existsSync, readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

// Ok so this is a hack to lazily import keytar only when needed
// since --http mode may not need it at all, and keytar can be a pain to install (looking at you alpine)
let keytar: typeof import('keytar') | null = null;
async function getKeytar() {
  if (keytar === undefined) {
    return null;
  }
  if (keytar === null) {
    try {
      keytar = await import('keytar');
      return keytar;
    } catch (error) {
      logger.info('keytar not available, using file-based credential storage');
      keytar = undefined as any;
      return null;
    }
  }
  return keytar;
}

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  llmTip?: string;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

const endpoints = {
  default: endpointsData,
};

const SERVICE_NAME = 'ms-365-mcp-server';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const ACCOUNT_METADATA_KEY = 'account-metadata';
const FALLBACK_DIR = path.dirname(fileURLToPath(import.meta.url));
const FALLBACK_PATH = path.join(FALLBACK_DIR, '..', '.token-cache.json');
const ACCOUNT_METADATA_PATH = path.join(FALLBACK_DIR, '..', '.account-metadata.json');

// Maps accountId -> { appId, tenantId } for multi-tenant support
interface AccountMetadata {
  appId: string;
  tenantId: string;
}

const DEFAULT_CONFIG: Configuration = {
  auth: {
    clientId: process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e',
    authority: `https://login.microsoftonline.com/${process.env.MS365_MCP_TENANT_ID || 'common'}`,
  },
};

interface ScopeHierarchy {
  [key: string]: string[];
}

const SCOPE_HIERARCHY: ScopeHierarchy = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
};

function buildScopesFromEndpoints(
  includeWorkAccountScopes: boolean = false,
  enabledToolsPattern?: string
): string[] {
  const scopesSet = new Set<string>();

  // Create regex for tool filtering if pattern is provided
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Building scopes with tool filter pattern: ${enabledToolsPattern}`);
    } catch (error) {
      logger.error(
        `Invalid tool filter regex pattern: ${enabledToolsPattern}. Building scopes without filter.`
      );
    }
  }

  endpoints.default.forEach((endpoint) => {
    // Skip endpoints that don't match the tool filter
    if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
      return;
    }

    // Skip endpoints that only have workScopes if not in work mode
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }

    // Add regular scopes
    if (endpoint.scopes && Array.isArray(endpoint.scopes)) {
      endpoint.scopes.forEach((scope) => scopesSet.add(scope));
    }

    // Add workScopes if in work mode
    if (includeWorkAccountScopes && endpoint.workScopes && Array.isArray(endpoint.workScopes)) {
      endpoint.workScopes.forEach((scope) => scopesSet.add(scope));
    }
  });

  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (lowerScopes.every((scope) => scopesSet.has(scope))) {
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
      scopesSet.add(higherScope);
    }
  });

  const scopes = Array.from(scopesSet);
  if (enabledToolsPattern) {
    logger.info(`Built ${scopes.length} scopes for filtered tools: ${scopes.join(', ')}`);
  }

  return scopes;
}

interface LoginTestResult {
  success: boolean;
  message: string;
  userData?: {
    displayName: string;
    userPrincipalName: string;
  };
}

class AuthManager {
  private config: Configuration;
  private scopes: string[];
  private msalApp: PublicClientApplication;
  private accessToken: string | null;
  private tokenExpiry: number | null;
  private oauthToken: string | null;
  private isOAuthMode: boolean;
  private accountMetadata: Map<string, AccountMetadata> = new Map();

  constructor(
    config: Configuration = DEFAULT_CONFIG,
    scopes: string[] = buildScopesFromEndpoints()
  ) {
    logger.info(`And scopes are ${scopes.join(', ')}`, scopes);
    this.config = config;
    this.scopes = scopes;
    this.msalApp = new PublicClientApplication(this.config);
    this.accessToken = null;
    this.tokenExpiry = null;

    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;
  }

  async loadTokenCache(): Promise<void> {
    try {
      let cacheData: string | undefined;

      try {
        const kt = await getKeytar();
        if (kt) {
          const cachedData = await kt.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
          if (cachedData) {
            cacheData = cachedData;
          }
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain access failed, falling back to file storage: ${(keytarError as Error).message}`
        );
      }

      if (!cacheData && existsSync(FALLBACK_PATH)) {
        cacheData = readFileSync(FALLBACK_PATH, 'utf8');
      }

      if (cacheData) {
        this.msalApp.getTokenCache().deserialize(cacheData);
      }

      // Load account metadata (appId/tenantId per account)
      await this.loadAccountMetadata();
    } catch (error) {
      logger.error(`Error loading token cache: ${(error as Error).message}`);
    }
  }

  private async loadAccountMetadata(): Promise<void> {
    try {
      let metadataStr: string | undefined;

      try {
        const kt = await getKeytar();
        if (kt) {
          const cachedData = await kt.getPassword(SERVICE_NAME, ACCOUNT_METADATA_KEY);
          if (cachedData) {
            metadataStr = cachedData;
          }
        }
      } catch (keytarError) {
        logger.warn(`Keychain access failed for account metadata: ${(keytarError as Error).message}`);
      }

      if (!metadataStr && existsSync(ACCOUNT_METADATA_PATH)) {
        metadataStr = readFileSync(ACCOUNT_METADATA_PATH, 'utf8');
      }

      if (metadataStr) {
        const parsed = JSON.parse(metadataStr) as Record<string, AccountMetadata>;
        this.accountMetadata = new Map(Object.entries(parsed));
        logger.info(`Loaded metadata for ${this.accountMetadata.size} accounts`);
      }
    } catch (error) {
      logger.error(`Error loading account metadata: ${(error as Error).message}`);
    }
  }

  private async saveAccountMetadata(): Promise<void> {
    try {
      const metadataObj = Object.fromEntries(this.accountMetadata);
      const metadataStr = JSON.stringify(metadataObj);

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, ACCOUNT_METADATA_KEY, metadataStr);
        } else {
          fs.writeFileSync(ACCOUNT_METADATA_PATH, metadataStr, { mode: 0o600 });
        }
      } catch (keytarError) {
        logger.warn(`Keychain save failed for account metadata: ${(keytarError as Error).message}`);
        fs.writeFileSync(ACCOUNT_METADATA_PATH, metadataStr, { mode: 0o600 });
      }
    } catch (error) {
      logger.error(`Error saving account metadata: ${(error as Error).message}`);
    }
  }

  async saveTokenCache(): Promise<void> {
    try {
      const cacheData = this.msalApp.getTokenCache().serialize();

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cacheData);
        } else {
          fs.writeFileSync(FALLBACK_PATH, cacheData, { mode: 0o600 });
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain save failed, falling back to file storage: ${(keytarError as Error).message}`
        );

        fs.writeFileSync(FALLBACK_PATH, cacheData, { mode: 0o600 });
      }
    } catch (error) {
      logger.error(`Error saving token cache: ${(error as Error).message}`);
    }
  }

  async setOAuthToken(token: string): Promise<void> {
    this.oauthToken = token;
    this.isOAuthMode = true;
  }

  async getToken(forceRefresh = false): Promise<string | null> {
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }

    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }

    const currentAccount = await this.getCurrentAccount();

    if (currentAccount) {
      // Get the appId/tenantId that was used to authenticate this account
      const metadata = this.accountMetadata.get(currentAccount.homeAccountId);
      let msalApp = this.msalApp;

      if (metadata) {
        // Create MSAL app with the correct appId/tenantId for this account
        const customConfig: Configuration = {
          auth: {
            clientId: metadata.appId,
            authority: `https://login.microsoftonline.com/${metadata.tenantId}`,
          },
        };
        msalApp = new PublicClientApplication(customConfig);
        // Load the token cache
        const cacheData = this.msalApp.getTokenCache().serialize();
        msalApp.getTokenCache().deserialize(cacheData);
        logger.info(`Using custom app for token refresh: appId=${metadata.appId}, tenant=${metadata.tenantId}`);
      }

      const silentRequest = {
        account: currentAccount,
        scopes: this.scopes,
      };

      try {
        const response = await msalApp.acquireTokenSilent(silentRequest);
        this.accessToken = response.accessToken;
        this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
        return this.accessToken;
      } catch {
        logger.error('Silent token acquisition failed');
        throw new Error('Silent token acquisition failed');
      }
    }

    throw new Error('No valid token found');
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      return null;
    }

    // Only return account if there's exactly one - no "selection" concept
    // If multiple accounts exist, caller MUST use getTokenForAccount with explicit accountId
    if (accounts.length === 1) {
      return accounts[0];
    }

    // Multiple accounts - caller must specify which one
    return null;
  }

  async acquireTokenByDeviceCode(
    hack?: (message: string) => void,
    options?: { appId?: string; tenantId?: string }
  ): Promise<string | null> {
    // Use custom app/tenant if provided, otherwise use defaults
    let msalApp = this.msalApp;
    if (options?.appId || options?.tenantId) {
      const customConfig: Configuration = {
        auth: {
          clientId: options.appId || this.config.auth.clientId,
          authority: `https://login.microsoftonline.com/${options.tenantId || 'common'}`,
        },
      };
      msalApp = new PublicClientApplication(customConfig);
      // Load existing token cache into custom app
      try {
        const cacheData = this.msalApp.getTokenCache().serialize();
        msalApp.getTokenCache().deserialize(cacheData);
      } catch {
        // Ignore cache load errors for custom app
      }
    }

    const deviceCodeRequest = {
      scopes: this.scopes,
      deviceCodeCallback: (response: { message: string }) => {
        // Extract device code from message
        const codeMatch = response.message.match(/code\s+([A-Z0-9]+)\s+to authenticate/i);
        const deviceCode = codeMatch ? codeMatch[1] : 'CHECK MESSAGE BELOW';

        const clearText = `
████████████████████████████████████████████████████████████
██                                                        ██
██   DEVICE CODE: ${deviceCode.padEnd(12)}                       ██
██                                                        ██
██   Go to: https://microsoft.com/devicelogin            ██
██                                                        ██
████████████████████████████████████████████████████████████

${response.message}

After login run the "verify login" command
`;
        if (hack) {
          hack(clearText);
        } else {
          console.log(clearText);
        }
        logger.info('Device code login initiated');
      },
    };

    try {
      logger.info('Requesting device code...');
      logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
      logger.info(`Using app: ${options?.appId || this.config.auth.clientId}, tenant: ${options?.tenantId || 'default'}`);
      const response = await msalApp.acquireTokenByDeviceCode(deviceCodeRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
      logger.info('Device code login successful');
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;

      // If using custom app, copy tokens back to main cache
      if (options?.appId || options?.tenantId) {
        const newCacheData = msalApp.getTokenCache().serialize();
        this.msalApp.getTokenCache().deserialize(newCacheData);
      }

      // Save account metadata (appId/tenantId) for this account
      if (response?.account) {
        const accountId = response.account.homeAccountId;
        const appIdUsed = options?.appId || this.config.auth.clientId;
        const tenantIdUsed = options?.tenantId || 'common';
        this.accountMetadata.set(accountId, { appId: appIdUsed, tenantId: tenantIdUsed });
        await this.saveAccountMetadata();
        logger.info(`Saved metadata for account ${response.account.username}: appId=${appIdUsed}, tenant=${tenantIdUsed}`);
      }

      await this.saveTokenCache();
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in device code flow: ${(error as Error).message}`);
      throw error;
    }
  }

  async testLogin(): Promise<LoginTestResult> {
    try {
      logger.info('Testing login...');
      const token = await this.getToken();

      if (!token) {
        logger.error('Login test failed - no token received');
        return {
          success: false,
          message: 'Login failed - no token received',
        };
      }

      logger.info('Token retrieved successfully, testing Graph API access...');

      try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (response.ok) {
          const userData = await response.json();
          logger.info('Graph API user data fetch successful');
          return {
            success: true,
            message: 'Login successful',
            userData: {
              displayName: userData.displayName,
              userPrincipalName: userData.userPrincipalName,
            },
          };
        } else {
          const errorText = await response.text();
          logger.error(`Graph API user data fetch failed: ${response.status} - ${errorText}`);
          return {
            success: false,
            message: `Login successful but Graph API access failed: ${response.status}`,
          };
        }
      } catch (graphError) {
        logger.error(`Error fetching user data: ${(graphError as Error).message}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${(graphError as Error).message}`,
        };
      }
    } catch (error) {
      logger.error(`Login test failed: ${(error as Error).message}`);
      return {
        success: false,
        message: `Login failed: ${(error as Error).message}`,
      };
    }
  }

  async logout(): Promise<boolean> {
    try {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.msalApp.getTokenCache().removeAccount(account);
      }
      this.accessToken = null;
      this.tokenExpiry = null;

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
        }
      } catch (keytarError) {
        logger.warn(`Keychain deletion failed: ${(keytarError as Error).message}`);
      }

      if (fs.existsSync(FALLBACK_PATH)) {
        fs.unlinkSync(FALLBACK_PATH);
      }

      return true;
    } catch (error) {
      logger.error(`Error during logout: ${(error as Error).message}`);
      throw error;
    }
  }

  // Multi-account support methods
  async listAccounts(): Promise<AccountInfo[]> {
    return await this.msalApp.getTokenCache().getAllAccounts();
  }

  async removeAccount(accountId: string): Promise<boolean> {
    const accounts = await this.listAccounts();
    const account = accounts.find((acc: AccountInfo) => acc.homeAccountId === accountId);

    if (!account) {
      logger.error(`Account with ID ${accountId} not found`);
      return false;
    }

    try {
      await this.msalApp.getTokenCache().removeAccount(account);
      // Clear any cached token
      this.accessToken = null;
      this.tokenExpiry = null;
      logger.info(`Removed account: ${account.username} (${accountId})`);
      return true;
    } catch (error) {
      logger.error(`Failed to remove account ${accountId}: ${(error as Error).message}`);
      return false;
    }
  }

  /**
   * Get a token for a specific account by ID.
   * Uses the correct appId/tenantId that was used when the account was authenticated.
   */
  async getTokenForAccount(accountId: string): Promise<string | null> {
    const accounts = await this.listAccounts();
    const account = accounts.find((acc: AccountInfo) => acc.homeAccountId === accountId);

    if (!account) {
      logger.error(`Account with ID ${accountId} not found`);
      return null;
    }

    // Get the appId/tenantId that was used to authenticate this account
    const metadata = this.accountMetadata.get(accountId);
    let msalApp = this.msalApp;

    if (metadata) {
      // Create MSAL app with the correct appId/tenantId for this account
      const customConfig: Configuration = {
        auth: {
          clientId: metadata.appId,
          authority: `https://login.microsoftonline.com/${metadata.tenantId}`,
        },
      };
      msalApp = new PublicClientApplication(customConfig);
      // Load the token cache
      const cacheData = this.msalApp.getTokenCache().serialize();
      msalApp.getTokenCache().deserialize(cacheData);
      logger.info(`Using custom app for account ${account.username}: appId=${metadata.appId}, tenant=${metadata.tenantId}`);
    }

    const silentRequest = {
      account: account,
      scopes: this.scopes,
    };

    try {
      const response = await msalApp.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (error) {
      logger.error(`Failed to get token for account ${accountId}: ${(error as Error).message}`);
      return null;
    }
  }

  /**
   * Get account metadata (appId, tenantId) for a specific account
   */
  getAccountMetadata(accountId: string): AccountMetadata | undefined {
    return this.accountMetadata.get(accountId);
  }

  /**
   * Get all account metadata
   */
  getAllAccountMetadata(): Map<string, AccountMetadata> {
    return this.accountMetadata;
  }

  /**
   * Check for appId conflicts - returns conflicts if attempting to use different appId for same tenant
   */
  async checkAppIdConflicts(tenantId: string, appId: string): Promise<{
    hasConflict: boolean;
    existingAccounts?: Array<{
      username: string;
      accountId: string;
      existingAppId: string;
      existingTenant: string;
    }>;
  }> {
    const accounts = await this.listAccounts();
    const conflicts: Array<{
      username: string;
      accountId: string;
      existingAppId: string;
      existingTenant: string;
    }> = [];

    // Normalize tenant for comparison (handle both domain and GUID)
    const normalizedTenant = tenantId.toLowerCase();

    for (const account of accounts) {
      const metadata = this.accountMetadata.get(account.homeAccountId);
      if (!metadata) continue;

      const existingTenant = metadata.tenantId.toLowerCase();

      // Check if same tenant but different appId
      if (
        (existingTenant === normalizedTenant ||
         account.tenantId?.toLowerCase() === normalizedTenant ||
         account.username?.toLowerCase().endsWith(`@${normalizedTenant}`))
        && metadata.appId !== appId
      ) {
        conflicts.push({
          username: account.username || 'unknown',
          accountId: account.homeAccountId,
          existingAppId: metadata.appId,
          existingTenant: metadata.tenantId,
        });
      }
    }

    return {
      hasConflict: conflicts.length > 0,
      existingAccounts: conflicts.length > 0 ? conflicts : undefined,
    };
  }

  /**
   * Validate all accounts have metadata - returns accounts missing metadata
   */
  async validateAccountMetadata(): Promise<{
    valid: boolean;
    accountsWithMetadata: Array<{
      username: string;
      accountId: string;
      appId: string;
      tenantId: string;
    }>;
    accountsMissingMetadata: Array<{
      username: string;
      accountId: string;
    }>;
  }> {
    const accounts = await this.listAccounts();
    const withMetadata: Array<{
      username: string;
      accountId: string;
      appId: string;
      tenantId: string;
    }> = [];
    const missingMetadata: Array<{
      username: string;
      accountId: string;
    }> = [];

    for (const account of accounts) {
      const metadata = this.accountMetadata.get(account.homeAccountId);
      if (metadata) {
        withMetadata.push({
          username: account.username || 'unknown',
          accountId: account.homeAccountId,
          appId: metadata.appId,
          tenantId: metadata.tenantId,
        });
      } else {
        missingMetadata.push({
          username: account.username || 'unknown',
          accountId: account.homeAccountId,
        });
      }
    }

    return {
      valid: missingMetadata.length === 0,
      accountsWithMetadata: withMetadata,
      accountsMissingMetadata: missingMetadata,
    };
  }
}

export default AuthManager;
export { buildScopesFromEndpoints };
