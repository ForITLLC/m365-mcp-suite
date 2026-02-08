#!/usr/bin/env node

import 'dotenv/config';
import { parseArgs } from './cli.js';
import logger from './logger.js';
import AuthManager, { buildScopesFromEndpoints } from './auth.js';
import MicrosoftGraphServer from './server.js';
import { version } from './version.js';

async function main(): Promise<void> {
  try {
    const args = parseArgs();

    const includeWorkScopes = args.orgMode || false;
    if (includeWorkScopes) {
      logger.info('Organization mode enabled - including work account scopes');
    }

    const scopes = buildScopesFromEndpoints(includeWorkScopes, args.enabledTools);
    const authManager = new AuthManager(undefined, scopes);
    await authManager.loadTokenCache();

    if (args.login) {
      await authManager.acquireTokenByDeviceCode();
      logger.info('Login completed, testing connection with Graph API...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.verifyLogin) {
      logger.info('Verifying login...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.logout) {
      await authManager.logout();
      console.log(JSON.stringify({ message: 'Logged out successfully' }));
      process.exit(0);
    }

    if (args.listAccounts) {
      const accounts = await authManager.listAccounts();
      const result = accounts.map((account) => {
        const metadata = authManager.getAccountMetadata(account.homeAccountId);
        return {
          id: account.homeAccountId,
          username: account.username,
          name: account.name,
          appId: metadata?.appId || 'UNKNOWN',
          tenantId: metadata?.tenantId || 'UNKNOWN',
        };
      });
      console.log(JSON.stringify({ accounts: result }));
      process.exit(0);
    }

    if (args.removeAccount) {
      const success = await authManager.removeAccount(args.removeAccount);
      if (success) {
        console.log(JSON.stringify({ message: `Removed account: ${args.removeAccount}` }));
      } else {
        console.log(JSON.stringify({ error: `Account not found: ${args.removeAccount}` }));
        process.exit(1);
      }
      process.exit(0);
    }

    const server = new MicrosoftGraphServer(authManager, args);
    await server.initialize(version);
    await server.start();
  } catch (error) {
    logger.error(`Startup error: ${error}`);
    process.exit(1);
  }
}

main();
