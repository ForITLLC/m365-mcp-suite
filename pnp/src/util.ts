import { exec, spawn } from 'child_process';
import path from 'path';
import { promises as fs } from 'fs';
import os from 'os';

// BLOCKED/HIDDEN COMMANDS - prevent accidental logout/removal
const BLOCKED_COMMANDS = ['logout', 'connection remove', 'connection use'];
const HIDDEN_COMMANDS = ['logout', 'connection remove'];

// Universal connection registry - shared across all M365 MCPs
const CONNECTIONS_FILE = path.join(os.homedir(), '.m365-connections.json');

// m365 CLI stores connections here (we READ from this, never write)
const CLI_CONNECTIONS_FILE = path.join(os.homedir(), '.cli-m365-all-connections.json');

const MCP_NAME = 'pnp-m365';

// Track active login processes to prevent zombies
const activeLoginProcesses = new Map<string, ReturnType<typeof spawn>>();

// Registry connection entry
interface ConnectionEntry {
    appId: string | null;
    tenant: string;
    description: string;
    mcps: string[];
    cliConnectionName: string | null;
    expectedEmail: string | null;
}

interface ConnectionsRegistry {
    connections: Record<string, ConnectionEntry>;
    _schema?: any;
}

// CLI connection from ~/.cli-m365-all-connections.json
interface CliConnection {
    name: string;
    identityName: string;
    identityId: string;
    identityTenantId: string;
    tenant: string;
    appId: string;
    authType: string;
    active?: boolean;
    accessTokens?: Record<string, any>;
}

async function loadRegistry(): Promise<ConnectionsRegistry> {
    try {
        const content = await fs.readFile(CONNECTIONS_FILE, 'utf-8');
        return JSON.parse(content);
    } catch {
        return { connections: {} };
    }
}

async function saveRegistry(registry: ConnectionsRegistry): Promise<void> {
    await fs.writeFile(CONNECTIONS_FILE, JSON.stringify(registry, null, 2));
}

async function loadCliConnections(): Promise<CliConnection[]> {
    try {
        const content = await fs.readFile(CLI_CONNECTIONS_FILE, 'utf-8');
        return JSON.parse(content);
    } catch {
        return [];
    }
}

// Find CLI connection that matches a registry entry
function findCliConnection(entry: ConnectionEntry, cliConnections: CliConnection[]): CliConnection | null {
    // First try by cliConnectionName if set
    if (entry.cliConnectionName) {
        const match = cliConnections.find(c => c.name === entry.cliConnectionName);
        if (match) return match;
    }

    // Find ALL CLI connections matching this appId+tenant
    const candidates = cliConnections.filter(c => {
        if (c.appId !== entry.appId) return false;
        if (c.tenant === entry.tenant) return true;
        // Multi-tenant app: check identityName domain
        if (c.identityName && entry.tenant) {
            const identityDomain = c.identityName.split('@')[1]?.toLowerCase();
            if (identityDomain === entry.tenant.toLowerCase()) return true;
        }
        return false;
    });

    if (candidates.length === 0) return null;
    if (candidates.length === 1) return candidates[0];

    // Multiple CLI connections for same appId+tenant (e.g. GA + Personal)
    // Use expectedEmail to pick the right one
    if (entry.expectedEmail) {
        const emailMatch = candidates.find(c =>
            c.identityName?.toLowerCase() === entry.expectedEmail!.toLowerCase()
        );
        if (emailMatch) return emailMatch;
    }

    // Fallback to first candidate
    return candidates[0];
}

// List all connections with their status
export async function listConnections(): Promise<string> {
    const registry = await loadRegistry();
    const cliConnections = await loadCliConnections();

    const entries = Object.entries(registry.connections)
        .filter(([_, entry]) => entry.mcps.includes(MCP_NAME));

    if (entries.length === 0) {
        return JSON.stringify({
            error: 'No connections configured for pnp-m365',
            hint: 'Add connections to ~/.m365-connections.json with "pnp-m365" in mcps array'
        }, null, 2);
    }

    const results = entries.map(([name, entry]) => {
        const cliConn = findCliConnection(entry, cliConnections);
        const connectedAs = cliConn?.identityName || null;
        const emailMismatch = entry.expectedEmail && connectedAs &&
            connectedAs.toLowerCase() !== entry.expectedEmail.toLowerCase();
        return {
            name,
            tenant: entry.tenant,
            appId: entry.appId,
            description: entry.description,
            loggedIn: !!cliConn,
            connectedAs,
            expectedEmail: entry.expectedEmail || null,
            warning: emailMismatch ? `WRONG ACCOUNT: expected ${entry.expectedEmail}, got ${connectedAs}` : null,
            cliConnectionName: cliConn?.name || null,
            needsSetup: !entry.appId ? 'Missing appId - needs app consent in tenant' : null
        };
    });

    return JSON.stringify(results, null, 2);
}

// Validate a specific connection
export async function validateConnection(connectionName: string): Promise<string> {
    const registry = await loadRegistry();
    const entry = registry.connections[connectionName];

    if (!entry) {
        const available = Object.keys(registry.connections).filter(
            k => registry.connections[k].mcps.includes(MCP_NAME)
        );
        return JSON.stringify({
            valid: false,
            error: `Connection "${connectionName}" not found`,
            available
        }, null, 2);
    }

    if (!entry.mcps.includes(MCP_NAME)) {
        return JSON.stringify({
            valid: false,
            error: `Connection "${connectionName}" not configured for pnp-m365`,
            configuredFor: entry.mcps
        }, null, 2);
    }

    if (!entry.appId) {
        return JSON.stringify({
            valid: false,
            error: `Connection "${connectionName}" has no appId configured`,
            hint: 'Add appId to ~/.m365-connections.json or get admin consent in tenant'
        }, null, 2);
    }

    const cliConnections = await loadCliConnections();
    const cliConn = findCliConnection(entry, cliConnections);

    if (!cliConn) {
        return JSON.stringify({
            valid: false,
            name: connectionName,
            tenant: entry.tenant,
            appId: entry.appId,
            error: 'Not logged in - run m365_login'
        }, null, 2);
    }

    return JSON.stringify({
        valid: true,
        name: connectionName,
        tenant: entry.tenant,
        appId: entry.appId,
        connectedAs: cliConn.identityName,
        cliConnectionName: cliConn.name
    }, null, 2);
}

// Login with device code - uses native CLI, no isolation nonsense
export async function loginWithDeviceCode(connectionName: string): Promise<string> {
    const registry = await loadRegistry();
    const entry = registry.connections[connectionName];

    if (!entry) {
        const available = Object.keys(registry.connections).filter(
            k => registry.connections[k].mcps.includes(MCP_NAME)
        );
        return JSON.stringify({
            error: `Connection "${connectionName}" not found in registry`,
            available,
            hint: 'Add connection to ~/.m365-connections.json first'
        }, null, 2);
    }

    if (!entry.appId) {
        return JSON.stringify({
            error: `Connection "${connectionName}" has no appId configured`,
            tenant: entry.tenant,
            hint: 'You need an app registration consented in this tenant. Add appId to ~/.m365-connections.json'
        }, null, 2);
    }

    // Kill any existing login process for this connection
    const existing = activeLoginProcesses.get(connectionName);
    if (existing && !existing.killed) {
        existing.kill();
        activeLoginProcesses.delete(connectionName);
    }

    // Build login command - NO M365_CLI_CONFIG_HOME, use native storage
    const loginCmd = `m365 login --authType deviceCode --appId ${entry.appId} --tenant ${entry.tenant}`;

    return new Promise((resolve) => {
        const subprocess = spawn(loginCmd, {
            shell: true,
            stdio: ['pipe', 'pipe', 'pipe']
        });

        activeLoginProcesses.set(connectionName, subprocess);

        // Clean up on exit (success or failure)
        subprocess.on('exit', () => {
            activeLoginProcesses.delete(connectionName);
            clearTimeout(killTimer);
        });

        // Kill after 15 minutes (device codes expire in 15 min)
        const killTimer = setTimeout(() => {
            if (!subprocess.killed) {
                subprocess.kill();
                activeLoginProcesses.delete(connectionName);
            }
        }, 15 * 60 * 1000);
        killTimer.unref();

        let output = '';
        let deviceCode = '';

        subprocess.stdout.on('data', (data) => {
            output += data.toString();
            const match = output.match(/enter the code ([A-Z0-9]+) to authenticate/i);
            if (match && !deviceCode) {
                deviceCode = match[1];
            }
        });

        subprocess.stderr.on('data', (data) => {
            output += data.toString();
            const match = output.match(/enter the code ([A-Z0-9]+) to authenticate/i);
            if (match && !deviceCode) {
                deviceCode = match[1];
            }
        });

        // Build sign-in hint
        const signInAs = entry.expectedEmail
            ? `SIGN IN AS: ${entry.expectedEmail}`
            : `Sign in with your @${entry.tenant} account`;

        // Poll for device code every 500ms, up to 15 seconds
        let attempts = 0;
        const maxAttempts = 30;
        const pollInterval = setInterval(() => {
            attempts++;
            if (deviceCode) {
                clearInterval(pollInterval);
                resolve(`
████████████████████████████████████████████████████████████
██                                                        ██
██   DEVICE CODE: ${deviceCode.padEnd(12)}                       ██
██                                                        ██
██   Go to: https://microsoft.com/devicelogin            ██
██                                                        ██
████████████████████████████████████████████████████████████

>>> ${signInAs} <<<

Connection: ${connectionName}
Tenant: ${entry.tenant}
App ID: ${entry.appId}
Description: ${entry.description}

Complete auth in browser. The CLI will store the token automatically.
Run m365_list_connections to verify after authenticating.
`);
            } else if (attempts >= maxAttempts) {
                clearInterval(pollInterval);
                resolve(JSON.stringify({
                    error: 'Failed to get device code within 15 seconds',
                    output: output.substring(0, 500)
                }, null, 2));
            }
        }, 500);
    });
}

// Run CLI command - ALWAYS requires connectionName, no defaults ever
export async function runCliCommand(command: string, connectionName: string): Promise<string> {
    // Validate connectionName is provided
    if (!connectionName) {
        const registry = await loadRegistry();
        const available = Object.keys(registry.connections).filter(
            k => registry.connections[k].mcps.includes(MCP_NAME)
        );
        return JSON.stringify({
            error: 'connectionName is REQUIRED - no defaults',
            available,
            hint: 'Every command must specify which connection to use'
        }, null, 2);
    }

    // Check for blocked commands
    const lowerCommand = command.toLowerCase();
    for (const blocked of BLOCKED_COMMANDS) {
        if (lowerCommand.includes(blocked)) {
            return `ERROR: '${blocked}' is blocked. Use CLI directly if needed.`;
        }
    }

    // Load registry and find connection
    const registry = await loadRegistry();
    const entry = registry.connections[connectionName];

    if (!entry) {
        const available = Object.keys(registry.connections).filter(
            k => registry.connections[k].mcps.includes(MCP_NAME)
        );
        return JSON.stringify({
            error: `Connection "${connectionName}" not found`,
            available
        }, null, 2);
    }

    if (!entry.mcps.includes(MCP_NAME)) {
        return JSON.stringify({
            error: `Connection "${connectionName}" not configured for pnp-m365`,
            configuredFor: entry.mcps
        }, null, 2);
    }

    // Find the CLI connection
    const cliConnections = await loadCliConnections();
    const cliConn = findCliConnection(entry, cliConnections);

    if (!cliConn) {
        return JSON.stringify({
            error: `Connection "${connectionName}" not logged in`,
            hint: `Run m365_login with connectionName="${connectionName}"`
        }, null, 2);
    }

    // Verify correct account is logged in
    if (entry.expectedEmail && cliConn.identityName &&
        cliConn.identityName.toLowerCase() !== entry.expectedEmail.toLowerCase()) {
        return JSON.stringify({
            error: `Wrong account logged in for "${connectionName}"`,
            expected: entry.expectedEmail,
            actual: cliConn.identityName,
            hint: `Logout and re-login with the correct account. Run m365_login with connectionName="${connectionName}"`
        }, null, 2);
    }

    let fullCommand = command;
    if (!fullCommand.includes('--output')) {
        fullCommand += ' --output json';
    }

    // Check if this connection is already active - if so, skip connection use
    // connection use outputs extra help text even on success
    let wrappedCommand = fullCommand;
    if (!cliConn.active) {
        // Only switch connections if needed
        wrappedCommand = `m365 connection use --name "${cliConn.name}" > /dev/null 2>&1 && ${fullCommand}`;
    }

    return new Promise((resolve, reject) => {
        exec(wrappedCommand, { timeout: 120000 }, (error, stdout, stderr) => {
            if (error) {
                resolve(`ERROR: ${stderr || error.message}`);
                return;
            }
            resolve(stdout.trim());
        });
    });
}

// Get command documentation
export async function getCommandDocs(commandName: string, docs: string): Promise<string> {
    try {
        const filePath = await checkGlobalPackage('@pnp/cli-microsoft365', `docs${path.sep}docs${path.sep}cmd${path.sep}${docs}`);
        if (!filePath) {
            throw new Error('@pnp/cli-microsoft365 package not found');
        }

        const fileContent = await fs.readFile(filePath, 'utf-8');
        return fileContent;
    } catch (error) {
        return `Failed to get docs for ${commandName}: ${error}`;
    }
}

// Get all available commands
export async function getAllCommands(): Promise<any[]> {
    try {
        const filePath = await checkGlobalPackage('@pnp/cli-microsoft365', 'allCommandsFull.json');
        if (!filePath) {
            throw new Error('@pnp/cli-microsoft365 package not found');
        }

        const fileContent = await fs.readFile(filePath, 'utf-8');
        const cliCommands = JSON.parse(fileContent);

        return cliCommands
            .filter((cmd: any) => !HIDDEN_COMMANDS.some(h => cmd.name.toLowerCase().includes(h)))
            .map((cmd: any) => ({
                name: `m365 ${cmd.name}`,
                description: cmd.description,
                docs: cmd.help
            }));
    } catch (error) {
        return [{ error: `Failed to get commands: ${error}` }];
    }
}

async function checkGlobalPackage(packageName: string, filePath: string): Promise<string | null> {
    return new Promise((resolve) => {
        exec('npm root -g', (err, npmRoot) => {
            if (err) {
                resolve(null);
                return;
            }
            const fullPath = path.join(npmRoot.trim(), packageName, filePath);
            fs.access(fullPath).then(() => resolve(fullPath)).catch(() => resolve(null));
        });
    });
}
