#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import * as util from './util.js';
import { logToolCall } from './logger.js';

const server = new McpServer({
    name: "pnp-m365-mcp",
    version: "2.0.0",
});

// List all commands available in CLI for M365
server.registerTool(
    'm365_get_commands',
    {
        title: 'List available M365 CLI commands',
        description: 'Gets all CLI for Microsoft 365 commands. Use m365_get_command_docs for details on a specific command.',
        inputSchema: {}
    },
    async ({}) => {
        const commands = await util.getAllCommands();
        return {
            content: [
                { type: 'text', text: "TIP: Run m365_get_command_docs for detailed usage before executing commands." },
                { type: 'text', text: JSON.stringify(commands, null, 2) }
            ]
        };
    }
);

// Get documentation for a specific command
server.registerTool(
    'm365_get_command_docs',
    {
        title: 'Get command documentation',
        description: 'Gets detailed documentation for a CLI for M365 command including examples and options.',
        inputSchema: {
            commandName: z.string().describe('Command name (e.g., "spo site list")'),
            docs: z.string().describe('Documentation file path from m365_get_commands')
        }
    },
    async ({ commandName, docs }) => ({
        content: [{ type: 'text', text: await util.getCommandDocs(commandName, docs) }]
    })
);

// List connections - shows what's configured and what's logged in
server.registerTool(
    'm365_list_connections',
    {
        title: 'List M365 connections',
        description: 'Lists all configured connections from ~/.m365-connections.json and their login status. Shows which are available for pnp-m365.',
        inputSchema: {}
    },
    async ({}) => ({
        content: [{ type: 'text', text: await util.listConnections() }]
    })
);

// Validate a connection
server.registerTool(
    'm365_validate_connection',
    {
        title: 'Validate a connection',
        description: 'Checks if a specific connection is properly configured and logged in.',
        inputSchema: {
            connectionName: z.string().describe('Connection name from ~/.m365-connections.json (e.g., "Contoso", "Personal")')
        }
    },
    async ({ connectionName }) => ({
        content: [{ type: 'text', text: await util.validateConnection(connectionName) }]
    })
);

// Login to a connection
server.registerTool(
    'm365_login',
    {
        title: 'Login to M365',
        description: 'Authenticates to M365 using device code flow. Connection must be configured in ~/.m365-connections.json first.',
        inputSchema: {
            connectionName: z.string().describe('Connection name from ~/.m365-connections.json (e.g., "Contoso", "Personal")')
        }
    },
    async ({ connectionName }) => {
        const start = Date.now();
        let result: string | undefined;
        let error: string | undefined;
        try {
            result = await util.loginWithDeviceCode(connectionName);
            return { content: [{ type: 'text', text: result }] };
        } catch (e: any) {
            error = e.message || String(e);
            throw e;
        } finally {
            logToolCall('pnp-m365', 'm365_login', { connectionName }, connectionName, result?.substring(0, 100), error, Date.now() - start);
        }
    }
);

// Run a command - REQUIRES connectionName, no defaults
server.registerTool(
    'm365_run_command',
    {
        title: 'Run M365 CLI command',
        description: 'Executes a CLI for M365 command. connectionName is REQUIRED - there are no defaults.',
        inputSchema: {
            command: z.string().describe('The m365 command to run (e.g., "m365 spo site list")'),
            connectionName: z.string().describe('REQUIRED: Connection name (e.g., "Contoso"). Use m365_list_connections to see available.')
        }
    },
    async ({ command, connectionName }) => {
        const start = Date.now();
        let result: string | undefined;
        let error: string | undefined;
        try {
            result = await util.runCliCommand(command, connectionName);
            return { content: [{ type: 'text', text: result }] };
        } catch (e: any) {
            error = e.message || String(e);
            throw e;
        } finally {
            logToolCall('pnp-m365', 'm365_run_command', { command, connectionName }, connectionName, result?.substring(0, 100), error, Date.now() - start);
        }
    }
);

const transport = new StdioServerTransport();
await server.connect(transport);
