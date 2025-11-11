#!/usr/bin/env node

/**
 * Test script for SharePoint MCP Server
 * Run this to verify your setup before using with Claude Desktop
 */

import { spawn } from 'child_process';
import readline from 'readline';

console.log('SharePoint MCP Server - Test Mode\n');
console.log('This will start the MCP server and allow you to test tool calls manually.\n');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Start the MCP server
const serverProcess = spawn('node', ['index.js'], {
  stdio: ['pipe', 'pipe', 'inherit']
});

let requestId = 1;

// Function to send MCP requests
function sendRequest(method, params = {}) {
  const request = {
    jsonrpc: '2.0',
    id: requestId++,
    method: method,
    params: params
  };
  
  console.log('\n→ Sending request:', JSON.stringify(request, null, 2));
  serverProcess.stdin.write(JSON.stringify(request) + '\n');
}

// Listen for responses
let buffer = '';
serverProcess.stdout.on('data', (data) => {
  buffer += data.toString();
  
  // Try to parse complete JSON responses
  const lines = buffer.split('\n');
  buffer = lines.pop(); // Keep incomplete line in buffer
  
  for (const line of lines) {
    if (line.trim()) {
      try {
        const response = JSON.parse(line);
        console.log('\n← Received response:', JSON.stringify(response, null, 2));
      } catch (e) {
        console.log('\n← Received:', line);
      }
    }
  }
});

serverProcess.on('error', (error) => {
  console.error('Error starting server:', error);
  process.exit(1);
});

serverProcess.on('exit', (code) => {
  console.log(`\nServer exited with code ${code}`);
  process.exit(code);
});

// Interactive menu
function showMenu() {
  console.log('\n=== Test Menu ===');
  console.log('1. Initialize connection');
  console.log('2. List available tools');
  console.log('3. Test authenticate_sharepoint');
  console.log('4. Test set_site_url');
  console.log('5. Test search_files');
  console.log('6. Test get_folder_structure');
  console.log('7. Custom request');
  console.log('8. Exit');
  console.log('================\n');
  
  rl.question('Select option: ', (answer) => {
    handleMenuChoice(answer.trim());
  });
}

function handleMenuChoice(choice) {
  switch (choice) {
    case '1':
      sendRequest('initialize', {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: {
          name: 'test-client',
          version: '1.0.0'
        }
      });
      setTimeout(showMenu, 1000);
      break;
      
    case '2':
      sendRequest('tools/list');
      setTimeout(showMenu, 1000);
      break;
      
    case '3':
      rl.question('Enter Client ID: ', (clientId) => {
        rl.question('Enter Tenant ID: ', (tenantId) => {
          sendRequest('tools/call', {
            name: 'authenticate_sharepoint',
            arguments: {
              clientId: clientId.trim(),
              tenantId: tenantId.trim()
            }
          });
          setTimeout(showMenu, 2000);
        });
      });
      break;
      
    case '4':
      rl.question('Enter SharePoint Site URL: ', (siteUrl) => {
        sendRequest('tools/call', {
          name: 'set_site_url',
          arguments: {
            siteUrl: siteUrl.trim()
          }
        });
        setTimeout(showMenu, 1000);
      });
      break;
      
    case '5':
      rl.question('Enter search query: ', (query) => {
        sendRequest('tools/call', {
          name: 'search_files',
          arguments: {
            query: query.trim(),
            maxResults: 5
          }
        });
        setTimeout(showMenu, 2000);
      });
      break;
      
    case '6':
      sendRequest('tools/call', {
        name: 'get_folder_structure',
        arguments: {
          folderPath: '',
          depth: 2
        }
      });
      setTimeout(showMenu, 2000);
      break;
      
    case '7':
      rl.question('Enter method: ', (method) => {
        rl.question('Enter params (JSON): ', (params) => {
          try {
            const parsedParams = params.trim() ? JSON.parse(params) : {};
            sendRequest(method.trim(), parsedParams);
          } catch (e) {
            console.error('Invalid JSON:', e.message);
          }
          setTimeout(showMenu, 1000);
        });
      });
      break;
      
    case '8':
      console.log('\nExiting...');
      serverProcess.kill();
      rl.close();
      process.exit(0);
      break;
      
    default:
      console.log('Invalid option');
      showMenu();
  }
}

// Start with initialization
console.log('Starting MCP server...\n');
setTimeout(() => {
  showMenu();
}, 1000);

// Handle Ctrl+C
process.on('SIGINT', () => {
  console.log('\n\nShutting down...');
  serverProcess.kill();
  rl.close();
  process.exit(0);
});
