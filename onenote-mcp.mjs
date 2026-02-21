#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Client } from '@microsoft/microsoft-graph-client';
import { DeviceCodeCredential } from '@azure/identity';
import { JSDOM } from 'jsdom';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import fetch from 'node-fetch';
import { z } from "zod";

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '.access-token.txt');
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Default: Microsoft Graph Explorer App ID
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'];

// --- Global State ---
let accessToken = null;
let graphClient = null;

// --- MCP Server Initialization ---
const server = new McpServer({
  name: 'onenote',
  version: '1.0.0', 
  description: 'OneNote MCP Server - Read, Write, and Edit OneNote content.'
});

// ============================================================================
// AUTHENTICATION & MICROSOFT GRAPH CLIENT MANAGEMENT
// ============================================================================

/**
 * Loads an existing access token from the local file system.
 */
function loadExistingToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
      try {
        const parsedToken = JSON.parse(tokenData); // New format: JSON object
        accessToken = parsedToken.token;
        console.error('Loaded existing token from file (JSON format).');
      } catch (parseError) {
        accessToken = tokenData; // Old format: plain token string
        console.error('Loaded existing token from file (plain text format).');
      }
    }
  } catch (error) {
    console.error(`Error loading token: ${error.message}`);
  }
}

/**
 * Initializes the Microsoft Graph client if an access token is available.
 * @returns {Client | null} The initialized Graph client or null.
 */
function initializeGraphClient() {
  if (accessToken && !graphClient) {
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
    console.error('Microsoft Graph client initialized.');
  }
  return graphClient;
}

/**
 * Ensures the Graph client is initialized and authenticated.
 * Loads token if not present, then initializes client.
 * @throws {Error} If no access token is available after attempting to load.
 * @returns {Promise<Client>} The initialized and authenticated Graph client.
 */
async function ensureGraphClient() {
  if (!accessToken) {
    loadExistingToken();
  }
  if (!accessToken) {
    throw new Error('No access token available. Please authenticate first using the "authenticate" tool.');
  }
  if (!graphClient) {
    initializeGraphClient();
  }
  return graphClient;
}

// ============================================================================
// HTML CONTENT PROCESSING UTILITIES
// ============================================================================

/**
 * Extracts readable plain text from HTML content.
 * Removes scripts, styles, and formats headings, paragraphs, lists, and tables.
 * @param {string} html - The HTML content string.
 * @returns {string} The extracted readable text.
 */
function extractReadableText(html) {
  try {
    if (!html) return '';
    const dom = new JSDOM(html);
    const document = dom.window.document;

    document.querySelectorAll('script, style').forEach(element => element.remove());

    let text = '';
    document.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(heading => {
      const headingText = heading.textContent?.trim();
      if (headingText) text += `\n${headingText}\n${'-'.repeat(headingText.length)}\n`;
    });
    document.querySelectorAll('p').forEach(paragraph => {
      const content = paragraph.textContent?.trim();
      if (content) text += `${content}\n\n`;
    });
    document.querySelectorAll('ul, ol').forEach(list => {
      text += '\n';
      list.querySelectorAll('li').forEach((item, index) => {
        const content = item.textContent?.trim();
        if (content) text += `${list.tagName === 'OL' ? index + 1 + '.' : '-'} ${content}\n`;
      });
      text += '\n';
    });
    document.querySelectorAll('table').forEach(table => {
      text += '\n📊 Table content:\n';
      table.querySelectorAll('tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'))
          .map(cell => cell.textContent?.trim())
          .join(' | ');
        if (cells.trim()) text += `${cells}\n`;
      });
      text += '\n';
    });

    if (!text.trim() && document.body) {
      text = document.body.textContent?.trim().replace(/\s+/g, ' ') || '';
    }
    return text.trim();
  } catch (error) {
    console.error(`Error extracting readable text: ${error.message}`);
    return 'Error: Could not extract readable text from HTML content.';
  }
}

/**
 * Extracts a short summary from HTML content.
 * @param {string} html - The HTML content string.
 * @param {number} [maxLength=300] - The maximum length of the summary.
 * @returns {string} A text summary.
 */
function extractTextSummary(html, maxLength = 300) {
  try {
    if (!html) return 'No content to summarize.';
    const dom = new JSDOM(html);
    const document = dom.window.document;
    const bodyText = document.body?.textContent?.trim().replace(/\s+/g, ' ') || '';
    if (!bodyText) return 'No text content found in HTML body.';
    const summary = bodyText.substring(0, maxLength);
    return summary.length < bodyText.length ? `${summary}...` : summary;
  } catch (error) {
    console.error(`Error extracting text summary: ${error.message}`);
    return 'Could not extract text summary.';
  }
}

/**
 * Converts plain text (with simple markdown) to HTML.
 * @param {string} text - The plain text to convert.
 * @returns {string} The HTML representation.
 */
function textToHtml(text) {
  if (!text) return '';
  if (text.includes('<html>') || text.includes('<!DOCTYPE html>')) return text; // Already HTML

  let html = String(text) // Ensure text is a string
    .replace(/&/g, '&').replace(/</g, '<').replace(/>/g, '>') // Basic HTML escaping first
    .replace(/```([\s\S]*?)```/g, (match, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/__(.*?)__/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>').replace(/_(.*?)_/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
    .replace(/^---+$/gm, '<hr>')
    .replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>')
    .replace(/^[\*\-\+] (.+)$/gm, '<li>$1</li>')
    .replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');

  html = html.split('\n').map(line => {
    const trimmed = line.trim();
    if (!trimmed) return '';
    if (/^<(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)/.test(trimmed) || /^<\/(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)>/.test(trimmed)) {
      return trimmed; // Already an HTML element we processed or a closing tag
    }
    return `<p>${trimmed}</p>`;
  }).filter(line => line).join('\n');

  html = html.replace(/(<li>.*?<\/li>(?:\s*<li>.*?<\/li>)*)/gs, '<ul>$1</ul>');
  html = html.replace(/(<blockquote>.*?<\/blockquote>(?:\s*<blockquote>.*?<\/blockquote>)*)/gs, '<blockquote>$1</blockquote>');
  
  return html;
}

// ============================================================================
// ONENOTE API UTILITIES
// ============================================================================

/**
 * Implements exponential backoff retry logic for rate limiting.
 * @param {Function} apiCall - The async function to execute.
 * @param {number} maxRetries - Maximum number of retries (default: 3).
 * @returns {Promise<any>} The result of the API call.
 */
async function retryWithBackoff(apiCall, maxRetries = 3) {
  let lastError;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await apiCall();
    } catch (error) {
      lastError = error;
      const is429 = error.statusCode === 429 || error.message?.includes('429') || error.message?.includes('throttled');
      if (is429 && attempt < maxRetries - 1) {
        const waitTime = Math.pow(2, attempt) * 1000; // 1s, 2s, 4s
        console.error(`Rate limit hit, waiting ${waitTime}ms before retry ${attempt + 1}/${maxRetries - 1}...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      } else {
        throw error;
      }
    }
  }
  throw lastError;
}

/**
 * Paginates through Microsoft Graph API results using @odata.nextLink.
 * @param {string} apiPath - The initial API path to query.
 * @param {number | null} maxPages - Maximum number of pages to fetch (null for unlimited).
 * @returns {Promise<Array>} Combined array of all results.
 */
async function paginateGraphRequest(apiPath, maxPages = null) {
  await ensureGraphClient();
  let allResults = [];
  let currentPath = apiPath;
  let pageCount = 0;

  while (currentPath && (maxPages === null || pageCount < maxPages)) {
    const response = await retryWithBackoff(() => graphClient.api(currentPath).get());
    
    if (response.value && Array.isArray(response.value)) {
      allResults = allResults.concat(response.value);
      pageCount++;
    }
    
    currentPath = response['@odata.nextLink'] || null;
    
    if (currentPath && pageCount % 10 === 0) {
      console.error(`Fetched ${pageCount} pages, ${allResults.length} items so far...`);
    }
  }
  
  return allResults;
}

/**
 * Fetches the content of a OneNote page.
 * @param {string} pageId - The ID of the page.
 * @param {'httpDirect' | 'direct'} [method='httpDirect'] - The method to use for fetching.
 * @returns {Promise<string>} The HTML content of the page.
 */
async function fetchPageContentAdvanced(pageId, method = 'httpDirect') {
  await ensureGraphClient();
  if (method === 'httpDirect') {
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!response.ok) throw new Error(`HTTP error fetching page content! Status: ${response.status} ${response.statusText}`);
    return await response.text();
  } else { // 'direct'
    return await graphClient.api(`/me/onenote/pages/${pageId}/content`).get();
  }
}

/**
 * Formats OneNote page information for display.
 * @param {object} page - The OneNote page object from Graph API.
 * @param {number | null} [index=null] - Optional index for numbered lists.
 * @param {boolean} [includeCreator=false] - Include creator/modifier information.
 * @returns {string} Formatted page information string.
 */
function formatPageInfo(page, index = null, includeCreator = false) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  const webUrl = page.links?.oneNoteWebUrl?.href || '';
  const urlLine = webUrl ? `
   🔗 <${webUrl}>` : '';
  
  let creatorLine = '';
  if (includeCreator && page.createdBy?.user?.displayName) {
    creatorLine = `
   Created by: ${page.createdBy.user.displayName}`;
  }
  if (includeCreator && page.lastModifiedBy?.user?.displayName) {
    creatorLine += `
   Modified by: ${page.lastModifiedBy.user.displayName}`;
  }
  
  return `${prefix}**${page.title}**
   ID: ${page.id}
   Created: ${new Date(page.createdDateTime).toLocaleDateString()}
   Modified: ${new Date(page.lastModifiedDateTime).toLocaleDateString()}${creatorLine}${urlLine}`;
}

// ============================================================================
// MCP TOOL DEFINITIONS
// ============================================================================

// --- Authentication Tools ---

server.tool(
  'authenticate',
  {
    // No input parameters expected for this tool
  },
  async () => {
    try {
      console.error('Starting device code authentication...');
      let deviceCodeInfo = null;
      const credential = new DeviceCodeCredential({
        clientId: clientId,
        userPromptCallback: (info) => {
          deviceCodeInfo = info;
          console.error(`\n=== AUTHENTICATION REQUIRED ===\n${info.message}\n================================\n`);
        }
      });

      const authPromise = credential.getToken(scopes);
      await new Promise(resolve => setTimeout(resolve, 2000)); // Allow time for userPromptCallback

      if (deviceCodeInfo) {
        const authMessage = `🔐 **AUTHENTICATION REQUIRED**

Please complete the following steps:
1. **Open this URL in your browser:** https://microsoft.com/devicelogin
2. **Enter this code:** ${deviceCodeInfo.userCode}
3. **Sign in with your Microsoft account that has OneNote access.**
4. **After completing authentication, use the 'saveAccessToken' tool.**

Token will be saved automatically upon successful browser authentication.`;

        authPromise.then(tokenResponse => {
          accessToken = tokenResponse.token;
          const tokenData = {
            token: accessToken,
            clientId: clientId,
            scopes: scopes,
            createdAt: new Date().toISOString(),
            expiresOn: tokenResponse.expiresOnTimestamp ? new Date(tokenResponse.expiresOnTimestamp).toISOString() : null
          };
          fs.writeFileSync(tokenFilePath, JSON.stringify(tokenData, null, 2));
          console.error('Token saved successfully!');
          initializeGraphClient();
        }).catch(error => {
          console.error(`Background authentication failed: ${error.message}`);
        });
        
        return { content: [{ type: 'text', text: authMessage }] };
      } else {
        return { isError: true, content: [{ type: 'text', text: 'Could not retrieve device code information. Please try again or check console logs.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Authentication failed: ${error.message}` }] };
    }
  }
);
// Note: For the above tool, the Zod schema `z.object({}).describe(...)` was simplified to `{}` as per the user's specific finding
// about the SDK's `server.tool(name, {param: z.type()}, handler)` signature.
// If the SDK *does* support a top-level describe on the Zod object itself, that would be:
// `z.object({}).describe('Start the authentication flow...')`

server.tool(
  'saveAccessToken',
  {
    // No input parameters
  },
  async () => {
    try {
      loadExistingToken();
      if (accessToken) {
        initializeGraphClient();
        const testResponse = await graphClient.api('/me').get();
        return {
          content: [{
            type: 'text',
            text: `✅ **Authentication Successful!**
Token loaded and verified.
**Account Info:**
- Name: ${testResponse.displayName || 'Unknown'}
- Email: ${testResponse.userPrincipalName || 'Unknown'}
🚀 You can now use OneNote tools!`
          }]
        };
      } else {
        return { isError: true, content: [{ type: 'text', text: `❌ **No Token Found.** Please run the 'authenticate' tool first.` }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to load or verify token: ${error.message}` }] };
    }
  }
);

// --- Page Reading Tools ---

server.tool(
  'listNotebooks',
  {
    // No input parameters
  },
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api('/me/onenote/notebooks').get();
      if (response.value && response.value.length > 0) {
        const notebookList = response.value.map((nb, i) => formatPageInfo(nb, i)).join('\n\n');
        return { content: [{ type: 'text', text: `📚 **Your OneNote Notebooks** (${response.value.length} found):\n\n${notebookList}` }] };
      } else {
        return { content: [{ type: 'text', text: '📚 No OneNote notebooks found.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: error.message.includes('authenticate') ? '🔐 Authentication Required. Run `authenticate` tool.' : `Failed to list notebooks: ${error.message}` }] };
    }
  }
);

server.tool(
  'listSections',
  {
    notebookId: z.string().describe('The ID of the notebook to list sections from.')
  },
  async ({ notebookId }) => {
    try {
      await ensureGraphClient();
      console.error(`Fetching sections for notebook ID: ${notebookId}`);
      const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebookId}/sections`);
      
      if (sections.length > 0) {
        const sectionList = sections.map((section, i) => `${i + 1}. **${section.displayName}**\n   ID: ${section.id}\n   Created: ${new Date(section.createdDateTime).toLocaleDateString()}`).join('\n\n');
        return { content: [{ type: 'text', text: `📂 **Sections in Notebook** (${sections.length} found):\n\n${sectionList}` }] };
      } else {
        return { content: [{ type: 'text', text: '📂 No sections found in this notebook.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to list sections: ${error.message}` }] };
    }
  }
);

server.tool(
  'listPagesInSection',
  {
    sectionId: z.string().describe('The ID of the section to list pages from.'),
    top: z.number().min(1).max(100).default(20).describe('Number of pages to return per request (max 100).').optional(),
    orderBy: z.enum(['created', 'modified']).default('modified').describe('Sort by created or modified date.').optional()
  },
  async ({ sectionId, top, orderBy }) => {
    try {
      await ensureGraphClient();
      const orderField = orderBy === 'created' ? 'createdDateTime' : 'lastModifiedDateTime';
      console.error(`Fetching pages from section ID: ${sectionId} (top: ${top}, orderBy: ${orderField})`);
      
      const apiPath = `/me/onenote/sections/${sectionId}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=${top}&$orderby=${orderField} desc`;
      const pages = await paginateGraphRequest(apiPath);
      
      if (pages.length > 0) {
        const pageList = pages.slice(0, 50).map((page, i) => formatPageInfo(page, i)).join('\n\n');
        const morePages = pages.length > 50 ? `\n\n... and ${pages.length - 50} more pages. Total: ${pages.length}` : '';
        return { content: [{ type: 'text', text: `📄 **Pages in Section** (${pages.length} found):\n\n${pageList}${morePages}` }] };
      } else {
        return { content: [{ type: 'text', text: '📄 No pages found in this section.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to list pages in section: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPages',
  {
    query: z.string().describe('The search term for page titles.').optional()
  },
  async ({ query }) => {
    try {
      await ensureGraphClient();
      console.error('Searching pages with pagination...');
      const pages = await paginateGraphRequest('/me/onenote/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links');
      
      let filteredPages = pages;
      if (query) {
        const searchTerm = query.toLowerCase();
        filteredPages = pages.filter(page => page.title && page.title.toLowerCase().includes(searchTerm));
      }
      
      if (filteredPages.length > 0) {
        const pageList = filteredPages.slice(0, 10).map((page, i) => formatPageInfo(page, i)).join('\n\n');
        const morePages = filteredPages.length > 10 ? `\n\n... and ${filteredPages.length - 10} more pages.` : '';
        const tip = filteredPages.length > 100 ? '\n\n💡 Tip: Use `searchPagesByDate` or `searchInNotebook` for more efficient filtered searches.' : '';
        return { content: [{ type: 'text', text: `🔍 **Search Results** ${query ? `for "${query}"` : ''} (${filteredPages.length} found):\n\n${pageList}${morePages}${tip}` }] };
      } else {
        return { content: [{ type: 'text', text: query ? `🔍 No pages found matching "${query}".` : '📄 No pages found.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search pages: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPagesByDate',
  {
    days: z.number().min(1).default(1).describe('Number of days back to search (1 = today only).').optional(),
    query: z.string().describe('Optional keyword to filter page titles.').optional(),
    dateField: z.enum(['created', 'modified', 'both']).default('both').describe('Filter by created, modified, or both dates.').optional(),
    includeContent: z.boolean().default(false).describe('Include page content preview (slower).').optional(),
    createdBy: z.string().describe('Optional: filter by creator name or email (e.g., "Josue" or "josue@elon.edu").').optional(),
    modifiedBy: z.string().describe('Optional: filter by modifier name or email (e.g., "Josue" or "josue@elon.edu").').optional()
  },
  async ({ days, query, dateField, includeContent, createdBy, modifiedBy }) => {
    try {
      await ensureGraphClient();
      const threshold = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
      console.error(`Searching pages from last ${days} day(s) (since ${threshold.toLocaleString()})...`);
      
      // Get all notebooks
      const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
      const notebooks = notebooksResponse.value || [];
      
      if (notebooks.length === 0) {
        return { content: [{ type: 'text', text: '📚 No notebooks found.' }] };
      }
      
      console.error(`Found ${notebooks.length} notebook(s), fetching sections...`);
      let allMatchingPages = [];
      let totalSectionsChecked = 0;
      
      // Iterate through notebooks and sections
      for (const notebook of notebooks) {
        const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
        totalSectionsChecked += sections.length;
        
        for (const section of sections) {
          try {
            const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$expand=lastModifiedBy,createdBy&$top=100`;
            const pages = await paginateGraphRequest(apiPath);
            
            // Filter by date
            const matchingPages = pages.filter(page => {
              const created = new Date(page.createdDateTime);
              const modified = new Date(page.lastModifiedDateTime);
              
              let dateMatch = false;
              if (dateField === 'created') {
                dateMatch = created >= threshold;
              } else if (dateField === 'modified') {
                dateMatch = modified >= threshold;
              } else { // 'both'
                dateMatch = created >= threshold || modified >= threshold;
              }
              
              if (!dateMatch) return false;
              
              // Optional keyword filter
              if (query) {
                const searchTerm = query.toLowerCase();
                if (!page.title || !page.title.toLowerCase().includes(searchTerm)) {
                  return false;
                }
              }
              
              // Optional createdBy filter
              if (createdBy) {
                const filterTerm = createdBy.toLowerCase();
                const creatorName = page.createdBy?.user?.displayName?.toLowerCase() || '';
                const creatorEmail = page.createdBy?.user?.email?.toLowerCase() || '';
                if (!creatorName.includes(filterTerm) && !creatorEmail.includes(filterTerm)) {
                  return false;
                }
              }
              
              // Optional modifiedBy filter
              if (modifiedBy) {
                const filterTerm = modifiedBy.toLowerCase();
                const modifierName = page.lastModifiedBy?.user?.displayName?.toLowerCase() || '';
                const modifierEmail = page.lastModifiedBy?.user?.email?.toLowerCase() || '';
                if (!modifierName.includes(filterTerm) && !modifierEmail.includes(filterTerm)) {
                  return false;
                }
              }
              
              return true;
            });
            
            // Add notebook and section context
            matchingPages.forEach(page => {
              page._notebook = notebook.displayName || notebook.name;
              page._section = section.displayName || section.name;
            });
            
            allMatchingPages = allMatchingPages.concat(matchingPages);
          } catch (sectionError) {
            console.error(`Error fetching pages from section ${section.displayName}: ${sectionError.message}`);
          }
        }
        
        console.error(`Checked notebook "${notebook.displayName}", ${allMatchingPages.length} matches so far...`);
      }
      
      console.error(`Search complete: ${allMatchingPages.length} matches from ${totalSectionsChecked} sections`);
      
      if (allMatchingPages.length === 0) {
        return { content: [{ type: 'text', text: query ? `🔍 No pages found matching "${query}" in the last ${days} day(s).` : `📄 No pages found in the last ${days} day(s).` }] };
      }
      
      // Sort by most recent first
      allMatchingPages.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      // Format results
      let resultText = `🔍 **Date Search Results** (${allMatchingPages.length} found in last ${days} day(s)):\n`;
      if (query) resultText += `Keyword: "${query}"\n`;
      if (createdBy) resultText += `Created by: "${createdBy}"\n`;
      if (modifiedBy) resultText += `Modified by: "${modifiedBy}"\n`;
      resultText += `Checked: ${notebooks.length} notebooks, ${totalSectionsChecked} sections\n\n`;
      
      const displayPages = allMatchingPages.slice(0, 20);
      for (let i = 0; i < displayPages.length; i++) {
        const page = displayPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   📚 ${page._notebook} / 📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Created: ${new Date(page.createdDateTime).toLocaleString()}`;
        if (page.createdBy?.user?.displayName) {
          resultText += ` by ${page.createdBy.user.displayName}`;
        }
        resultText += `\n   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}`;
        if (page.lastModifiedBy?.user?.displayName) {
          resultText += ` by ${page.lastModifiedBy.user.displayName}`;
        }
        
        if (includeContent) {
          try {
            const htmlContent = await fetchPageContentAdvanced(page.id, 'httpDirect');
            const preview = extractTextSummary(htmlContent, 200);
            resultText += `\n   Preview: ${preview}`;
          } catch (contentError) {
            resultText += `\n   Preview: (error loading content)`;
          }
        }
        
        resultText += '\n\n';
      }
      
      if (allMatchingPages.length > 20) {
        resultText += `... and ${allMatchingPages.length - 20} more pages.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search pages by date: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPageContent',
  {
    query: z.string().describe('Text to search for within page content.'),
    days: z.number().min(1).describe('Optional: limit to pages from last N days.').optional(),
    notebookId: z.string().describe('Optional: limit to specific notebook.').optional(),
    maxPages: z.number().min(1).max(50).default(20).describe('Max number of pages to search (default 20).').optional()
  },
  async ({ query, days, notebookId, maxPages }) => {
    try {
      await ensureGraphClient();
      const searchTerm = query.toLowerCase();
      const threshold = days ? new Date(Date.now() - days * 24 * 60 * 60 * 1000) : null;
      console.error(`Searching page content for "${query}"...`);
      
      let notebooks = [];
      if (notebookId) {
        const nb = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
        notebooks = [nb];
      } else {
        const response = await graphClient.api('/me/onenote/notebooks').get();
        notebooks = response.value || [];
      }
      
      let matchingPages = [];
      let pagesSearched = 0;
      
      for (const notebook of notebooks) {
        const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
        
        for (const section of sections) {
          if (matchingPages.length >= maxPages) break;
          
          try {
            const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=20`;
            const pages = await paginateGraphRequest(apiPath);
            
            for (const page of pages) {
              if (matchingPages.length >= maxPages) break;
              
              // Date filter if specified
              if (threshold) {
                const modified = new Date(page.lastModifiedDateTime);
                if (modified < threshold) continue;
              }
              
              pagesSearched++;
              
              // Fetch and search content
              try {
                const htmlContent = await fetchPageContentAdvanced(page.id, 'httpDirect');
                const textContent = extractReadableText(htmlContent).toLowerCase();
                
                if (textContent.includes(searchTerm)) {
                  page._notebook = notebook.displayName || notebook.name;
                  page._section = section.displayName || section.name;
                  page._snippet = extractSnippet(textContent, searchTerm, 150);
                  matchingPages.push(page);
                }
              } catch (contentError) {
                console.error(`Error reading page ${page.title}: ${contentError.message}`);
              }
            }
          } catch (sectionError) {
            console.error(`Error in section ${section.displayName}: ${sectionError.message}`);
          }
        }
        
        if (matchingPages.length >= maxPages) break;
      }
      
      console.error(`Content search complete: ${matchingPages.length} matches from ${pagesSearched} pages searched`);
      
      if (matchingPages.length === 0) {
        return { content: [{ type: 'text', text: `🔍 No pages found containing "${query}" in their content.\nSearched ${pagesSearched} pages.` }] };
      }
      
      let resultText = `🔍 **Content Search Results** (${matchingPages.length} matches for "${query}"):\n`;
      resultText += `Searched ${pagesSearched} pages\n\n`;
      
      for (let i = 0; i < matchingPages.length; i++) {
        const page = matchingPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   📚 ${page._notebook} / 📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n`;
        resultText += `   Snippet: "...${page._snippet}..."\n\n`;
      }
      
      if (matchingPages.length >= maxPages) {
        resultText += `\n💡 Showing first ${maxPages} matches. Use 'maxPages' parameter or narrow search with 'days' or 'notebookId'.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search page content: ${error.message}` }] };
    }
  }
);

/**
 * Extracts a snippet around the search term.
 * @param {string} text - The full text to search.
 * @param {string} searchTerm - The term to find.
 * @param {number} maxLength - Maximum snippet length.
 * @returns {string} The snippet.
 */
function extractSnippet(text, searchTerm, maxLength = 150) {
  const index = text.indexOf(searchTerm);
  if (index === -1) return text.substring(0, maxLength);
  
  const start = Math.max(0, index - Math.floor(maxLength / 2));
  const end = Math.min(text.length, start + maxLength);
  let snippet = text.substring(start, end).trim();
  
  if (start > 0) snippet = '...' + snippet;
  if (end < text.length) snippet = snippet + '...';
  
  return snippet;
}

server.tool(
  'getMyRecentChanges',
  {
    sinceDate: z.string().describe('Start date in YYYY-MM-DD format, or "monday" for last Monday.').optional(),
    days: z.number().min(1).describe('Alternative: number of days back to search (e.g., 3 for last 3 days).').optional(),
    notebookId: z.string().describe('Optional: limit to specific notebook.').optional(),
    includeCreator: z.boolean().default(false).describe('Show who created/modified each page.').optional()
  },
  async ({ sinceDate, days, notebookId, includeCreator }) => {
    try {
      await ensureGraphClient();
      
      // Get current user info
      const userInfo = await graphClient.api('/me').get();
      const currentUserEmail = userInfo.userPrincipalName?.toLowerCase();
      
      // Calculate threshold date
      let threshold;
      if (sinceDate === 'monday') {
        const today = new Date();
        const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ...
        const daysBack = dayOfWeek === 0 ? 6 : dayOfWeek === 1 ? 0 : dayOfWeek - 1;
        threshold = new Date(today.getTime() - daysBack * 24 * 60 * 60 * 1000);
        threshold.setHours(0, 0, 0, 0);
      } else if (sinceDate) {
        threshold = new Date(sinceDate);
      } else if (days) {
        threshold = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
      } else {
        // Default to last 3 days
        threshold = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
      }
      
      console.error(`Finding your changes since ${threshold.toLocaleDateString()}...`);
      
      // Get notebooks
      let notebooks = [];
      if (notebookId) {
        const nb = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
        notebooks = [nb];
      } else {
        const response = await graphClient.api('/me/onenote/notebooks').get();
        notebooks = response.value || [];
      }
      
      let myChanges = [];
      
      for (const notebook of notebooks) {
        const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
        
        for (const section of sections) {
          try {
            const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$expand=lastModifiedBy,createdBy&$top=100`;
            const pages = await paginateGraphRequest(apiPath);
            
            const myPages = pages.filter(page => {
              const modified = new Date(page.lastModifiedDateTime);
              if (modified < threshold) return false;
              
              // Check if current user modified it
              const modifierEmail = page.lastModifiedBy?.user?.email?.toLowerCase();
              return modifierEmail === currentUserEmail;
            });
            
            myPages.forEach(page => {
              page._notebook = notebook.displayName || notebook.name;
              page._section = section.displayName || section.name;
            });
            
            myChanges = myChanges.concat(myPages);
          } catch (sectionError) {
            console.error(`Error in section ${section.displayName}: ${sectionError.message}`);
          }
        }
      }
      
      console.error(`Found ${myChanges.length} pages you modified`);
      
      if (myChanges.length === 0) {
        return { content: [{ type: 'text', text: `📝 No pages found that you modified since ${threshold.toLocaleDateString()}.` }] };
      }
      
      // Sort by most recent
      myChanges.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      let resultText = `📝 **Your Recent Changes** (${myChanges.length} pages since ${threshold.toLocaleDateString()}):\n\n`;
      
      // Group by notebook
      const byNotebook = {};
      myChanges.forEach(page => {
        const nb = page._notebook;
        if (!byNotebook[nb]) byNotebook[nb] = [];
        byNotebook[nb].push(page);
      });
      
      for (const [notebookName, pages] of Object.entries(byNotebook)) {
        resultText += `\n📚 **${notebookName}** (${pages.length} pages)\n`;
        
        pages.forEach((page, i) => {
          const webUrl = page.links?.oneNoteWebUrl?.href || '';
          resultText += `\n${i + 1}. **${page.title}**\n`;
          resultText += `   📂 ${page._section}\n`;
          if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
          resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n`;
          if (includeCreator && page.createdBy?.user?.displayName) {
            resultText += `   Created by: ${page.createdBy.user.displayName}\n`;
          }
        });
      }
      
      resultText += `\n\n💡 **Standup Tip:** Use this list to fill out your "Async" section!`;
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get recent changes: ${error.message}` }] };
    }
  }
);

server.tool(
  'createDailyNote',
  {
    notebookName: z.string().describe('Name of the notebook (e.g., "SQLNikon").'),
    sectionName: z.string().describe('Name of the section (e.g., "Elon").'),
    date: z.string().describe('Date in YYYY-MM-DD format, or "today" for current date.').optional(),
    title: z.string().describe('Optional custom title (defaults to M/D/YY format).').optional(),
    content: z.string().describe('Optional initial content for the note.').optional()
  },
  async ({ notebookName, sectionName, date, title, content }) => {
    try {
      await ensureGraphClient();
      
      // Determine the date
      let noteDate;
      if (date === 'today' || !date) {
        noteDate = new Date();
      } else {
        noteDate = new Date(date);
      }
      
      // Format title as M/D/YY (e.g., "2/20/26")
      const defaultTitle = `${noteDate.getMonth() + 1}/${noteDate.getDate()}/${noteDate.getFullYear().toString().slice(-2)}`;
      const pageTitle = title || defaultTitle;
      
      console.error(`Creating daily note "${pageTitle}" in ${notebookName}/${sectionName}...`);
      
      // Find the notebook
      const notebooks = await graphClient.api('/me/onenote/notebooks').get();
      const notebook = (notebooks.value || []).find(nb => 
        nb.displayName?.toLowerCase() === notebookName.toLowerCase()
      );
      
      if (!notebook) {
        return { isError: true, content: [{ type: 'text', text: `❌ Notebook "${notebookName}" not found. Available notebooks:\n${(notebooks.value || []).map(nb => `  - ${nb.displayName}`).join('\n')}` }] };
      }
      
      // Find the section
      const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
      const section = sections.find(s => 
        s.displayName?.toLowerCase() === sectionName.toLowerCase()
      );
      
      if (!section) {
        return { isError: true, content: [{ type: 'text', text: `❌ Section "${sectionName}" not found in notebook "${notebookName}". Available sections:\n${sections.map(s => `  - ${s.displayName}`).join('\n')}` }] };
      }
      
      // Check if page already exists
      const existingPages = await paginateGraphRequest(`/me/onenote/sections/${section.id}/pages?$select=id,title`);
      const existingPage = existingPages.find(p => p.title === pageTitle);
      
      if (existingPage) {
        return { content: [{ type: 'text', text: `ℹ️ **Note Already Exists**\n\nA page titled "${pageTitle}" already exists in ${notebookName}/${sectionName}.\n\nPage ID: ${existingPage.id}\n\nUse 'appendToPage' or 'updatePageContent' to modify it.` }] };
      }
      
      // Create the page
      const htmlContent = content ? textToHtml(content) : '<p>Notes for today...</p>';
      const pageHtml = `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(pageTitle)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(pageTitle)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${section.id}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);
      
      const webUrl = response.links?.oneNoteWebUrl?.href || '';
      
      return {
        content: [{
          type: 'text',
          text: `✅ **Daily Note Created!**\n\n**Title:** ${response.title}\n**Notebook:** ${notebookName}\n**Section:** ${sectionName}\n**Page ID:** ${response.id}\n**Created:** ${new Date(response.createdDateTime).toLocaleString()}${webUrl ? `\n\n🔗 <${webUrl}>` : ''}`
        }]
      };
    } catch (error) {
      console.error(`CREATE DAILY NOTE ERROR: ${error.message}`, error.stack);
      return { isError: true, content: [{ type: 'text', text: `❌ **Error creating daily note:** ${error.message}` }] };
    }
  }
);

server.tool(
  'searchInNotebook',
  {
    notebookId: z.string().describe('The ID of the notebook to search within.'),
    query: z.string().describe('Optional keyword to filter page titles.').optional(),
    days: z.number().min(1).describe('Optional: search pages from last N days.').optional(),
    top: z.number().min(1).max(100).default(100).describe('Max pages per section.').optional()
  },
  async ({ notebookId, query, days, top }) => {
    try {
      await ensureGraphClient();
      console.error(`Searching in notebook ID: ${notebookId}`);
      
      // Get notebook info
      const notebook = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
      const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebookId}/sections`);
      
      if (sections.length === 0) {
        return { content: [{ type: 'text', text: `📂 No sections found in notebook "${notebook.displayName}".` }] };
      }
      
      console.error(`Found ${sections.length} section(s), searching pages...`);
      let allMatchingPages = [];
      const threshold = days ? new Date(Date.now() - days * 24 * 60 * 60 * 1000) : null;
      
      for (const section of sections) {
        try {
          const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=${top}`;
          const pages = await paginateGraphRequest(apiPath);
          
          const matchingPages = pages.filter(page => {
            // Date filter if specified
            if (threshold) {
              const created = new Date(page.createdDateTime);
              const modified = new Date(page.lastModifiedDateTime);
              if (created < threshold && modified < threshold) {
                return false;
              }
            }
            
            // Keyword filter if specified
            if (query) {
              const searchTerm = query.toLowerCase();
              return page.title && page.title.toLowerCase().includes(searchTerm);
            }
            
            return true;
          });
          
          matchingPages.forEach(page => {
            page._section = section.displayName || section.name;
          });
          
          allMatchingPages = allMatchingPages.concat(matchingPages);
        } catch (sectionError) {
          console.error(`Error in section ${section.displayName}: ${sectionError.message}`);
        }
      }
      
      console.error(`Search complete: ${allMatchingPages.length} matches`);
      
      if (allMatchingPages.length === 0) {
        return { content: [{ type: 'text', text: `🔍 No matching pages found in notebook "${notebook.displayName}".` }] };
      }
      
      // Sort by most recent
      allMatchingPages.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      let resultText = `🔍 **Search Results in "${notebook.displayName}"** (${allMatchingPages.length} found):\n`;
      if (query) resultText += `Keyword: "${query}"\n`;
      if (days) resultText += `Time range: Last ${days} day(s)\n`;
      resultText += `Checked: ${sections.length} sections\n\n`;
      
      const displayPages = allMatchingPages.slice(0, 20);
      for (let i = 0; i < displayPages.length; i++) {
        const page = displayPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n\n`;
      }
      
      if (allMatchingPages.length > 20) {
        resultText += `... and ${allMatchingPages.length - 20} more pages.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search in notebook: ${error.message}` }] };
    }
  }
);

server.tool(
  'getPageContent',
  {
    pageId: z.string().describe('The ID of the page to retrieve content from.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text (readable), html (raw), or summary (brief).')
      .optional()
  },
  async ({ pageId, format }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
      let resultText = '';

      if (format === 'html') {
        resultText = `📄 **${pageInfo.title}** (HTML Format)\n\n${htmlContent}`;
      } else if (format === 'summary') {
        const summary = extractTextSummary(htmlContent, 300);
        resultText = `📄 **${pageInfo.title}** (Summary)\n\n${summary}`;
      } else { // 'text'
        const textContent = extractReadableText(htmlContent);
        resultText = `📄 **${pageInfo.title}**\n📅 Modified: ${new Date(pageInfo.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
      }
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get page content for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'getPageByTitle',
  {
    title: z.string().describe('The title (or partial title) of the page to find.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text, html, or summary.')
      .optional()
  },
  async ({ title, format }) => {
    try {
      await ensureGraphClient();
      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      const matchingPage = (pagesResponse.value || []).find(p => p.title && p.title.toLowerCase().includes(title.toLowerCase()));

      if (!matchingPage) {
        const availablePages = (pagesResponse.value || []).slice(0, 10).map(p => `- ${p.title}`).join('\n');
        return { isError: true, content: [{ type: 'text', text: `❌ No page found with title containing "${title}".\n\nAvailable pages (up to 10):\n${availablePages || 'None'}` }] };
      }

      const htmlContent = await fetchPageContentAdvanced(matchingPage.id, 'httpDirect');
      let resultText = '';
      if (format === 'html') {
        resultText = `📄 **${matchingPage.title}** (HTML Format)\n\n${htmlContent}`;
      } else if (format === 'summary') {
        const summary = extractTextSummary(htmlContent, 300);
        resultText = `📄 **${matchingPage.title}** (Summary)\n\n${summary}`;
      } else { // 'text'
        const textContent = extractReadableText(htmlContent);
        resultText = `📄 **${matchingPage.title}**\n📅 Modified: ${new Date(matchingPage.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
      }
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get page by title "${title}": ${error.message}` }] };
    }
  }
);

// --- Page Editing & Content Manipulation Tools ---

server.tool(
  'updatePageContent',
  {
    pageId: z.string().describe('The ID of the page to update.'),
    content: z.string().describe('New page content (HTML or markdown-style text).'),
    preserveTitle: z.boolean()
      .default(true)
      .describe('Keep the original title (default: true).')
      .optional()
  },
  async ({ pageId, content: newContent, preserveTitle }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Updating content for page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const htmlContentForUpdate = textToHtml(newContent);
      const finalHtml = `
        <div>
          ${preserveTitle ? `<h1>${pageInfo.title}</h1>` : ''}
          ${htmlContentForUpdate}
          <hr>
          <p><em>Updated via OneNote MCP on ${new Date().toLocaleString()}</em></p>
        </div>
      `;
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'replace', content: finalHtml }])
      });
      
      if (!response.ok) throw new Error(`Update failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Page Content Updated!**\nPage: ${pageInfo.title}\nUpdated: ${new Date().toLocaleString()}\nContent Length: ${newContent.length} chars.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to update page content for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'appendToPage',
  {
    pageId: z.string().describe('The ID of the page to append content to.'),
    content: z.string().describe('Content to append (HTML or markdown-style).'),
    addTimestamp: z.boolean().default(true).describe('Add a timestamp (default: true).').optional(),
    addSeparator: z.boolean().default(true).describe('Add a visual separator (default: true).').optional()
  },
  async ({ pageId, content: newContent, addTimestamp, addSeparator }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Appending content to page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const htmlContentToAppend = textToHtml(newContent);
      let appendHtml = '';
      if (addSeparator) appendHtml += '<hr>';
      if (addTimestamp) appendHtml += `<p><em>Added on ${new Date().toLocaleString()}</em></p>`;
      appendHtml += htmlContentToAppend;
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'append', content: appendHtml }])
      });
      
      if (!response.ok) throw new Error(`Append failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Content Appended!**\nPage: ${pageInfo.title}\nAppended: ${new Date().toLocaleString()}\nLength: ${newContent.length} chars.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to append content to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'updatePageTitle',
  {
    pageId: z.string().describe('The ID of the page whose title is to be updated.'),
    newTitle: z.string().describe('The new title for the page.')
  },
  async ({ pageId, newTitle }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const oldTitle = pageInfo.title;
      console.error(`Updating page title from "${oldTitle}" to "${newTitle}" for page ID "${pageId}"`);
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'title', action: 'replace', content: newTitle }])
      });
      
      if (!response.ok) throw new Error(`Title update failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Page Title Updated!**\nOld Title: ${oldTitle}\nNew Title: ${newTitle}\nUpdated: ${new Date().toLocaleString()}` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to update page title for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'replaceTextInPage',
  {
    pageId: z.string().describe('The ID of the page to modify.'),
    findText: z.string().describe('The text to find and replace.'),
    replaceText: z.string().describe('The text to replace with.'),
    caseSensitive: z.boolean().default(false).describe('Case-sensitive search (default: false).').optional()
  },
  async ({ pageId, findText, replaceText, caseSensitive }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
      console.error(`Replacing text in page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const flags = caseSensitive ? 'g' : 'gi';
      const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
      const matches = (htmlContent.match(regex) || []).length;
      
      if (matches === 0) {
        return { content: [{ type: 'text', text: `ℹ️ **No matches found** for "${findText}" in page: ${pageInfo.title}.` }] };
      }
      
      const updatedContent = htmlContent.replace(regex, replaceText);
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'replace', content: `<div>${updatedContent}</div>` }])
      });
      
      if (!response.ok) throw new Error(`Replace failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Text Replaced!**\nPage: ${pageInfo.title}\nFound: "${findText}" (${matches} occurrences)\nReplaced with: "${replaceText}".` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to replace text in page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'addNoteToPage',
  {
    pageId: z.string().describe('The ID of the page to add a note to.'),
    note: z.string().describe('The note/comment content.'),
    noteType: z.enum(['note', 'todo', 'important', 'question'])
      .default('note')
      .describe('Type of note (note, todo, important, question).')
      .optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the note (top or bottom).')
      .optional()
  },
  async ({ pageId, note, noteType, position }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Adding ${noteType} to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);
      
      const icons = { note: '📝', todo: '✅', important: '🚨', question: '❓' };
      const colors = { note: '#e3f2fd', todo: '#e8f5e8', important: '#ffebee', question: '#fff3e0' };
      const noteHtml = `
        <div style="border-left: 4px solid #2196f3; background-color: ${colors[noteType]}; padding: 10px; margin: 10px 0;">
          <p><strong>${icons[noteType]} ${noteType.charAt(0).toUpperCase() + noteType.slice(1)}</strong> - <em>${new Date().toLocaleString()}</em></p>
          <p>${textToHtml(note)}</p>
        </div>`;
      
      const action = position === 'top' ? 'prepend' : 'append';
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: action, content: noteHtml }])
      });
      
      if (!response.ok) throw new Error(`Add note failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **${noteType.charAt(0).toUpperCase() + noteType.slice(1)} Added!**\nPage: ${pageInfo.title}\nPosition: ${position}.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to add note to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'addTableToPage',
  {
    pageId: z.string().describe('The ID of the page to add a table to.'),
    tableData: z.string().describe('Table data in CSV format (header row, then data rows).'),
    title: z.string().describe('Optional title for the table.').optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the table (top or bottom).')
      .optional()
  },
  async ({ pageId, tableData, title, position }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Adding table to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);
      
      const rows = tableData.trim().split('\n').map(row => row.split(',').map(cell => cell.trim()));
      if (rows.length < 2) throw new Error('Table data must have at least a header row and one data row.');
      
      const headerRow = rows[0];
      const dataRows = rows.slice(1);
      let tableHtml = title ? `<h3>📊 ${textToHtml(title)}</h3>` : '';
      tableHtml += `<table style="border-collapse: collapse; width: 100%; margin: 10px 0;"><thead><tr style="background-color: #f5f5f5;">${headerRow.map(cell => `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${textToHtml(cell)}</th>`).join('')}</tr></thead><tbody>${dataRows.map(row => `<tr>${row.map(cell => `<td style="border: 1px solid #ddd; padding: 8px;">${textToHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
      
      const action = position === 'top' ? 'prepend' : 'append';
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: action, content: tableHtml }])
      });
      
      if (!response.ok) throw new Error(`Add table failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Table Added!**\nPage: ${pageInfo.title}\nTitle: ${title || 'Untitled'}\nPosition: ${position}.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to add table to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

// --- Page Creation Tool ---
server.tool(
  'createPage',
  {
    title: z.string().min(1, { message: "Title cannot be empty." }).describe('The title for the new page.'),
    content: z.string().min(1, { message: "Content cannot be empty." }).describe('The content for the new page (HTML or markdown-style).')
  },
  async ({ title, content }) => {
    try {
      await ensureGraphClient();
      console.error(`Attempting to create page with title: "${title}"`);
      
      const sectionsResponse = await graphClient.api('/me/onenote/sections').get();
      if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
        throw new Error('No sections found in your OneNote. Cannot create a page.');
      }
      const targetSectionId = sectionsResponse.value[0].id;
      const targetSectionName = sectionsResponse.value[0].displayName;
      
      const htmlContent = textToHtml(content);
      const pageHtml = `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(title)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(title)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${targetSectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);
      
      return {
        content: [{
          type: 'text',
          text: `✅ **Page Created Successfully!**
**Title:** ${response.title}
**Page ID:** ${response.id}
**In Section:** ${targetSectionName}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`
        }]
      };
    } catch (error) {
      console.error(`CREATE PAGE ERROR: ${error.message}`, error.stack);
      return { isError: true, content: [{ type: 'text', text: `❌ **Error creating page:** ${error.message}` }] };
    }
  }
);



// ============================================================================
// SERVER STARTUP
// ============================================================================

/**
 * Main function to initialize and start the MCP server.
 */
async function main() {
  loadExistingToken(); // Attempt to load token at startup
  if (accessToken) {
    initializeGraphClient(); // Initialize client if token was loaded
  }

  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('🚀✨ OneNote Ultimate MCP Server is now LIVE! ✨🚀');
    console.error(`   Client ID: ${clientId.substring(0, 8)}... (Using ${process.env.AZURE_CLIENT_ID ? 'environment variable' : 'default'})`);
    console.error('   Ready to manage your OneNote like never before!');
    console.error('--- Available Tool Categories ---');
    console.error('   🔐 Auth: authenticate, saveAccessToken');
    console.error('   📚 Read: listNotebooks, listSections, listPagesInSection, searchPages, searchPagesByDate, searchPageContent, searchInNotebook, getPageContent, getPageByTitle');
    console.error('   📝 Productivity: getMyRecentChanges, createDailyNote');
    console.error('   ✏️ Edit: updatePageContent, appendToPage, updatePageTitle, replaceTextInPage, addNoteToPage, addTableToPage');
    console.error('   ➕ Create: createPage');
    console.error('---------------------------------');
    
    process.on('SIGINT', () => {
      console.error('\n🔌 OneNote MCP Server shutting down gracefully...');
      process.exit(0);
    });
    process.on('SIGTERM', () => {
      console.error('\n🔌 OneNote MCP Server terminated...');
      process.exit(0);
    });

  } catch (error) {
    console.error(`💀 Critical error starting server: ${error.message}`, error.stack);
    process.exit(1);
  }
}

main();