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
 * @returns {string} Formatted page information string.
 */
function formatPageInfo(page, index = null) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  return `${prefix}**${page.title}**
   ID: ${page.id}
   Created: ${new Date(page.createdDateTime).toLocaleDateString()}
   Modified: ${new Date(page.lastModifiedDateTime).toLocaleDateString()}`;
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
  'searchPages',
  {
    query: z.string().describe('The search term for page titles.').optional()
  },
  async ({ query }) => {
    try {
      await ensureGraphClient();
      const apiResponse = await graphClient.api('/me/onenote/pages').get();
      let pages = apiResponse.value || [];
      if (query) {
        const searchTerm = query.toLowerCase();
        pages = pages.filter(page => page.title && page.title.toLowerCase().includes(searchTerm));
      }
      if (pages.length > 0) {
        const pageList = pages.slice(0, 10).map((page, i) => formatPageInfo(page, i)).join('\n\n');
        const morePages = pages.length > 10 ? `\n\n... and ${pages.length - 10} more pages.` : '';
        return { content: [{ type: 'text', text: `🔍 **Search Results** ${query ? `for "${query}"` : ''} (${pages.length} found):\n\n${pageList}${morePages}` }] };
      } else {
        return { content: [{ type: 'text', text: query ? `🔍 No pages found matching "${query}".` : '📄 No pages found.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search pages: ${error.message}` }] };
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
    console.error('   📚 Read: listNotebooks, searchPages, getPageContent, getPageByTitle');
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