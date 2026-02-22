# OneNote MCP Server - Design Document

## 1. Overview

This system is a **Model Context Protocol (MCP) server** that acts as a bridge between AI assistants and Microsoft OneNote. It exposes OneNote operations (read, write, search, edit) as discrete tools that an MCP-compatible client can invoke over a stdio transport.

### Core Responsibilities

- Authenticate users against Microsoft Identity Platform via device code flow
- Provide structured tool interfaces for OneNote CRUD operations
- Translate between user-friendly formats (markdown, plain text, CSV) and OneNote's HTML-based storage
- Handle API rate limiting, pagination, and error recovery transparently

---

## 2. Architecture

```
┌─────────────────────┐     stdio (JSON-RPC)     ┌──────────────────────┐
│                     │◄────────────────────────► │                      │
│   MCP Client        │                           │   OneNote MCP Server │
│   (AI Assistant)    │                           │                      │
│                     │                           │  ┌────────────────┐  │
└─────────────────────┘                           │  │  Tool Router   │  │
                                                  │  └───────┬────────┘  │
                                                  │          │           │
                                      ┌───────────┼──────────┼───────┐   │
                                      │           │          │       │   │
                                      ▼           ▼          ▼       │   │
                                 ┌─────────┐ ┌────────┐ ┌────────┐  │   │
                                 │  Auth   │ │  Read  │ │ Write  │  │   │
                                 │ Module  │ │ Module │ │ Module │  │   │
                                 └────┬────┘ └───┬────┘ └───┬────┘  │   │
                                      │          │          │       │   │
                                      ▼          ▼          ▼       │   │
                                 ┌──────────────────────────────┐   │   │
                                 │     Graph API Client Layer   │   │   │
                                 │  (retry, pagination, auth)   │   │   │
                                 └──────────────┬───────────────┘   │   │
                                                │                   │   │
                                 ┌──────────────┴───────────────┐   │   │
                                 │   Content Processing Layer   │   │   │
                                 │  (HTML↔Text, Markdown→HTML)  │   │   │
                                 └──────────────────────────────┘   │   │
                                                                    │   │
                                 ┌──────────────────────────────┐   │   │
                                 │   Token Store (filesystem)   │   │   │
                                 └──────────────────────────────┘   │   │
                                                                    │   │
                                                  ┌─────────────────┘   │
                                                  ▼                     │
                                     ┌──────────────────────┐          │
                                     │  Microsoft Graph API │          │
                                     │  (graph.microsoft.com)│          │
                                     └──────────────────────┘          │
                                                                       │
                                                  └────────────────────┘
```

### Layered Design

| Layer | Responsibility |
|---|---|
| **Transport** | Stdio-based JSON-RPC communication (MCP protocol) |
| **Tool Router** | Maps tool names to handler functions; validates input schemas |
| **Tool Handlers** | Business logic for each operation; grouped by domain |
| **Graph API Client** | HTTP communication with Microsoft Graph; manages auth headers, retries, pagination |
| **Content Processing** | Bidirectional format conversion (HTML, Markdown, Plain Text, CSV) |
| **Token Store** | Persistent credential storage with expiry management |

---

## 3. Component Design

### 3.1 Authentication Module

**Purpose:** Manage the OAuth 2.0 device code flow and token lifecycle.

**Behavior:**

1. **Device Code Flow:** Initiate OAuth via device code grant. Present user code and verification URL to the client. Poll for token completion in the background.
2. **Token Persistence:** Store the token as JSON with metadata (client ID, scopes, creation timestamp, expiry timestamp) on the local filesystem.
3. **Token Loading:** On startup, attempt to load a persisted token. Validate expiry before use. Delete and prompt re-authentication if expired.
4. **Lazy Initialization:** The Graph API client is only initialized when a token becomes available, not at startup.

**Security Requirements:**

- Token file MUST be written with restrictive permissions (owner read/write only)
- Token expiry MUST be validated before use
- Client ID SHOULD be configurable via environment variable; warn if using a shared/default ID

**Data Model - Stored Token:**

```
TokenData {
  token:     string     // The OAuth bearer token
  clientId:  string     // Azure App Registration ID used
  scopes:    string[]   // Granted permission scopes
  createdAt: datetime   // When the token was obtained
  expiresOn: datetime?  // When the token expires (null if unknown)
}
```

**Tools:**

| Tool | Input | Output |
|---|---|---|
| `authenticate` | _(none)_ | Device code instructions (URL + user code) |
| `saveAccessToken` | _(none)_ | Verification result with account display name and email |

---

### 3.2 Read Module

**Purpose:** Query and retrieve OneNote data without modification.

**OneNote Data Hierarchy:**

```
User Account
  └── Notebook (id, displayName)
        └── Section (id, displayName)
              └── Page (id, title, createdDateTime, lastModifiedDateTime, links, content)
```

**Tools:**

| Tool | Input | Output | Notes |
|---|---|---|---|
| `listNotebooks` | _(none)_ | List of notebooks with IDs and metadata | |
| `listSections` | `notebookId` | List of sections in the notebook | |
| `listPagesInSection` | `sectionId`, `top?`, `orderBy?` | Paginated page list | Supports sorting by created or modified date |
| `searchPages` | `query?` | Pages matching title substring | Fetches all pages then filters client-side |
| `searchPagesByDate` | `days?`, `query?`, `dateField?`, `includeContent?`, `createdBy?`, `modifiedBy?` | Filtered page list | Iterates all notebooks/sections; supports user attribution filter |
| `searchPageContent` | `query`, `days?`, `notebookId?`, `maxPages?` | Pages where body text contains query | Fetches and parses HTML content of each page |
| `searchInNotebook` | `notebookId`, `query?`, `days?`, `top?` | Filtered pages within one notebook | Scoped search for performance |
| `getPageContent` | `pageId`, `format?` | Page content in requested format | Formats: `text`, `html`, `summary` |
| `getPageByTitle` | `title`, `format?` | Content of first page matching title | Partial title match (case-insensitive) |
| `getMyRecentChanges` | `sinceDate?`, `days?`, `notebookId?`, `includeCreator?` | Pages modified by current user | Compares modifier email with authenticated user; groups results by notebook |

**Design Decisions:**

- **Client-side filtering for title search:** The OneNote Graph API does not support `$filter` on page titles, so all pages must be fetched and filtered in-memory.
- **Content search is expensive:** `searchPageContent` must fetch the HTML body of each candidate page. A `maxPages` cap prevents runaway API usage.
- **Date threshold calculation:** Supports absolute dates (`YYYY-MM-DD`), relative keywords (`"monday"`), and day counts.

---

### 3.3 Write Module

**Purpose:** Create and modify OneNote pages and their content.

**Tools:**

| Tool | Input | Output | Notes |
|---|---|---|---|
| `createPage` | `title`, `content` | Created page metadata | Targets the first available section |
| `createDailyNote` | `notebookName`, `sectionName`, `date?`, `title?`, `content?` | Created page metadata | Auto-formats date as `M/D/YY`; duplicate detection |
| `updatePageContent` | `pageId`, `content`, `preserveTitle?` | Confirmation | Replaces entire page body |
| `appendToPage` | `pageId`, `content`, `addTimestamp?`, `addSeparator?` | Confirmation | Appends to end of body |
| `updatePageTitle` | `pageId`, `newTitle` | Confirmation | PATCH title target |
| `replaceTextInPage` | `pageId`, `findText`, `replaceText`, `caseSensitive?` | Match count + confirmation | Regex-based find/replace on raw HTML; escapes special regex chars from user input |
| `addNoteToPage` | `pageId`, `note`, `noteType?`, `position?` | Confirmation | Styled callout block; types: note, todo, important, question |
| `addTableToPage` | `pageId`, `tableData`, `title?`, `position?` | Confirmation | Parses CSV input into HTML table |

**OneNote PATCH Operations:**

The OneNote API uses a specific PATCH format for page modifications:

```
PATCH /me/onenote/pages/{id}/content
Content-Type: application/json

[
  {
    "target": "body" | "title",
    "action": "replace" | "append" | "prepend",
    "content": "<html content>"
  }
]
```

**Security Requirements:**

- Page titles MUST be HTML-escaped (not markdown-converted) before insertion into `<title>` and `<h1>` tags
- Table cell values and table titles MUST be HTML-escaped
- Markdown link conversion MUST validate URL protocols (allow only `http`, `https`, `mailto`)
- HTML passthrough MUST strip `<script>` tags and inline event handler attributes

---

### 3.4 Content Processing Layer

**Purpose:** Convert between the formats users provide and the HTML that OneNote stores.

#### 3.4.1 Markdown/Text to HTML (`textToHtml`)

**Input:** Plain text, optionally with markdown formatting
**Output:** HTML suitable for OneNote page content

**Supported Markdown Features:**

| Markdown Syntax | HTML Output |
|---|---|
| `# Heading` (1-3 levels) | `<h1>` through `<h3>` |
| `**bold**` or `__bold__` | `<strong>` |
| `*italic*` or `_italic_` | `<em>` |
| `` `code` `` | `<code>` |
| ` ```block``` ` | `<pre><code>` |
| `[text](url)` | `<a href="url">` (protocol-validated) |
| `---` | `<hr>` |
| `> quote` | `<blockquote>` |
| `- item` or `* item` | `<ul><li>` |
| `1. item` | `<li>` (rendered as unordered in practice) |

**Behavioral Rules:**

1. If input already contains `<html>` or `<!DOCTYPE html>`, treat as pre-existing HTML. Sanitize by removing `<script>` tags and `on*` event attributes, then pass through.
2. Otherwise, HTML-escape the raw input first (`&`, `<`, `>`), then apply markdown transformations.
3. Wrap any remaining bare lines in `<p>` tags.
4. Consolidate adjacent `<li>` elements into `<ul>` wrappers.

#### 3.4.2 HTML to Readable Text (`extractReadableText`)

**Input:** Raw HTML from OneNote page content
**Output:** Formatted plain text

**Processing Steps:**

1. Remove `<script>` and `<style>` elements
2. Extract headings as underlined text blocks
3. Extract paragraphs as double-newline-separated text
4. Convert lists (`<ul>`, `<ol>`) to `- item` or `1. item` format
5. Convert tables to pipe-separated rows
6. Fallback: if no structured content found, use `body.textContent`

#### 3.4.3 HTML to Summary (`extractTextSummary`)

**Input:** Raw HTML + max length (default 300 chars)
**Output:** Truncated plain text with `...` suffix if truncated

#### 3.4.4 HTML Escaping (`escapeHtml`)

**Input:** Arbitrary string
**Output:** String safe for insertion into HTML text nodes and attributes

**Required Escape Mappings:**

| Character | Escape |
|---|---|
| `&` | `&amp;` |
| `<` | `&lt;` |
| `>` | `&gt;` |
| `"` | `&quot;` |
| `'` | `&#39;` |

**Usage:** Mandatory for page titles, table cell values, and any value inserted into HTML attributes.

---

### 3.5 Graph API Client Layer

**Purpose:** Manage all HTTP communication with the Microsoft Graph API.

#### 3.5.1 Client Initialization

- Lazily initialized when a valid token is available
- Auth provider callback supplies the bearer token on each request

#### 3.5.2 Retry with Exponential Backoff (`retryWithBackoff`)

**Trigger:** HTTP 429 (Too Many Requests) or throttling error messages
**Strategy:**

| Attempt | Wait Before Retry |
|---|---|
| 1st retry | 1 second |
| 2nd retry | 2 seconds |
| 3rd retry | 4 seconds |

Formula: `wait = 2^attempt * 1000ms`
Default max retries: 3

Non-429 errors are thrown immediately without retry.

#### 3.5.3 Pagination (`paginateGraphRequest`)

Microsoft Graph returns large result sets across multiple pages using `@odata.nextLink`.

**Algorithm:**

```
results = []
currentPath = initialApiPath
while currentPath exists AND (maxPages is null OR pageCount < maxPages):
    response = GET currentPath (with retry)
    append response.value to results
    currentPath = response["@odata.nextLink"]
    log progress every 10 pages
return results
```

**Parameters:**
- `apiPath`: Initial Graph API endpoint (e.g., `/me/onenote/pages`)
- `maxPages`: Optional cap on number of pages to fetch (null = unlimited)

#### 3.5.4 Page Content Fetching (`fetchPageContentAdvanced`)

Two methods supported:

| Method | How | When to Use |
|---|---|---|
| `httpDirect` (default) | Raw HTTP GET with bearer token to `graph.microsoft.com/v1.0/me/onenote/pages/{id}/content` | Primary method; returns raw HTML string |
| `direct` | Graph SDK client `.api().get()` | Fallback |

---

### 3.6 Transport Layer

**Protocol:** Model Context Protocol (MCP) over stdio
**Encoding:** JSON-RPC 2.0 over stdin/stdout
**Logging:** All diagnostic output goes to stderr (not stdout, which is reserved for MCP protocol messages)

**Lifecycle:**

1. Initialize MCP server with metadata (name, version, description)
2. Register all tool definitions with input schemas
3. Attempt to load persisted token
4. Create stdio transport and connect
5. Handle SIGINT/SIGTERM for graceful shutdown

---

## 4. Input Validation

All tool inputs are validated using schema definitions before the handler executes.

**Schema Types Used:**

| Type | Constraints |
|---|---|
| `string` | Required or optional; `.min(1)` where empty is invalid |
| `number` | `.min(1)`, `.max(100)` for pagination limits |
| `enum` | Restricted value sets (e.g., `"created" \| "modified"`, `"text" \| "html" \| "summary"`) |
| `boolean` | With defaults (e.g., `preserveTitle` defaults to `true`) |

**Validation must happen at the tool router level**, before any business logic executes.

---

## 5. Error Handling Strategy

| Error Type | Behavior |
|---|---|
| No token available | Throw with message directing user to authenticate |
| Token expired | Delete token file; prompt re-authentication |
| HTTP 429 (rate limit) | Retry with exponential backoff (up to 3 times) |
| Graph API error | Return `isError: true` with error message to client |
| Section-level failure during bulk operations | Log to stderr; continue with next section |
| Content parsing failure | Return fallback text; log error to stderr |

**Principle:** Bulk operations (date search, content search) should be fault-tolerant at the section/page level. A single failed section should not abort the entire search.

---

## 6. Security Considerations

### 6.1 Token Management

- Store tokens with owner-only file permissions (e.g., mode `0600` on Unix, ACL-restricted on Windows)
- Validate token expiry on every load; do not use expired tokens
- Delete expired token files automatically
- Never log or return the token value in tool responses

### 6.2 Input Sanitization

- **Page titles** in HTML templates: Use strict HTML escaping only (not markdown conversion)
- **Markdown links**: Validate URL protocol against allowlist (`http:`, `https:`, `mailto:`)
- **HTML passthrough**: Strip `<script>` elements and `on*` event handler attributes
- **Table cells**: HTML-escape all values
- **Regex construction from user input**: Escape all special regex characters before creating RegExp objects

### 6.3 OAuth Configuration

- Provide the client ID via environment variable; never hardcode production credentials
- Warn at startup if using a shared/default client ID
- Request the minimum necessary OAuth scopes for the operations being performed

### 6.4 Error Messages

- Error messages returned to the MCP client should describe the failure clearly but avoid leaking internal resource paths, token fragments, or stack traces
- Full stack traces should only go to stderr (server-side logging)

---

## 7. Configuration

| Setting | Source | Default | Description |
|---|---|---|---|
| Azure Client ID | Environment variable (`AZURE_CLIENT_ID`) | Microsoft Graph Explorer app ID | The Azure AD app registration to authenticate as |
| OAuth Scopes | Hardcoded | `Notes.Read`, `Notes.ReadWrite`, `Notes.Create`, `User.Read` | Permissions requested during auth |
| Token File Path | Derived | `.access-token.txt` in server directory | Where the token is persisted |
| Max Retries | Hardcoded | 3 | Retry count for rate-limited requests |

---

## 8. Data Flow Examples

### 8.1 Search Pages by Date

```
Client calls "searchPagesByDate" { days: 3, query: "standup" }
    │
    ▼
Validate input schema
    │
    ▼
Ensure authenticated (load token → init Graph client)
    │
    ▼
Calculate date threshold (now - 3 days)
    │
    ▼
GET /me/onenote/notebooks → list of notebooks
    │
    ▼
For each notebook:
    GET /me/onenote/notebooks/{id}/sections → list of sections
        │
        For each section:
            GET /me/onenote/sections/{id}/pages (with pagination)
                │
                Filter: modified/created >= threshold
                Filter: title contains "standup" (case-insensitive)
                │
                Attach notebook + section context to each match
    │
    ▼
Sort all matches by lastModifiedDateTime descending
    │
    ▼
Format top 20 results as structured text
    │
    ▼
Return { content: [{ type: "text", text: "..." }] }
```

### 8.2 Create Daily Note

```
Client calls "createDailyNote" { notebookName: "Work", sectionName: "Daily", date: "today" }
    │
    ▼
Validate input → resolve date → format title as "M/D/YY"
    │
    ▼
Find notebook by name (case-insensitive match)
    │ (error if not found → return available notebooks)
    ▼
Find section by name within that notebook
    │ (error if not found → return available sections)
    ▼
Check for existing page with same title
    │ (if exists → return info about existing page, do not duplicate)
    ▼
Convert content to HTML (textToHtml) or use default placeholder
Escape title for HTML insertion (escapeHtml)
    │
    ▼
Build full HTML document with title, content, and creation timestamp
    │
    ▼
POST /me/onenote/sections/{id}/pages (Content-Type: application/xhtml+xml)
    │
    ▼
Return confirmation with page ID, title, and web URL
```

---

## 9. API Surface Summary

### Authentication (2 tools)

| Tool | Purpose |
|---|---|
| `authenticate` | Start device code OAuth flow |
| `saveAccessToken` | Verify and load stored token |

### Read Operations (9 tools)

| Tool | Purpose |
|---|---|
| `listNotebooks` | Enumerate all notebooks |
| `listSections` | Enumerate sections in a notebook |
| `listPagesInSection` | Paginated page listing |
| `searchPages` | Title-based search across all notebooks |
| `searchPagesByDate` | Date-range + keyword + author filter |
| `searchPageContent` | Full-text body search |
| `searchInNotebook` | Scoped search within one notebook |
| `getPageContent` | Retrieve page in text/html/summary format |
| `getPageByTitle` | Find and retrieve page by title match |

### Productivity (2 tools)

| Tool | Purpose |
|---|---|
| `getMyRecentChanges` | Current user's recent modifications |
| `createDailyNote` | Auto-dated daily note with duplicate check |

### Write Operations (6 tools)

| Tool | Purpose |
|---|---|
| `createPage` | Create page in first available section |
| `updatePageContent` | Replace page body |
| `appendToPage` | Append content with optional timestamp |
| `updatePageTitle` | Change page title |
| `replaceTextInPage` | Find-and-replace in page HTML |
| `addNoteToPage` | Add styled callout block |
| `addTableToPage` | Insert table from CSV data |

**Total: 19 tools**

---

## 10. Implementation Notes for Alternative Languages

This design is language-agnostic. Key considerations when porting:

- **MCP SDK:** An MCP server SDK or library is needed for the target language. The protocol is JSON-RPC 2.0 over stdio. If no SDK exists, implement the protocol directly.
- **Microsoft Graph Client:** Use the official Microsoft Graph SDK for your language, or make raw HTTP calls to `https://graph.microsoft.com/v1.0/`.
- **HTML Parsing:** For `extractReadableText`, you need a DOM parser (e.g., `HtmlAgilityPack` in C#, `BeautifulSoup` in Python, `Jsoup` in Java).
- **Schema Validation:** The input schemas can be implemented with any validation framework (e.g., `FluentValidation` in C#, `pydantic` in Python).
- **Token Storage:** Consider using OS-native credential storage (e.g., Windows Credential Manager, macOS Keychain, `libsecret` on Linux) instead of filesystem storage for stronger security.
- **Async Patterns:** All Graph API calls should be async. Use the language's native async patterns (`async/await` in C#, `asyncio` in Python, etc.).
