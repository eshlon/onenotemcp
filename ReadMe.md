# OneNote MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

The OneNote MCP Server is a powerful Model Context Protocol (MCP) server that enables AI language models (LLMs) like Claude, and other AI assistants, to securely interact with your Microsoft OneNote data. It allows for reading, writing, searching, and comprehensive editing of your OneNote notebooks, sections, and pages directly through your AI interface.

This server provides a rich set of tools for advanced OneNote management, including robust text extraction, HTML content processing, and fine-grained page manipulation.

## Features

*   **Authentication:** Secure device code flow for Microsoft Graph API access.
*   **Read Operations:**
    *   List notebooks, sections, and pages with full pagination support.
    *   Search pages by title across all notebooks.
    *   **Advanced date-based searching** across all notebooks and sections.
    *   Filter pages by creation or modification date (e.g., find all notes from today).
    *   **Content search** - search within page content, not just titles.
    *   **Track modifications by user** - see who created or last modified pages.
    *   Scoped searches within specific notebooks for better performance.
    *   Get page content in various formats (full HTML, readable text, summary).
*   **Productivity Tools:**
    *   **Standup prep helper** - quickly see what pages YOU modified since last standup.
    *   **Quick daily note creator** - create dated notes with one command.
    *   Automatic date formatting (M/D/YY format).
    *   Smart duplicate detection (won't recreate existing daily notes).
*   **Pagination & Performance:**
    *   Handles accounts with 10+ notebooks and 100+ sections efficiently.
    *   Automatic retry logic with exponential backoff for rate limit handling.
    *   Smart pagination using @odata.nextLink for large result sets.
    *   Progress logging during long-running searches.
*   **Write & Edit Operations:**
    *   Create new OneNote pages with custom HTML or markdown content.
    *   Update entire page content, preserving or replacing the title.
    *   Append content to existing pages with optional timestamps and separators.
    *   Update page titles.
    *   Find and replace text within pages (case-sensitive or insensitive).
    *   Add formatted notes (like callouts or todos) to pages.
    *   Insert structured tables into pages from CSV data.
*   **Advanced Content Processing:**
    *   Sophisticated HTML to readable text extraction.
    *   Markdown-to-HTML conversion for page content.
*   **Robust Input Validation:** Uses Zod for defining and validating tool input schemas.


## Prerequisites

*   **Node.js:** Version 18.x or later is recommended. (Install from [nodejs.org](https://nodejs.org/))
*   **npm:** Usually comes bundled with Node.js.
*   **Git:** For cloning the repository. (Install from [git-scm.com](https://git-scm.com/))
*   **Microsoft Account:** An active Microsoft account with access to OneNote.
*   **Azure Application Registration (Recommended for Production/Shared Use):**
    *   While the server defaults to using the Microsoft Graph Explorer's public Client ID for easy testing, for regular or shared use, it is **strongly recommended** to create your own Azure App Registration.
    *   Ensure your app registration has the following delegated Microsoft Graph API permissions: `Notes.Read`, `Notes.ReadWrite`, `Notes.Create`, `User.Read`.
    *   You will need the "Application (client) ID" from your app registration.

## Installation

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/[your-github-username]/onenote-ultimate-mcp-server.git
    cd onenote-ultimate-mcp-server
    ```
    *(Replace `[your-github-username]/onenote-ultimate-mcp-server` with your actual repository URL)*

2.  **Install Dependencies:**
    ```bash
    npm install
    ```

## Configuration

1.  **Azure Client ID:**
    This server requires an Azure Application Client ID to authenticate with Microsoft Graph.
    *   **Recommended for Production/Shared Use:** Set the `AZURE_CLIENT_ID` environment variable to your own Azure App's "Application (client) ID".
        ```bash
        export AZURE_CLIENT_ID="your-actual-azure-app-client-id" 
        ```
        (On Windows, use `set AZURE_CLIENT_ID=your-actual-azure-app-client-id`)
    *   **For Quick Testing:** If the `AZURE_CLIENT_ID` environment variable is not set, the server will default to using the Microsoft Graph Explorer's public Client ID. This is suitable for initial testing but not recommended for prolonged or shared use.
    *   Alternatively, you can modify the `clientId` variable directly in `onenote-mcp.mjs`, but using an environment variable is preferred.

2.  **`.gitignore`:**
    The project includes a `.gitignore` file. Ensure it contains at least the following to prevent committing sensitive files:
    ```
    node_modules/
    .DS_Store
    *.log
    .access-token.txt
    .env
    ```
    The `.access-token.txt` file will be created by the server to store your authentication token.

## Running the MCP Server

Once configured, start the server from the project's root directory:

```bash
node onenote-mcp.mjs
```

You should see console output indicating the server has started and listing the available tool categories.

## Connecting to an MCP Client

You can connect this server to any MCP-compatible client, such as Claude Desktop or Cursor.

**Example for Claude Desktop or Cursor:**

1.  Open your MCP client's configuration file.
    *   **Claude Desktop (macOS):** `~/Library/Application Support/Claude/claude_desktop_config.json`
    *   **Claude Desktop (Windows):** `%APPDATA%\Claude\claude_desktop_config.json`
    *   **Cursor:** Preferences -> MCP tab.

2.  Add or update the `mcpServers` configuration:

    ```json
    {
      "mcpServers": {
        "onenote": {
          "command": "node",
          "args": ["/full/path/to/your/onenote-ultimate-mcp-server/onenote-mcp.mjs"],
          "env": {
            // Recommended: Set AZURE_CLIENT_ID here if not set globally
            "AZURE_CLIENT_ID": "YOUR_AZURE_APP_CLIENT_ID_HERE" 
          }
        }
      }
    }
    ```

    *   Replace `/full/path/to/your/onenote-ultimate-mcp-server/` with the **absolute path** to where you cloned the repository.
    *   Replace `YOUR_AZURE_APP_CLIENT_ID_HERE` with your Azure App's Client ID, especially if you are not setting it as a system-wide environment variable.

3.  Restart your MCP client (Claude Desktop/Cursor).

## Authentication Flow

The first time you try to use a OneNote tool through your AI assistant, or by explicitly invoking the `authenticate` tool:

1.  **Invoke `authenticate` Tool:** Your AI assistant will call the `authenticate` tool on the server.
2.  **Device Code Prompt:** The server will output a URL (typically `https://microsoft.com/devicelogin`) and a user code to its `stderr`. Your MCP client (e.g., Claude Desktop) should display this information to you.
3.  **Browser Authentication:** Open the provided URL in a web browser and enter the user code.
4.  **Sign In & Grant Permissions:** Sign in with your Microsoft account that has OneNote access and grant the requested permissions.
5.  **Token Saved:** Upon successful browser authentication, the server will automatically receive and save the access token to an `.access-token.txt` file in its directory.
6.  **Verify (Optional but Recommended):** Invoke the `saveAccessToken` tool through your AI assistant. This tool doesn't actually save (as it's already saved by the background process) but rather loads and verifies the token, confirming successful authentication and displaying your account info.

The saved token will be used for subsequent sessions until it expires, at which point you may need to re-authenticate.

## Available MCP Tools

This server exposes the following tools to your AI assistant. All tools require authentication unless otherwise noted.

---

### Authentication Tools

#### `authenticate`

Starts the OAuth 2.0 device code flow. The server generates a one-time user code and provides a URL (`https://microsoft.com/devicelogin`). After you complete sign-in in your browser, the token is saved automatically in the background. No input parameters required.

**Behavior:** The tool returns immediately with the device code instructions. Token retrieval continues asynchronously. If the browser sign-in succeeds, the token is persisted to `.access-token.txt` and the Graph client is initialized without further action.

#### `saveAccessToken`

Loads the persisted token from disk, verifies it against the Graph API, and displays the authenticated account name and email. Use this after `authenticate` to confirm the token was saved correctly. No input parameters required.

**Behavior:** Calls the `/me` Graph endpoint to verify the token. If no token file exists, returns an error directing the user to run `authenticate` first.

---

### Reading Tools

#### `listNotebooks`

Returns all OneNote notebooks accessible to the authenticated user, including each notebook's ID, display name, creation date, and modification date.

**Behavior:** Single API call to `/me/onenote/notebooks`. No pagination needed since users typically have a small number of notebooks.

#### `listSections`

Lists all sections within a specific notebook.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `notebookId` | string | Yes | The ID of the notebook (obtained from `listNotebooks`) |

**Behavior:** Uses pagination to retrieve all sections, so it handles notebooks with many sections.

#### `listPagesInSection`

Lists pages within a section, sorted by date, with pagination support.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `sectionId` | string | Yes | | The section ID (from `listSections`) |
| `top` | number | No | 20 | Pages per request (1-100) |
| `orderBy` | enum | No | `"modified"` | Sort by `"created"` or `"modified"` date (descending) |

**Behavior:** Fetches pages using the Graph API's `$top` and `$orderby` parameters. Displays up to 50 pages in the response; if more exist, a count of remaining pages is shown.

#### `searchPages`

Searches for pages by title across all notebooks.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `query` | string | No | Substring to match against page titles (case-insensitive). If omitted, returns all pages. |

**Behavior:** Fetches *all* pages via pagination, then filters client-side by title. This is necessary because the OneNote Graph API does not support `$filter` on page titles. Returns the first 10 matches. For large accounts, consider using `searchInNotebook` or `searchPagesByDate` instead to reduce API calls.

**Tip:** If you have many pages and searches feel slow, use `searchInNotebook` to scope the search to a single notebook, or `searchPagesByDate` to narrow by date range.

#### `searchPagesByDate`

Searches across all notebooks and sections for pages created or modified within a time window. Supports keyword filtering and creator/modifier attribution filtering.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `days` | number | No | 1 | How many days back to search (1 = today only) |
| `query` | string | No | | Keyword filter on page titles (case-insensitive) |
| `dateField` | enum | No | `"both"` | Which date to check: `"created"`, `"modified"`, or `"both"` |
| `includeContent` | boolean | No | `false` | Include a 200-character content preview for each match (slower, requires fetching page body) |
| `createdBy` | string | No | | Filter by creator name or email (partial, case-insensitive) |
| `modifiedBy` | string | No | | Filter by last modifier name or email (partial, case-insensitive) |

**Behavior:** Iterates through every notebook and section, fetching pages with `$expand=lastModifiedBy,createdBy`. Filters by date threshold, then optionally by keyword and user attribution. Results are sorted by most recently modified first. Returns up to 20 matches.

**Performance note:** This tool makes multiple API calls (one per section). For accounts with many notebooks/sections, it may take several seconds. Use the `notebookId`-scoped `searchInNotebook` tool when possible.

#### `searchPageContent`

Searches for text within the actual body content of pages, not just titles.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `query` | string | Yes | | Text to search for (case-insensitive) |
| `days` | number | No | | Limit to pages modified within the last N days |
| `notebookId` | string | No | | Limit to a specific notebook |
| `maxPages` | number | No | 20 | Maximum number of matching pages to return (1-50) |

**Behavior:** Iterates through notebooks and sections, fetching the HTML content of each page and converting it to plain text for searching. Returns matching pages with a contextual snippet showing the search term in context.

**Performance note:** This is the most expensive search tool because it must download and parse the HTML body of every candidate page. Use the `days`, `notebookId`, and `maxPages` parameters to limit the scope. The tool stops early once `maxPages` matches are found.

#### `searchInNotebook`

Scoped search within a single notebook. Faster than cross-notebook searches.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `notebookId` | string | Yes | | The notebook to search within |
| `query` | string | No | | Keyword filter on page titles (case-insensitive) |
| `days` | number | No | | Limit to pages from the last N days |
| `top` | number | No | 100 | Max pages to fetch per section (1-100) |

**Behavior:** Fetches all sections in the specified notebook, then pages within each section. Applies optional date and keyword filters. Results sorted by most recently modified. Returns up to 20 matches.

#### `getPageContent`

Retrieves the full content of a single page by its ID.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page ID (from any search or list tool) |
| `format` | enum | No | `"text"` | Output format: `"text"` (readable plain text), `"html"` (raw HTML), or `"summary"` (first 300 characters) |

**Behavior:** Fetches the page's HTML content via direct HTTP request, then converts it to the requested format. The `text` format strips HTML tags and formats headings, lists, and tables as readable plain text. The `html` format returns the raw OneNote HTML. The `summary` format extracts the first 300 characters of body text.

#### `getPageByTitle`

Finds a page by title and retrieves its content in one step.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `title` | string | Yes | | Full or partial page title (case-insensitive) |
| `format` | enum | No | `"text"` | Output format: `"text"`, `"html"`, or `"summary"` |

**Behavior:** Fetches the first page of results from `/me/onenote/pages`, finds the first page whose title contains the search string (case-insensitive), then fetches and returns its content. If no match is found, shows up to 10 available page titles to help the user refine their search.

**Limitation:** Only searches the first page of API results (default page size). For more thorough title searching, use `searchPages` first to find the page ID, then `getPageContent`.

---

### Productivity Tools

#### `getMyRecentChanges`

Shows pages that the currently authenticated user has modified since a given date. Designed for standup preparation.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `sinceDate` | string | No | | Start date in `YYYY-MM-DD` format, or `"monday"` to use last Monday |
| `days` | number | No | 3 | Alternative: number of days back to search |
| `notebookId` | string | No | | Limit to a specific notebook |
| `includeCreator` | boolean | No | `false` | Show who originally created each page |

**Behavior:** Fetches the current user's email via `/me`, then iterates through all notebooks/sections looking for pages where the `lastModifiedBy` email matches. Results are grouped by notebook for easy scanning. If both `sinceDate` and `days` are provided, `sinceDate` takes precedence.

**Tip:** Use `sinceDate: "monday"` at your Thursday standup to see everything you touched that week.

#### `createDailyNote`

Creates a new dated note in a specific notebook and section. Automatically formats the date and checks for duplicates.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `notebookName` | string | Yes | | Notebook name (case-insensitive match) |
| `sectionName` | string | Yes | | Section name within that notebook (case-insensitive match) |
| `date` | string | No | `"today"` | Date in `YYYY-MM-DD` format, or `"today"` |
| `title` | string | No | | Custom title (defaults to `M/D/YY` format, e.g., `"2/20/26"`) |
| `content` | string | No | | Initial page content (markdown or HTML). Defaults to placeholder text. |

**Behavior:** Looks up the notebook and section by name (case-insensitive). If the notebook or section isn't found, returns a list of available options. Checks if a page with the same title already exists in the section; if so, returns info about the existing page instead of creating a duplicate. The page is created with the content converted from markdown to HTML.

---

### Editing & Creating Tools

#### `createPage`

Creates a new page in the first available section across all notebooks.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `title` | string | Yes | Page title (cannot be empty) |
| `content` | string | Yes | Page content in markdown or HTML (cannot be empty) |

**Behavior:** Fetches all sections via `/me/onenote/sections` and uses the first one found. Content is converted from markdown to HTML. The title is HTML-escaped for safe insertion. Returns the created page's ID, title, section name, and creation timestamp.

**Note:** If you need to target a specific notebook and section, use `createDailyNote` instead (it accepts notebook/section names).

#### `updatePageContent`

Replaces the entire body content of an existing page.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page to update |
| `content` | string | Yes | | New content (markdown or HTML) |
| `preserveTitle` | boolean | No | `true` | Keep the original title as an `<h1>` at the top |

**Behavior:** Fetches the current page metadata, converts the new content to HTML, and sends a PATCH request with `action: "replace"` targeting the body. A timestamp footer is appended. This is a destructive operation that replaces all existing body content.

#### `appendToPage`

Adds content to the end of an existing page without affecting existing content.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page to append to |
| `content` | string | Yes | | Content to add (markdown or HTML) |
| `addTimestamp` | boolean | No | `true` | Prepend a timestamp above the new content |
| `addSeparator` | boolean | No | `true` | Add a horizontal rule before the new content |

**Behavior:** Converts content to HTML and sends a PATCH request with `action: "append"` targeting the body. The separator and timestamp appear above the new content to visually distinguish additions.

#### `updatePageTitle`

Changes the title of an existing page.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `pageId` | string | Yes | The page to rename |
| `newTitle` | string | Yes | The new title |

**Behavior:** Sends a PATCH request with `action: "replace"` targeting the `title`. The old title is displayed in the confirmation message.

#### `replaceTextInPage`

Finds and replaces text within a page's HTML content.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page to modify |
| `findText` | string | Yes | | Text to search for |
| `replaceText` | string | Yes | | Replacement text |
| `caseSensitive` | boolean | No | `false` | Whether to match case exactly |

**Behavior:** Fetches the page's raw HTML content, builds a regex from the search text (special regex characters are escaped automatically), and performs a global replace. Reports the number of occurrences found and replaced. If no matches are found, no changes are made.

**Important:** The find-and-replace operates on the raw HTML. This means it can match text inside HTML tags or attributes. For most plain-text replacements this works as expected, but be cautious with strings that might appear in HTML markup.

#### `addNoteToPage`

Adds a formatted, color-coded callout block to a page.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page to add the note to |
| `note` | string | Yes | | The note content (markdown or HTML) |
| `noteType` | enum | No | `"note"` | Visual style: `"note"` (blue), `"todo"` (green), `"important"` (red), `"question"` (orange) |
| `position` | enum | No | `"bottom"` | Where to insert: `"top"` or `"bottom"` |

**Behavior:** Creates a styled `<div>` block with a colored left border, background color, icon, type label, and timestamp. The note content is converted from markdown to HTML. Sent as a PATCH with `action: "prepend"` or `"append"` depending on position.

#### `addTableToPage`

Inserts a formatted HTML table from CSV data.

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `pageId` | string | Yes | | The page to add the table to |
| `tableData` | string | Yes | | Table in CSV format: first line is headers, subsequent lines are data rows. Columns separated by commas. |
| `title` | string | No | | Optional heading above the table |
| `position` | enum | No | `"bottom"` | Where to insert: `"top"` or `"bottom"` |

**Behavior:** Parses the CSV string by splitting on newlines (rows) and commas (columns). Requires at least two rows (one header + one data row). All cell values are HTML-escaped. The table is rendered with borders and padding.

**Example `tableData` input:**
```
Task,Status,Owner
Design review,Done,Alice
Backend API,In Progress,Bob
```

## Example Interactions with AI

Once connected and authenticated, you can ask your AI assistant to perform tasks like:

**Basic Operations:**
*   "List my OneNote notebooks."
*   "Show me the sections in my Work notebook."
*   "List the most recent pages in section XYZ."
*   "Create a new OneNote page titled 'Meeting Ideas' with the content 'Brainstorm new marketing strategies'."

**Advanced Search & Discovery:**
*   "Show me all notes I created or modified today." (uses `searchPagesByDate`)
*   "Find all pages created in the last 7 days that mention 'project'."
*   "Search for pages with 'meeting' in the title from the last 3 days."
*   "In my Personal notebook, find all pages from this week."
*   "Search for 'Project Phoenix' in page content" (searches actual content, not just titles!)

**Productivity & Standup Prep:**
*   "What pages did I modify since Monday?" (perfect for Thu standup prep!)
*   "Show me my recent changes from the last 3 days"
*   "Create today's note in SQLNikon/Elon" (auto-formats date)
*   "Create a daily note for 2/20/26 in Data Team/Stand Ups"

**Content Operations:**
*   "Can you find my OneNote page about 'Project Phoenix' and tell me its summary?"
*   "Append 'Follow up with John Doe' to the OneNote page with ID 'your-page-id-here'."
*   "In my OneNote page 'Recipe Ideas', replace all instances of 'sugar' with 'sweetener'."



## Troubleshooting

*   **Authentication Issues:**
    *   Ensure your `AZURE_CLIENT_ID` (if set) is correct and has the required API permissions.
    *   If the device code flow fails, try in a different browser or an incognito/private window.
    *   Token expiry: If tools stop working, you may need to re-run the `authenticate` tool.
*   **Server Not Starting:**
    *   Check Node.js version (`node -v`).
    *   Ensure all dependencies are installed (`npm install`).
*   **MCP Client Issues (e.g., Claude Desktop, Cursor):**
    *   Verify the `command` and `args` (especially the absolute path to `onenote-mcp.mjs`) in your client's MCP server configuration are correct.
    *   Restart the MCP client after making configuration changes.
    *   Check the MCP client's logs and the server's console output for errors.


## Security Notes

*   **Access Token Security:** The `.access-token.txt` file contains a token that grants access to your OneNote data according to the defined scopes. Protect this file as you would any sensitive credential. Ensure it is included in your `.gitignore` file.
*   **Azure Client ID:** If you create your own Azure App Registration, keep its client secret (if any generated for other flows) secure. For this device code flow, a client secret is not used by this script.
*   **Permissions:** This server requests `Notes.ReadWrite` and `Notes.Create` permissions. Be aware of the access you are granting.

## Acknowledgements

This project was developed with inspiration and by adapting patterns from the following open-source projects:

*   **[onenote-mcp](https://github.com/danosb/onenote-mcp) by danosb:** This project served as an early inspiration and provided reference for structuring a OneNote MCP server, particularly for initial concepts around authentication and basic OneNote operations.


*   **[azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks:** The core authentication flow using Device Code Credentials, token storage/retrieval strategy, and foundational patterns for wrapping Microsoft Graph API calls for OneNote (such as listing entities and creating pages) as MCP tools were significantly informed by or adapted from this project. This project is licensed under the MIT License.

The extensive set of editing tools, advanced text extraction and HTML processing utilities, Zod schema integration, and the overall refined structure of this server are original contributions. 

Development of this server was also assisted by AI language models, including Anthropic's Claude and Google's Gemini, for tasks such as code generation, refactoring, debugging, and documentation.

We are grateful to the authors of the referenced projects and the developers of the AI tools for their contributions to the open-source and development communities.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
