# OneNote MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

The OneNote Ultimate MCP Server is a powerful Model Context Protocol (MCP) server that enables AI language models (LLMs) like Claude, and other AI assistants, to securely interact with your Microsoft OneNote data. It allows for reading, writing, searching, and comprehensive editing of your OneNote notebooks, sections, and pages directly through your AI interface.

This server provides a rich set of tools for advanced OneNote management, including robust text extraction, HTML content processing, and fine-grained page manipulation.

## Features

*   **Authentication:** Secure device code flow for Microsoft Graph API access.
*   **Read Operations:**
    *   List notebooks, sections, and pages.
    *   Search pages by title.
    *   Get page content in various formats (full HTML, readable text, summary).
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

This server exposes the following tools to your AI assistant:

**Authentication:**
*   `authenticate`: Initiates the device code authentication flow with Microsoft Graph.
*   `saveAccessToken`: Loads and verifies the locally saved access token.

**Reading OneNote Data:**
*   `listNotebooks`: Lists all your OneNote notebooks.
*   `searchPages`: Searches for pages by title across all notebooks. (Arg: `query` (optional string))
*   `getPageContent`: Retrieves the content of a specific OneNote page. (Args: `pageId` (string), `format` (enum: "text", "html", "summary", optional, default: "text"))
*   `getPageByTitle`: Finds a page by its title and retrieves its content. (Args: `title` (string), `format` (enum: "text", "html", "summary", optional, default: "text"))

**Editing & Creating OneNote Pages:**
*   `createPage`: Creates a new OneNote page in the first available section. (Args: `title` (string), `content` (string - HTML or markdown))
*   `updatePageContent`: Replaces the entire content of an existing page. (Args: `pageId` (string), `content` (string), `preserveTitle` (boolean, optional, default: true))
*   `appendToPage`: Adds new content to the end of an existing page. (Args: `pageId` (string), `content` (string), `addTimestamp` (boolean, optional, default: true), `addSeparator` (boolean, optional, default: true))
*   `updatePageTitle`: Changes the title of an existing page. (Args: `pageId` (string), `newTitle` (string))
*   `replaceTextInPage`: Finds and replaces text within a page. (Args: `pageId` (string), `findText` (string), `replaceText` (string), `caseSensitive` (boolean, optional, default: false))
*   `addNoteToPage`: Adds a formatted, timestamped note/comment to a page. (Args: `pageId` (string), `note` (string), `noteType` (enum: "note", "todo", "important", "question", optional, default: "note"), `position` (enum: "top", "bottom", optional, default: "bottom"))
*   `addTableToPage`: Adds a formatted table to a page from CSV data. (Args: `pageId` (string), `tableData` (string - CSV), `title` (string, optional), `position` (enum: "top", "bottom", optional, default: "bottom"))

## Example Interactions with AI

Once connected and authenticated, you can ask your AI assistant to perform tasks like:

*   "List my OneNote notebooks."
*   "Create a new OneNote page titled 'Meeting Ideas' with the content 'Brainstorm new marketing strategies'."
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
