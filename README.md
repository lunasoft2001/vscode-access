# Access Explorer

> Created by **[Juanjo Luna](https://blog.luna-soft.es/)** · [luna-soft](https://blog.luna-soft.es/) · [LinkedIn](https://www.linkedin.com/in/luna-soft/) · [GitHub](https://github.com/lunasoft2001)  
> Explore and edit Microsoft Access databases (.accdb / .mdb) directly from VS Code.

🌐 [Español](README.es.md) · [Deutsch](README.de.md)

---

## Features

- Multiple connections to `.accdb` / `.mdb` databases
- Object tree per connection: Tables, Queries, Forms, Reports, Macros, VBA Modules, Relationships, References
- Integrated SQL editor with results grid (sort, filter, copy, export CSV)
- CRUD support: INSERT / UPDATE / DELETE with confirmation dialog
- Edit and save VBA code directly into Access
- Edit form/report control properties
- Compile VBA and display errors in the editor
- Compact & Repair database
- UI available in **English**, **German** and **Spanish** (follows VS Code display language)

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| Windows | Required by Microsoft Access |
| Microsoft Access | 2016 or later |
| Python 3.9+ | Required to run the MCP server |
| [MCP-Access](https://github.com/unmateria/MCP-Access) | External MCP server (Python process) launched automatically by the extension |

### Install MCP-Access

Clone or download the server from [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access):

```powershell
git clone https://github.com/unmateria/MCP-Access.git
cd MCP-Access
pip install -e .
```

---

## Installation

### Option A — From VSIX file (recommended)

1. Download `access-explorer-x.x.x.vsix` from the [Releases](../../releases) section.
2. In VS Code: **Extensions → ··· → Install from VSIX...** and select the file.
3. Or from the terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Option B — Build from source

```powershell
git clone https://github.com/lunasoft2001/vscode-access.git
cd vscode-access
npm install
npm run compile
npx vsce package        # generates the .vsix file
code --install-extension access-explorer-x.x.x.vsix
```

Requires Node.js 18+ and `@vscode/vsce` (`npm install -g @vscode/vsce`).

---

## Configuration

After installing the extension, configure it in VS Code **Settings** (`Ctrl+,`):

| Setting | Description | Example |
|---------|-------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Absolute path to `access_mcp_server.py` | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Python command | `python` or `py` |
| `accessExplorer.mcp.toolPrefix` | Tool prefix used by the MCP server | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout for general MCP calls (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout for SQL execution (ms) | `600000` |

---

## Quick Start

1. Open the **Access** view in the activity bar (moon icon).
2. Click **Access: Add Connection** and select an `.accdb` file.
3. Expand categories to explore database objects.
4. Right-click a connection to **Compile VBA** or **Compact & Repair**.
5. Open a VBA module → edit → save with **Access: Save Code to Access** (upload icon in editor toolbar).
6. Open the SQL editor (**Access: New SQL Query**), write your query and run it with the ▶ button.

---

## Credits

**Access Explorer** is developed and maintained by [Juanjo Luna](https://blog.luna-soft.es/) — [luna-soft](https://blog.luna-soft.es/).

This extension uses **[MCP-Access](https://github.com/unmateria/MCP-Access)** as the backend server to communicate with Microsoft Access.  
MCP-Access is an independent project distributed under its own license. All rights belong to their respective authors.

Communication protocol: [Model Context Protocol (MCP)](https://modelcontextprotocol.io).

---

## License

© 2026 Juanjo Luna — [luna-soft](https://blog.luna-soft.es/)

This extension is licensed under the **[Polyform Noncommercial License 1.0.0](LICENSE)**.

**Free for personal, educational, and non-commercial use.**  
**Commercial use** (including integration into commercial products, services, or use by a for-profit entity for business purposes) **requires a separate written commercial license** from the author.

To inquire about commercial licensing, contact: [juanjo@luna-soft.es](mailto:juanjo@luna-soft.es)

