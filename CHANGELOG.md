# Changelog

All notable changes to this project are documented in this file.

## 1.0.19 - 2026-04-20

- Added automatic registration of the MCP-Access server in VS Code's user-level `mcp.json` on startup, so Copilot and other MCP-aware tools discover it without manual configuration.
- Added new command `Access: Register MCP Server in VS Code` to force registration and reload the window.
- The extension now resolves the correct Python executable (venv) when writing the `mcp.json` entry, preventing `ModuleNotFoundError: No module named 'mcp'` at server startup.
- Rewrote `BulkExportService` to use a separate clean runner database (`SecondBrainRunner.accdb`) instead of injecting VBA into the target database, avoiding `Application.Run` failures caused by pre-existing compile errors in the target project.
- VBA entry points now accept `targetDbPath` as first argument and call `Application.OpenCurrentDatabase` to switch context before exporting.

## 1.0.18 - 2026-04-20

- Hardened MCP-Access setup on Windows by preinstalling `cryptography` with `--only-binary :all:` to avoid source builds that require Rust/MSVC toolchains.
- Added automatic retry logic when dependency installation fails with common `cryptography` build errors (`link.exe`, `maturin`, wheel build failures).
- Improved installation diagnostics with explicit Windows ARM64 guidance and a no-Git ZIP-based manual fallback.

## 1.0.14 - 2026-04-06

- Added `Access: Show MCP Runtime` command to make runtime paths visible from the extension UI.
- Added one-click actions to copy a ready-to-use `mcp.json` snippet and open the managed runtime folder.
- Added runtime diagnostics view with resolved Python command and MCP server script path.

## 1.0.13 - 2026-04-06

- Added automatic detection and installation of `Pillow` when `from PIL import Image` fails during screenshot operations.
- Added explicit warning and guided repair flow for MCP-Access when `mcp_access` module is missing.
- Centralized import probe now validates `mcp`, `win32com.client`, `PIL` and `mcp_access.server` before starting the server.
- Updated setup help messages to include `Pillow` in the required Python packages.

## 1.0.12 - 2026-04-06

- Added automatic best-effort enabling of Access Trust Center VBA access (`AccessVBOM=1`) during prerequisite bootstrap.
- Improved environment guidance when VBA project access is still blocked by policy.

## 1.0.11 - 2026-04-06

- Switched MCP-Access default runtime location to extension-managed storage (`globalStorage`).
- Prioritized the managed runtime script before legacy user-level paths.
- Kept backward compatibility with custom `accessExplorer.mcp.serverScriptPath` and legacy paths.

## 1.0.10 - 2026-04-06

- Improved MCP-Access auto-installation for machines without Git.
- Added fallback download via ZIP from GitHub when `git clone` is unavailable or fails.
- Reduced cases where users had to install MCP-Access manually.

## 1.0.9 - 2026-04-06

- Added forced prerequisite validation when the extension activates.
- Added automatic Python installation attempt via `winget` when Python is missing.
- Made the extension continue with MCP-Access installation only after Python is available.
- Improved first-run behavior so missing prerequisites are handled before the user hits a generic connection error.

## 1.0.8 - 2026-04-06

- Fixed MCP-Access prerequisite setup to install the actual required Python packages: `mcp` and `pywin32`.
- Added guided environment diagnostics that explain the required order: Python, Python packages, Microsoft Access, and Trust Center VBA access.
- Improved connection failure messages for closed MCP server processes.
- Updated README installation instructions to match the real MCP-Access setup.

## 1.0.7 - 2026-04-06

- Fixed MCP-Access auto-install when repository is not pip-installable in editable mode (`pip install -e`).
- Added fallback to install dependencies from `requirements*.txt` files.
- Improved diagnostics text with manual alternative for non-packaged repositories.

## 1.0.6 - 2026-04-06

- Added automatic MCP-Access prerequisite detection when `access_mcp_server.py` is missing.
- Added guided auto-install flow for MCP-Access (clone/update repo, create venv, install dependencies).
- Added Python command validation with quick actions to open settings or download Python.
- Added detailed troubleshooting document and output logs when automatic installation fails.

## 1.0.5 - 2026-04-06

- General stability and usability improvements.
