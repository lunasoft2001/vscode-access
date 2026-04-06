# Changelog

All notable changes to this project are documented in this file.

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
