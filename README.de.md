# Access Explorer

> Erstellt von **[Juanjo Luna](https://blog.luna-soft.es/)** · [luna-soft](https://blog.luna-soft.es/) · [LinkedIn](https://www.linkedin.com/in/luna-soft/) · [GitHub](https://github.com/lunasoft2001)  
> Microsoft Access-Datenbanken (.accdb / .mdb) direkt aus VS Code erkunden und bearbeiten.

🌐 [English](README.md) · [Español](README.es.md)

---

## Funktionen

- Mehrere Verbindungen zu `.accdb` / `.mdb`-Datenbanken
- Objektbaum je Verbindung: Tabellen, Abfragen, Formulare, Berichte, Makros, VBA-Module, Beziehungen, Referenzen
- Integrierter SQL-Editor mit Ergebnisraster (sortieren, filtern, kopieren, CSV-Export)
- CRUD-Unterstützung: INSERT / UPDATE / DELETE mit Bestätigungsdialog
- VBA-Code direkt in Access bearbeiten und speichern
- Eigenschaften von Formular-/Berichtssteuerelementen bearbeiten
- VBA kompilieren und Fehler im Editor anzeigen
- Datenbank komprimieren und reparieren
- Benutzeroberfläche auf **Englisch**, **Deutsch** und **Spanisch** (folgt der VS Code-Anzeigesprache)

---

## Voraussetzungen

| Anforderung | Details |
|-------------|---------|
| Windows | Von Microsoft Access benötigt |
| Microsoft Access | 2016 oder neuer |
| Python 3.9+ | Zum Starten des MCP-Servers erforderlich |
| [MCP-Access](https://github.com/unmateria/MCP-Access) | Externer MCP-Server (Python-Prozess), der automatisch von der Erweiterung gestartet wird |

### MCP-Access installieren

Server klonen oder herunterladen von [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access):

```powershell
git clone https://github.com/unmateria/MCP-Access.git
cd MCP-Access
pip install -e .
```

---

## Installation

### Option A — Aus VSIX-Datei (empfohlen)

1. `access-explorer-x.x.x.vsix` aus dem Bereich [Releases](../../releases) herunterladen.
2. In VS Code: **Erweiterungen → ··· → Aus VSIX installieren...** und Datei auswählen.
3. Oder über das Terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Option B — Aus dem Quellcode erstellen

```powershell
git clone https://github.com/lunasoft2001/vscode-access.git
cd vscode-access
npm install
npm run compile
npx vsce package        # erstellt die .vsix-Datei
code --install-extension access-explorer-x.x.x.vsix
```

Erfordert Node.js 18+ und `@vscode/vsce` (`npm install -g @vscode/vsce`).

---

## Konfiguration

Nach der Installation in den VS Code **Einstellungen** konfigurieren (`Ctrl+,`):

| Einstellung | Beschreibung | Beispiel |
|-------------|--------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Absoluter Pfad zu `access_mcp_server.py` | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Python-Befehl | `python` oder `py` |
| `accessExplorer.mcp.toolPrefix` | Tool-Präfix des MCP-Servers | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout für allgemeine MCP-Aufrufe (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout für SQL-Ausführung (ms) | `600000` |

---

## Schnellstart

1. Die **Access**-Ansicht in der Aktivitätsleiste öffnen (Mond-Symbol).
2. **Access: Add Connection** klicken und eine `.accdb`-Datei auswählen.
3. Kategorien aufklappen, um Datenbankobjekte zu erkunden.
4. Rechtsklick auf eine Verbindung für **VBA kompilieren** oder **Komprimieren und reparieren**.
5. Ein VBA-Modul öffnen → bearbeiten → mit **Access: Save Code to Access** speichern (Upload-Symbol in der Editor-Symbolleiste).
6. SQL-Editor öffnen (**Access: New SQL Query**), Abfrage schreiben und mit der Schaltfläche ▶ ausführen.

---

## Danksagungen

**Access Explorer** wurde entwickelt von [Juanjo Luna](https://blog.luna-soft.es/) — [luna-soft](https://blog.luna-soft.es/).

Diese Erweiterung verwendet **[MCP-Access](https://github.com/unmateria/MCP-Access)** als Backend-Server zur Kommunikation mit Microsoft Access.  
MCP-Access ist ein unabhängiges Projekt unter eigener Lizenz. Alle Rechte liegen bei den jeweiligen Autoren.

Kommunikationsprotokoll: [Model Context Protocol (MCP)](https://modelcontextprotocol.io).

---

## Lizenz

© 2026 Juanjo Luna — [luna-soft](https://blog.luna-soft.es/)

Diese Erweiterung ist lizenziert unter der **[Polyform Noncommercial License 1.0.0](LICENSE)**.

**Kostenlos für persönliche, bildungsbezogene und nicht-kommerzielle Nutzung.**  
**Kommerzielle Nutzung** (einschließlich der Integration in kommerzielle Produkte oder Dienste oder der Nutzung durch eine gewinnorientierte Einheit) **erfordert eine separate schriftliche kommerzielle Lizenz** des Autors.

Für Anfragen zur kommerziellen Lizenzierung wenden Sie sich an: [juanjo@luna-soft.es](mailto:juanjo@luna-soft.es)
