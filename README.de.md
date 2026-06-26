# Access Explorer

> Erstellt von **[Juanjo Luna](https://blog.luna-soft.es/)** · [luna-soft](https://blog.luna-soft.es/) · [LinkedIn](https://www.linkedin.com/in/luna-soft/) · [GitHub](https://github.com/lunasoft2001)  
> Microsoft Access-Datenbanken (.accdb / .mdb) direkt aus VS Code erkunden und bearbeiten.

🌐 [English](README.md) · [Español](README.es.md)

---

Access Explorer richtet sich an Teams und Entwickler, die echte Access-Projekte pflegen und dabei die modernen Editor-Workflows von VS Code nutzen wollen.

## Was Du Damit Machen Kannst

- **Schnell erkunden**: mehrere `.accdb` / `.mdb`-Dateien verbinden und einen vollständigen Objektbaum nutzen (Tabellen, Abfragen, Formulare, Berichte, Makros, VBA-Module, Beziehungen, Referenzen).
- **SQL sicher ausführen**: einzelne SQL-Statements oder Batch-Skripte ausführen, mit Bestätigung bei destruktiven Operationen.
- **Daten übertragen**: Tabellen per CSV/XLSX importieren oder exportieren.
- **VBA komfortabel pflegen**: Code in VS Code öffnen, bearbeiten, zurück nach Access speichern, kompilieren und Fehler prüfen.
- **Refactoring mit Transparenz**: Verwendungen über VBA, gespeicherte Abfragen und Steuerelement-Eigenschaften finden.
- **Datenbanken stabil halten**: Komprimieren/Reparieren ausführen und MCP-Runtime-Status/Version prüfen.
- **Mehrsprachig arbeiten**: UI in **Englisch**, **Deutsch** und **Spanisch** (entsprechend der VS Code-Sprache).

---

## In 30 Sekunden

1. Öffne die **Access**-Ansicht in der Aktivitätsleiste.
2. Starte **Access: Add Connection** und wähle eine `.accdb`-Datei.
3. Klappe Objekte auf und öffne eine Abfrage oder ein Modul.
4. Für SQL: nutze **Access: Temporäre SQL-Abfrage** oder **Access: SQL-Abfrage ausführen**.
5. Für VBA: bearbeiten, dann **Access: Code in Access speichern**, danach **Access: Modul kompilieren**.
6. Für Datenaustausch: **Access: Daten importieren/exportieren** ausführen.

---

## Voraussetzungen

| Anforderung | Details |
|-------------|---------|
| Windows | Von Microsoft Access benötigt |
| Microsoft Access | 2010 oder neuer (jede Version mit VBE-Unterstuetzung) |
| Python 3.9+ | Zum Starten des MCP-Servers erforderlich |
| VBA-Zugriff in Access | `Trust access to the VBA project object model` im Trust Center aktivieren |
| [MCP-Access](https://github.com/unmateria/MCP-Access) | Wird beim ersten Verwenden automatisch von der Erweiterung verwaltet; die manuelle Einrichtung ist nur für fortgeschrittene Fälle (siehe Installation) |

---

## Installation

Installiere zuerst die Erweiterung. Für die normale Nutzung musst du **MCP-Access nicht separat installieren**: Beim ersten Verbinden mit einer Datenbank richtet die Erweiterung den MCP-Runtime automatisch ein.

### Option A — Aus dem Marketplace installieren (empfohlen)

1. Öffne VS Code und installiere **Access Explorer** über die Erweiterungen-Ansicht.
2. Öffne eine `.accdb`-Datei und führe **Access: Add Connection** aus.
3. Erlaube bei Nachfrage die automatische Installation des MCP-Runtimes.

### Option B — Aus einer VSIX-Datei installieren

1. `access-explorer-x.x.x.vsix` aus dem Bereich [Releases](../../releases) herunterladen.
2. In VS Code: **Erweiterungen → ··· → Aus VSIX installieren...** und Datei auswählen.
3. Oder über das Terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Option C — Aus dem Quellcode erstellen

```powershell
git clone https://github.com/lunasoft2001/vscode-access.git
cd vscode-access
npm install
npm run compile
npx vsce package        # erstellt die .vsix-Datei
code --install-extension access-explorer-x.x.x.vsix
```

Erfordert Node.js 18+ und `@vscode/vsce` (`npm install -g @vscode/vsce`).

### Erster Start

Beim ersten Verbinden mit einer Datenbank richtet Access Explorer den MCP-Runtime automatisch ein:

1. MCP-Access wird heruntergeladen.
2. Eine Python-Virtualenv wird erstellt.
3. Die erforderlichen Pakete werden installiert.
4. Der MCP-Runtime wird im von der Erweiterung verwalteten Ordner abgelegt.

Wenn Python oder Git fehlen, zeigt die Erweiterung geführte Wiederherstellungsaktionen an, statt still zu scheitern.

<details>
<summary>Manuelle MCP-Installation (nur wenn die automatische Einrichtung fehlschlägt oder wenn du einen separaten Runtime willst)</summary>

Server klonen oder herunterladen von [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access):

```powershell
# Option A: mit Git
git clone https://github.com/unmateria/MCP-Access.git
cd MCP-Access

# Option B: ohne Git (ZIP-Download)
$zip = "$env:TEMP\MCP-Access-main.zip"
Invoke-WebRequest https://github.com/unmateria/MCP-Access/archive/refs/heads/main.zip -OutFile $zip
Expand-Archive -Path $zip -DestinationPath . -Force
cd .\MCP-Access-main

py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install --only-binary :all: cryptography   # empfohlen auf Windows ARM64
.\.venv\Scripts\python.exe -m pip install mcp pywin32 Pillow
```

Anschließend `accessExplorer.mcp.serverScriptPath` in den VS Code-Einstellungen auf den vollständigen Pfad zu `access_mcp_server.py` setzen.

</details>

---

## Konfiguration

Nach der Installation sind folgende Einstellungen in den VS Code **Einstellungen** verfügbar (`Ctrl+,`). Alle Einstellungen sind optional — die Erweiterung funktioniert sofort mit der automatischen Einrichtung.

| Einstellung | Beschreibung | Beispiel |
|-------------|--------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Absoluter Pfad zu `access_mcp_server.py` (leer lassen für automatisch verwaltetes Runtime) | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Python-Befehl (leer lassen für automatisch verwaltetes Venv) | `python` oder `py` |
| `accessExplorer.mcp.toolPrefix` | Tool-Präfix des MCP-Servers | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout für allgemeine MCP-Aufrufe (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout für SQL-Ausführung (ms) | `600000` |

---

## Typische Workflows

### 1) SQL-Wartung und Diagnose

1. Öffne **Access: Temporäre SQL-Abfrage**.
2. Führe einzelne Statements oder mehrere Statements als Batch aus.
3. Prüfe Ergebnisraster oder Batch-Bericht.

### 2) VBA bearbeiten-kompilieren Schleife

1. Öffne ein Modul/Formular/Bericht aus dem Baum.
2. Bearbeite den Code in VS Code.
3. Speichere mit **Access: Code in Access speichern**.
4. Kompiliere mit **Access: Modul kompilieren** oder **Access: VBA kompilieren**.

### 3) Datenmigration mit CSV/XLSX

1. Starte **Access: Daten importieren/exportieren**.
2. Wähle Import oder Export.
3. Wähle Tabelle, Datei und Optionen (Kopfzeilen, Bereich/Spec je nach Typ).
4. Prüfe den erzeugten Bericht.

---

## Befehlsübersicht

| Befehl | Zweck |
|--------|-------|
| `Access: Add Connection` | Neue Access-Datenbank verbinden |
| `Access: Objekte suchen` | Objekte schnell finden und öffnen |
| `Access: Verwendungen suchen` | Text in VBA, Abfragen und Steuerelement-Eigenschaften suchen |
| `Access: Temporäre SQL-Abfrage` | Temporären SQL-Editor öffnen |
| `Access: SQL-Abfrage ausführen` | SQL direkt ausführen |
| `Access: Aktives SQL ausführen` | Auswahl (oder ganzes SQL-Dokument) ausführen |
| `Access: Daten importieren/exportieren` | Tabellendaten via CSV/XLSX übertragen |
| `Access: Code in Access speichern` | VBA/Form/Report-Code zurück nach Access schreiben |
| `Access: Modul kompilieren` / `Access: VBA kompilieren` | VBA validieren und Fehler anzeigen |
| `Access: Komprimieren und reparieren` | Wartung und Bereinigung der Datenbankdatei |
| `Access: MCP-Runtime anzeigen` | Runtime-Pfad/Version/Quelle und Update-Status prüfen |

---

## Versionsverlauf

Dieses README enthält jetzt eine kurze Zusammenfassung der Änderungen pro Version.
Für alle Details siehe [CHANGELOG.md](CHANGELOG.md).

- **v1.2.0**: Umfassende Überarbeitung der Marketplace-Dokumentation und klareres Onboarding (Was Du machen kannst, 30-Sekunden-Start, Workflows, Befehlsübersicht). Bessere Sichtbarkeit neuer Funktionen wie SQL-Batch-Ausführung, CSV/XLSX-Datenübertragung und Verwendungen suchen.
- **v1.1.4**: MCP-Versionserkennung verbessert — extrahiert jetzt die Version aus Commit-Nachrichten, wenn Git-Tags nicht verfügbar sind (z. B. "v0.7.43"), wodurch Benutzer bessere Sichtbarkeit des MCP-Builds erhalten.
- **v1.1.3**: Marketplace-Release-Bump für die neuesten Access Explorer-Fixes.
- **v1.0.14**: Fuegt den Befehl `Access: MCP-Runtime anzeigen` mit kopierbarem `mcp.json`-Block und Runtime-Ordner-Ansicht hinzu.
- **v1.0.12**: Aktiviert den VBA-Zugriff im Access Trust Center (`AccessVBOM`) beim Setup automatisch (best effort).
- **v1.0.11**: Verwendet standardmäßig einen von der Erweiterung verwalteten MCP-Runtime-Speicher (`globalStorage`).
- **v1.0.10**: Fügt ZIP-Fallback für MCP-Access hinzu, wenn Git nicht verfügbar ist.

Ab den nächsten Versionen wird dieser Abschnitt bei jedem Release aktualisiert (Fixes, Verbesserungen und neue Funktionen).

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
