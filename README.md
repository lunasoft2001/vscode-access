# Access Explorer

> Desarrollado por **luna-soft** — Explora y edita bases de datos Microsoft Access (.accdb / .mdb) directamente desde VS Code.

---

## Características

- Conexiones múltiples a bases `.accdb` / `.mdb`
- Árbol de objetos por conexión: Tablas, Consultas, Formularios, Informes, Macros, Módulos VBA, Relaciones, Referencias
- Editor SQL integrado con resultados en rejilla (ordenar, filtrar, copiar, exportar CSV)
- Soporte CRUD: INSERT / UPDATE / DELETE con confirmación
- Editar y guardar código VBA directamente en Access
- Editar propiedades de controles de formularios/informes
- Compilar VBA y mostrar errores en el editor
- Compactar y reparar la base de datos
- Interfaz en **inglés**, **alemán** y **español** (según el idioma de VS Code)

---

## Requisitos previos

| Requisito | Detalle |
|-----------|---------|
| Windows | Requerido por Microsoft Access |
| Microsoft Access | 2016 o superior |
| Python 3.9+ | Necesario para el servidor MCP |
| [MCP-Access](https://github.com/modelcontextprotocol/servers) | Servidor MCP local para Access |

### Instalar MCP-Access

```powershell
pip install mcp-access
# o clona el repositorio y usa: pip install -e .
```

---

## Instalación de la extensión

### Opción A — Desde archivo VSIX (recomendado)

1. Descarga el archivo `access-explorer-x.x.x.vsix` desde la sección [Releases](../../releases) del repositorio.
2. En VS Code: **Extensions → ··· → Install from VSIX...** y selecciona el archivo.
3. O desde la terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Opción B — Compilar desde el código fuente

```powershell
git clone <url-del-repositorio>
cd AccessExtension
npm install
npm run compile
npx vsce package        # genera el .vsix
code --install-extension access-explorer-x.x.x.vsix
```

Requiere Node.js 18+ y `@vscode/vsce` (`npm install -g @vscode/vsce`).

---

## Configuración

Una vez instalada la extensión, configura en **Settings** de VS Code (`Ctrl+,`):

| Setting | Descripción | Ejemplo |
|---------|-------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Ruta absoluta a `access_mcp_server.py` | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Comando Python | `python` o `py` |
| `accessExplorer.mcp.toolPrefix` | Prefijo de herramientas del servidor | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout general (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout para SQL (ms) | `600000` |

---

## Uso rápido

1. Abre la vista **Access** en la barra lateral (icono de luna).
2. Pulsa **Access: Add Connection** y selecciona una base `.accdb`.
3. Expande las categorías para explorar objetos.
4. Clic derecho en una conexión para **Compilar VBA** o **Compactar y reparar**.
5. Abre un módulo VBA → edita → guarda con **Access: Save Code to Access** (icono de subida en el editor).
6. Abre el editor SQL (**Access: New SQL Query**), escribe tu consulta y ejecútala con el botón ▶.

---

## Créditos

Desarrollado por **luna-soft**.  
Basado en el protocolo [Model Context Protocol (MCP)](https://modelcontextprotocol.io) y el servidor [MCP-Access](https://github.com/modelcontextprotocol/servers).
