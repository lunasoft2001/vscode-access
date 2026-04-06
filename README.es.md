# Access Explorer

> Creado por **[Juanjo Luna](https://blog.luna-soft.es/)** · [luna-soft](https://blog.luna-soft.es/) · [LinkedIn](https://www.linkedin.com/in/luna-soft/) · [GitHub](https://github.com/lunasoft2001)  
> Explora y edita bases de datos Microsoft Access (.accdb / .mdb) directamente desde VS Code.

🌐 [English](README.md) · [Deutsch](README.de.md)

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
| Acceso VBA en Access | Activa `Trust access to the VBA project object model` en el Centro de confianza |
| [MCP-Access](https://github.com/unmateria/MCP-Access) | Servidor MCP externo (proceso Python) que la extensión lanza automáticamente |

### Instalar MCP-Access

Clona o descarga el servidor desde [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access):

```powershell
git clone https://github.com/unmateria/MCP-Access.git
cd MCP-Access
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install mcp pywin32
```

En Microsoft Access activa:

```text
Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Configuración de macros > Trust access to the VBA project object model
```

---

## Instalación

### Opción A — Desde archivo VSIX (recomendado)

1. Descarga `access-explorer-x.x.x.vsix` desde la sección [Releases](../../releases).
2. En VS Code: **Extensiones → ··· → Instalar desde VSIX...** y selecciona el archivo.
3. O desde la terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Opción B — Compilar desde el código fuente

```powershell
git clone https://github.com/lunasoft2001/vscode-access.git
cd vscode-access
npm install
npm run compile
npx vsce package        # genera el archivo .vsix
code --install-extension access-explorer-x.x.x.vsix
```

Requiere Node.js 18+ y `@vscode/vsce` (`npm install -g @vscode/vsce`).

---

## Configuración

Una vez instalada, configura en **Ajustes** de VS Code (`Ctrl+,`):

| Setting | Descripción | Ejemplo |
|---------|-------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Ruta absoluta a `access_mcp_server.py` | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Comando Python | `python` o `py` |
| `accessExplorer.mcp.toolPrefix` | Prefijo de herramientas del servidor MCP | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout para llamadas MCP generales (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout para ejecución SQL (ms) | `600000` |

---

## Uso rápido

1. Abre la vista **Access** en la barra lateral (icono de luna).
2. Pulsa **Access: Add Connection** y selecciona una base `.accdb`.
3. Expande las categorías para explorar los objetos.
4. Clic derecho en una conexión para **Compilar VBA** o **Compactar y reparar**.
5. Abre un módulo VBA → edita → guarda con **Access: Save Code to Access** (icono de subida en la barra del editor).
6. Abre el editor SQL (**Access: New SQL Query**), escribe tu consulta y ejecútala con el botón ▶.

---

## Créditos

**Access Explorer** es una extensión desarrollada por [Juanjo Luna](https://blog.luna-soft.es/) — [luna-soft](https://blog.luna-soft.es/).

Esta extensión utiliza **[MCP-Access](https://github.com/unmateria/MCP-Access)** como servidor backend para comunicarse con Microsoft Access.  
MCP-Access es un proyecto independiente distribuido bajo su propia licencia. Todos los derechos sobre MCP-Access pertenecen a sus respectivos autores.

Protocolo de comunicación: [Model Context Protocol (MCP)](https://modelcontextprotocol.io).

---

## Licencia

© 2026 Juanjo Luna — [luna-soft](https://blog.luna-soft.es/)

Esta extensión se distribuye bajo la licencia **[Polyform Noncommercial License 1.0.0](LICENSE)**.

**Gratuita para uso personal, educativo y no comercial.**  
**El uso comercial** (incluyendo la integración en productos o servicios comerciales, o el uso por parte de una entidad con ánimo de lucro) **requiere una licencia comercial escrita del autor**.

Para consultar licencias comerciales, contacta con: [juanjo@luna-soft.es](mailto:juanjo@luna-soft.es)
