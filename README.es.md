# Access Explorer

> Creado por **[Juanjo Luna](https://blog.luna-soft.es/)** · [luna-soft](https://blog.luna-soft.es/) · [LinkedIn](https://www.linkedin.com/in/luna-soft/) · [GitHub](https://github.com/lunasoft2001)  
> Explora y edita bases de datos Microsoft Access (.accdb / .mdb) directamente desde VS Code.

🌐 [English](README.md) · [Deutsch](README.de.md)

---

Access Explorer está pensado para mantener proyectos reales de Access con la experiencia de edición de VS Code, sin perder compatibilidad con Access.

## Qué Puedes Hacer

- **Explorar rápido**: conecta múltiples `.accdb` / `.mdb` y navega un árbol completo (Tablas, Consultas, Formularios, Informes, Macros, Módulos VBA, Relaciones, Referencias).
- **Trabajar SQL con seguridad**: ejecuta SQL suelto o lotes de varias sentencias, con confirmación para operaciones destructivas.
- **Mover datos**: importa y exporta datos de tablas en CSV/XLSX desde la propia extensión.
- **Mantener VBA cómodamente**: abre, edita y guarda código VBA desde VS Code; compila y revisa errores.
- **Refactorizar con contexto**: busca usos en VBA, SQL de consultas guardadas y propiedades de controles.
- **Mantener la base sana**: ejecuta Compactar/Reparar y consulta estado/versión del runtime MCP.
- **Trabajar en tu idioma**: interfaz en **inglés**, **alemán** y **español** (respeta el idioma de VS Code).

---

## En 30 Segundos

1. Abre la vista **Access** en la barra lateral.
2. Ejecuta **Access: Add Connection** y selecciona una base `.accdb`.
3. Expande objetos y abre una consulta o módulo.
4. Para SQL: usa **Access: SQL temporal** o **Access: Ejecutar consulta SQL**.
5. Para VBA: edita y pulsa **Access: Guardar código en Access**, luego **Access: Compilar módulo**.
6. Para intercambio de datos: ejecuta **Access: Importar/Exportar datos**.

---

## Requisitos previos

| Requisito | Detalle |
|-----------|---------|
| Windows | Requerido por Microsoft Access |
| Microsoft Access | 2010 o superior (cualquier version compatible con VBE) |
| Python 3.9+ | Necesario para el servidor MCP |
| Acceso VBA en Access | Activa `Trust access to the VBA project object model` en el Centro de confianza |
| [MCP-Access](https://github.com/unmateria/MCP-Access) | Lo gestiona automáticamente la extensión al primer uso; la instalación manual es solo para casos avanzados (ver Instalación) |

---

## Instalación

Instala primero la extensión. **No hace falta instalar MCP-Access por separado** para el uso normal: la primera vez que conectes una base de datos, la extensión descargará y configurará el runtime MCP automáticamente.

### Opción A — Instalar desde Marketplace (recomendado)

1. Abre VS Code e instala **Access Explorer** desde la vista de Extensiones.
2. Abre un archivo `.accdb` y ejecuta **Access: Add Connection**.
3. Cuando se solicite, permite la instalación automática del runtime MCP.

### Opción B — Instalar desde un archivo VSIX

1. Descarga `access-explorer-x.x.x.vsix` desde la sección [Releases](../../releases).
2. En VS Code: **Extensiones → ··· → Instalar desde VSIX...** y selecciona el archivo.
3. O desde la terminal:
```powershell
code --install-extension access-explorer-x.x.x.vsix
```

### Opción C — Compilar desde el código fuente

```powershell
git clone https://github.com/lunasoft2001/vscode-access.git
cd vscode-access
npm install
npm run compile
npx vsce package        # genera el archivo .vsix
code --install-extension access-explorer-x.x.x.vsix
```

Requiere Node.js 18+ y `@vscode/vsce` (`npm install -g @vscode/vsce`).

### Primer uso

Al conectarte por primera vez a una base de datos, Access Explorer configura automáticamente el runtime MCP:

1. Descarga MCP-Access.
2. Crea un entorno virtual de Python.
3. Instala los paquetes necesarios.
4. Guarda el runtime en la carpeta gestionada por la extensión.

Si falta Python o Git, la extensión mostrará acciones guiadas para recuperarse en lugar de fallar sin explicación.

<details>
<summary>Instalación manual del MCP (solo si la instalación automática falla o si quieres un runtime independiente)</summary>

Clona o descarga el servidor desde [github.com/unmateria/MCP-Access](https://github.com/unmateria/MCP-Access):

```powershell
# Opción A: con Git
git clone https://github.com/unmateria/MCP-Access.git
cd MCP-Access

# Opción B: sin Git (descarga ZIP)
$zip = "$env:TEMP\MCP-Access-main.zip"
Invoke-WebRequest https://github.com/unmateria/MCP-Access/archive/refs/heads/main.zip -OutFile $zip
Expand-Archive -Path $zip -DestinationPath . -Force
cd .\MCP-Access-main

py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install --only-binary :all: cryptography   # recomendado en Windows ARM64
.\.venv\Scripts\python.exe -m pip install mcp pywin32 Pillow
```

Luego establece `accessExplorer.mcp.serverScriptPath` en los Ajustes de VS Code con la ruta completa a `access_mcp_server.py`.

</details>

---

## Configuración

Una vez instalada, los siguientes ajustes están disponibles en **Ajustes** de VS Code (`Ctrl+,`). Todos son opcionales: la extensión funciona de inmediato con su configuración automática.

| Setting | Descripción | Ejemplo |
|---------|-------------|---------|
| `accessExplorer.mcp.serverScriptPath` | Ruta absoluta a `access_mcp_server.py` (dejar vacío para usar el runtime gestionado automáticamente) | `C:\\tools\\mcp-access\\access_mcp_server.py` |
| `accessExplorer.mcp.pythonCommand` | Comando Python (dejar vacío para usar el entorno virtual gestionado automáticamente) | `python` o `py` |
| `accessExplorer.mcp.toolPrefix` | Prefijo de herramientas del servidor MCP | `access` |
| `accessExplorer.mcp.requestTimeoutMs` | Timeout para llamadas MCP generales (ms) | `30000` |
| `accessExplorer.mcp.sqlQueryTimeoutMs` | Timeout para ejecución SQL (ms) | `600000` |

---

## Flujos de Trabajo Comunes

### 1) Mantenimiento y diagnóstico SQL

1. Abre **Access: SQL temporal**.
2. Ejecuta una sentencia o varias en lote.
3. Revisa la rejilla de resultados o el informe de batch.

### 2) Ciclo editar-compilar de VBA

1. Abre un módulo/formulario/informe desde el árbol.
2. Edita en VS Code.
3. Guarda con **Access: Guardar código en Access**.
4. Compila con **Access: Compilar módulo** o **Access: Compilar VBA**.

### 3) Migración de datos con CSV/XLSX

1. Ejecuta **Access: Importar/Exportar datos**.
2. Elige Importar o Exportar.
3. Selecciona tabla, archivo y opciones (cabeceras, rango/spec según aplique).
4. Revisa el informe generado.

---

## Comandos Clave

| Comando | Para qué sirve |
|---------|-----------------|
| `Access: Add Connection` | Conectar una nueva base de Access |
| `Access: Buscar objetos` | Buscar y abrir objetos rápidamente |
| `Access: Buscar usos` | Buscar texto en VBA, consultas y propiedades de controles |
| `Access: SQL temporal` | Abrir un editor SQL de trabajo |
| `Access: Ejecutar consulta SQL` | Lanzar SQL directamente |
| `Access: Ejecutar SQL activo` | Ejecutar selección SQL (o documento completo) |
| `Access: Importar/Exportar datos` | Transferir datos de tablas vía CSV/XLSX |
| `Access: Guardar código en Access` | Subir cambios de código VBA/form/report a Access |
| `Access: Compilar módulo` / `Access: Compilar VBA` | Validar VBA y mostrar errores |
| `Access: Compactar y reparar` | Mantenimiento y limpieza del archivo |
| `Access: Mostrar runtime MCP` | Ver ruta/versión/origen del runtime y estado de actualización |

---

## Historial de versiones

Este README incluye ahora un resumen corto de cambios por versión.
Para el detalle completo, consulta [CHANGELOG.md](CHANGELOG.md).

- **v1.1.4**: Detección de versión MCP mejorada — ahora extrae la versión de los mensajes de commit cuando no hay tags de Git (p. ej., "v0.7.43"), dando a los usuarios mejor visibilidad del build de MCP.
- **v1.1.3**: Marketplace release bump para publicar los últimos fixes de Access Explorer.
- **v1.0.18**: Endurece la instalación en Windows/ARM64 forzando wheel binaria para `cryptography`; añade reintento automático tras errores de compilación y guía manual sin Git (ZIP).
- **v1.0.13**: Instala automáticamente `Pillow` si falta PIL al capturar pantallas; flujo guiado de reparación cuando falta el módulo `mcp_access`.
- **v1.0.12**: Activa automáticamente (best effort) el acceso VBA del Trust Center (`AccessVBOM`) durante el setup.
- **v1.0.11**: Usa por defecto runtime MCP gestionado por la extensión (`globalStorage`).
- **v1.0.10**: Añade fallback por ZIP de MCP-Access cuando Git no está disponible.

A partir de las próximas versiones, esta sección se actualizará en cada release (fixes, mejoras y nuevas funcionalidades).

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
