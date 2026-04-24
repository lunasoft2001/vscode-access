# PRD — Capacidades de edición completas en Access Explorer

**Versión base**: 1.0.21  
**Fecha**: 2026-04-23  
**Objetivo**: Exponer en la extensión VS Code todas las capacidades de edición disponibles en el servidor MCP-Access, convirtiéndola en un IDE funcional para bases de datos Microsoft Access.

---

## 1. Estado actual (ya implementado)

| Capacidad | Comando / Mecanismo |
|---|---|
| Ver objetos del árbol (tablas, consultas, formularios, informes, módulos, macros) | Tree view |
| Ver detalles de cualquier objeto (read-only) | `showDetails` |
| Abrir código VBA de módulos/forms/reports en editor | `showDetails` → abre documento virtual |
| **Guardar código VBA** editado de vuelta a Access | `saveCodeToAccess` → `setCode` |
| Editor SQL (SELECT) | `openSqlEditor` + `executeActiveSql` |
| Ejecutar DML (INSERT/UPDATE/DELETE) | `executeDml` |
| Compilar proyecto VBA completo | `compileVba` |
| Compactar y reparar | `compactRepair` |
| Exportar vault SecondBrain | `secondBrain.full/category/object` |
| Exportar objetos VBA a ficheros | `exportObjects.full/category` |

---

## 2. Capacidades MCP disponibles NO expuestas aún

Inventario de métodos del cliente MCP que aún no tienen comando en la extensión:

| Método MCP | Descripción |
|---|---|
| `setControlProps` | Editar propiedades de cualquier control en un form/report |
| `getObjectScreenshot` | Captura de pantalla de un form/report abierto en Access |
| `getControlDefinition` | Definición detallada de un control concreto |
| `getControlAssociatedProcedures` | Procedimientos de evento asociados a un control |
| `getFormReportProperties` | Propiedades globales del form/report |
| `evalVba` | Evaluar expresión VBA arbitraria |
| `runVba` | Ejecutar procedimiento VBA por nombre |
| `compileModule` | Compilar un módulo VBA individual |
| `deleteVbaModule` | Eliminar un módulo VBA del proyecto |
| `createDatabase` | Crear una nueva base de datos .accdb |
| `closeAccess` | Cerrar Access desde la extensión |
| `getTableDataPreview` | Previsualizar datos de una tabla |
| `getTableFields` | Campos de una tabla (ya parcialmente usado en SecondBrain) |
| `getTableIndexes` | Índices de una tabla |
| `listReferences` | Referencias VBA del proyecto |
| `listLinkedTables` | Tablas vinculadas |
| `listStartupOptions` | Opciones de inicio de la base de datos |

---

## 3. Funcionalidades propuestas

### Grupo A — Edición VBA (prioridad alta)

#### A1. Crear nuevo módulo VBA
- **Trigger**: clic derecho en categoría "Módulos" → "Nuevo módulo…"
- **Flujo**: pide nombre → crea fichero virtual vacío en editor → al guardar (`saveCodeToAccess`) llama `setCode` con `object_type: "module"`
- **MCP**: `setCode`
- **Impacto**: bajo (reutiliza infraestructura existente)

#### A2. Eliminar módulo VBA
- **Trigger**: clic derecho sobre nodo de módulo → "Eliminar módulo"
- **Flujo**: confirmación modal → `deleteVbaModule` → refresh árbol
- **MCP**: `deleteVbaModule`
- **Impacto**: bajo

#### A3. Compilar módulo individual
- **Trigger**: clic derecho sobre nodo de módulo → "Compilar módulo"
- **Flujo**: `compileModule(connection, moduleName)` → notificación de resultado
- **MCP**: `compileModule`
- **Impacto**: bajo

#### A4. Consola VBA (REPL)
- **Trigger**: comando `Access: Abrir consola VBA`
- **Flujo**: panel WebView con campo de texto libre; el usuario escribe expresión/bloque y pulsa Ejecutar → `evalVba` → muestra resultado; historial de sesión
- **MCP**: `evalVba`, `runVba`
- **Impacto**: medio (requiere WebView nuevo)

---

### Grupo B — Edición de consultas (prioridad alta)

#### B1. Guardar consulta editada de vuelta a Access
- **Trigger**: en el editor SQL, comando "Guardar como consulta…" (o guardar si el documento es una consulta existente)
- **Flujo**: pide nombre (o usa el actual) → llama herramienta MCP `set_query_sql` si existe, o `eval_vba` con `CurrentDb.QueryDefs(name).SQL = "..."`
- **MCP**: `evalVba` (workaround) o nuevo método wrapper
- **Impacto**: medio

#### B2. Crear nueva consulta
- **Trigger**: `accessExplorer.newQuery` desde árbol o paleta
- **Flujo**: abre editor SQL vacío → al guardar, pregunta nombre → crea `QueryDef` en Access
- **MCP**: `evalVba`
- **Impacto**: medio

#### B3. Eliminar consulta
- **Trigger**: clic derecho sobre consulta → "Eliminar consulta"
- **Flujo**: confirmación → `evalVba("CurrentDb.QueryDefs.Delete '...'")` → refresh
- **MCP**: `evalVba`
- **Impacto**: bajo

---

### Grupo C — Inspección visual de formularios e informes (prioridad media)

#### C1. Captura de pantalla de form/report
- **Trigger**: clic derecho sobre form/report → "Ver captura"
- **Flujo**: `getObjectScreenshot` → muestra imagen en panel WebView o en notificación
- **MCP**: `getObjectScreenshot`
- **Impacto**: medio (requiere renderizado de imagen)

#### C2. Inspector de controles
- **Trigger**: al abrir detalles de un form/report, mostrar lista de controles con propiedades editables
- **Flujo**: `getControls` → tabla de controles en WebView; al editar un campo → `setControlProps`
- **MCP**: `getControls`, `getControlDefinition`, `setControlProps`
- **Impacto**: alto (requiere WebView con grilla editable)

#### C3. Navegar control → procedimiento de evento
- **Trigger**: en el inspector de controles, botón "Ir al evento" junto a cada control
- **Flujo**: `getControlAssociatedProcedures` → lista de procedimientos → abre el módulo de code en editor en la línea correspondiente
- **MCP**: `getControlAssociatedProcedures`, `getProcedureDocument`
- **Impacto**: medio

---

### Grupo D — Gestión de tablas (prioridad media)

#### D1. Previsualizar datos de tabla
- **Trigger**: clic derecho sobre tabla → "Previsualizar datos"
- **Flujo**: `getTableDataPreview(connection, tableName, limit=200)` → WebView con tabla paginada
- **MCP**: `getTableDataPreview`
- **Impacto**: medio (requiere WebView con tabla)

#### D2. Ver estructura (campos e índices)
- **Trigger**: ya parcialmente en `showDetails`; ampliar para mostrar índices
- **Flujo**: `getTableFields` + `getTableIndexes` → panel de detalles mejorado
- **MCP**: `getTableFields`, `getTableIndexes`
- **Impacto**: bajo (ampliar vista existente)

#### D3. Editar datos de tabla (DML guiado)
- **Trigger**: en la previsualización de datos, botón "Nueva fila" / edición inline
- **Flujo**: genera INSERT/UPDATE desde UI → `executeDml`
- **MCP**: `executeDml`
- **Impacto**: alto

---

### Grupo E — Gestión de la base de datos (prioridad baja-media)

#### E1. Crear nueva base de datos
- **Trigger**: paleta de comandos → "Access: Nueva base de datos…"
- **Flujo**: diálogo `saveDialog` para elegir ruta .accdb → `createDatabase` → añadir como conexión automáticamente
- **MCP**: `createDatabase`
- **Impacto**: bajo

#### E2. Ver referencias VBA
- **Trigger**: nodo "Referencias" en el árbol (bajo la conexión)
- **Flujo**: `listReferences` → lista en árbol con nombre, GUID y ruta
- **MCP**: `listReferences`
- **Impacto**: bajo

#### E3. Ver opciones de inicio
- **Trigger**: nodo "Inicio" o menú contextual de conexión → "Opciones de inicio"
- **Flujo**: `listStartupOptions` → WebView con propiedades editables vía `evalVba`
- **MCP**: `listStartupOptions`, `evalVba`
- **Impacto**: medio

#### E4. Gestión de tablas vinculadas
- **Trigger**: nodo "Tablas vinculadas" en árbol
- **Flujo**: `listLinkedTables` → lista con cadena de conexión; opción de relink vía `evalVba`
- **MCP**: `listLinkedTables`, `evalVba`
- **Impacto**: medio

---

## 4. Priorización y hoja de ruta

| Prioridad | ID | Nombre | Esfuerzo | Valor |
|---|---|---|---|---|
| 🔴 1 | A1 | Crear módulo VBA | Bajo | Alto |
| 🔴 2 | A2 | Eliminar módulo VBA | Bajo | Alto |
| 🔴 3 | A3 | Compilar módulo individual | Bajo | Alto |
| 🔴 4 | B1 | Guardar consulta editada | Medio | Alto |
| 🟠 5 | D1 | Previsualizar datos de tabla | Medio | Alto |
| 🟠 6 | D2 | Ver campos e índices mejorado | Bajo | Medio |
| 🟠 7 | A4 | Consola VBA REPL | Medio | Alto |
| 🟠 8 | B2 | Crear nueva consulta | Medio | Medio |
| 🟠 9 | B3 | Eliminar consulta | Bajo | Medio |
| 🟡 10 | C1 | Captura de pantalla form/report | Medio | Medio |
| 🟡 11 | C3 | Navegar control → evento | Medio | Medio |
| 🟡 12 | E1 | Crear nueva base de datos | Bajo | Medio |
| 🟡 13 | E2 | Ver referencias VBA | Bajo | Bajo |
| 🟡 14 | E4 | Gestión tablas vinculadas | Medio | Medio |
| 🟢 15 | C2 | Inspector de controles editable | Alto | Alto |
| 🟢 16 | D3 | Editar datos tabla (DML guiado) | Alto | Medio |
| 🟢 17 | E3 | Opciones de inicio editables | Medio | Bajo |

---

## 5. Consideraciones técnicas

### Patrón de documento virtual (reutilizar para A1, B1, B2)
Ya existe para módulos. Extender `codeDocuments` en `extension.ts` para soportar `query` como tipo. Al guardar, detectar el tipo y llamar la herramienta correspondiente.

### WebView reutilizable (para C1, C2, D1, D3, A4)
Crear un `WebviewPanel` genérico con:
- Modo tabla (D1, D3, C2)
- Modo imagen (C1)
- Modo consola (A4)

### Seguridad DML
Toda operación de escritura en datos (`executeDml`, `setControlProps`, `setCode`) debe pedir confirmación explícita si el objeto no tiene copia en SecondBrain reciente.

### Access debe estar cerrado para edición VBA
Igual que ahora: si Access está abierto al intentar `setCode`/`deleteVbaModule`, usar el flujo de `restartAccessProcesses` ya implementado.

---

## 6. Fuera de alcance (v1.x)

- Editor visual WYSIWYG de formularios (requeriría renderizado COM)
- Debugger VBA paso a paso
- Control de versiones integrado (git diff de módulos VBA) — cubierto parcialmente por SecondBrain
- Soporte Access ADP / SQL Server backend
