# Fase 1b — Tabla N°, consulta previa y resumen

Archivos modificados: `Código.js` y `docs/index.html`. **No** se hizo `git push`, `clasp push` ni deploy.
Para que la página nueva funcione hay que publicar después el Apps Script (nueva versión del despliegue `/exec`)
y agregar en la fila 1 de "Base de datos" una columna con el encabezado **Tabla N°**.

## Cómo se hizo
Con scripts de Node en `/tmp/fase1b/` (en Windows `%TEMP%\fase1b\`):

| Script | Qué hace |
|---|---|
| `aplicar-codigo.js` | Reemplaza `Código.js` con `codigo-nuevo.js`. Antes comprueba que el archivo sea la versión anterior. |
| `aplicar-html.js` | Reemplaza 3 bloques de `docs/index.html`: el CSS desde `#lector` hasta `#pinError`, la `<section id="app">` (más el panel nuevo) y el `<script>` inline. Conserva la `API_URL`. Se detiene si cambia el SHA-256 del fondo SVG + `<header>` o del CSS base. |
| `verificar-html.js` | Hace `node --check` del script, revisa los id, los handlers, los textos y el formateo del código. |
| `probar-codigo.js` | Prueba `Código.js` con `SpreadsheetApp`, `LockService`, `PropertiesService`, `ContentService` y `Utilities` simulados (vía `vm`). |

## API (`Código.js`)
Todas las acciones exigen PIN (si no es válido, responden `{estado:'pin'}`).

| Acción | Entrada | Respuesta |
|---|---|---|
| `validarPin` | `pin` | `pin_ok` |
| `consultar` *(nueva)* | `codigo` | No escribe nada ni toma el lock. Responde `pendiente` {nombre, codigo} · `repetido` {nombre, codigo, hora, tabla} · `no_existe` · `duplicado` |
| `registrar` *(cambia)* | `codigo`, `tabla` | `ok` {nombre, codigo, hora, tabla} · `repetido` · `no_existe` · `duplicado` · `tabla_invalida` · `tabla_repetida` {tabla, nombre, codigo, codigoSolicitado} · `error` {mensaje} |
| `resumen` *(nueva)* | — | `resumen` {total, asistentes, pendientes} |

- **Tabla N°**: un solo número entero positivo por código. Se acepta `7` o `"7"`. Se rechazan vacío, 0, negativos,
  decimales, listas, `"1,2"` y textos, con `tabla_invalida` (esto se revisa antes de tomar el lock).
- **Columna**: se busca por su encabezado en la fila 1, sin importar mayúsculas, tildes, espacios de más ni el uso de `º`/`°`
  (`Tabla N°`, `TABLA Nº` y `Tábla N°` sirven). Si no existe, responde un `error` con el mensaje:
  *No existe la columna "Tabla N°" en la fila 1 de la hoja "Base de datos". Agrégala y vuelve a intentar.*
- **Dentro del lock**, `registrar` vuelve a leer la hoja completa y comprueba otra vez `no_existe`, `duplicado` y `repetido`.
  Después revisa si **otra fila ya tiene ese número de tabla**. También cuenta si la tabla quedó guardada como texto, por ejemplo `"12"`.
  Si alguien la tiene, responde `tabla_repetida` con el nombre y el código de esa persona, y no escribe nada.
  Si todo está bien, escribe Tabla N°, I = hora, H = TRUE y luego `flush`.
- **Resumen**: `total` = filas con código en la columna A (no cuenta las vacías); `asistentes` = de esas, las que tienen H = TRUE;
  `pendientes` = total − asistentes.
- La hoja se lee con un solo `getValues()` (encabezados + datos) en lugar de una lectura por celda.

## Página (`docs/index.html`)
1. **Panel superpuesto** (`#capa`, fijo sobre la pantalla con fondo semitransparente; la cámara se ve detrás).
   - Al escanear, tomar foto o escribir un código, la cámara se **pausa** sin cerrarse, se abre el panel y se llama a `consultar`.
   - `pendiente`: muestra el nombre, el código, el campo **Tabla N°** (`type=number`, `step=1`, `min=1`, teclado numérico)
     y los botones **Registrar** y **Cancelar**. Solo deja escribir dígitos (bloquea `e`, `+`, `-`, `.` y `,`).
     Antes de enviar se valida con `/^[1-9][0-9]*$/`. Enter registra y Escape cancela.
   - **Verde** (registrado): se cierra solo a los 3 s. **Amarillo** (ya había ingresado, con hora y tabla) y **rojo**
     (no existe, duplicado, error, tabla repetida): esperan al botón **Entendido**.
   - **`tabla_repetida`** se muestra en rojo: *"La Tabla N° X ya está asignada"*, con el nombre y el código de quien la tiene,
     y no registra nada. Al pulsar Entendido vuelve al campo Tabla N° de la misma persona, para escribir otro número
     o pulsar Cancelar. Lo mismo pasa con `tabla_invalida`.
   - Al cerrar el panel, la cámara se reanuda. Hay 4 s de espera antes de volver a aceptar el mismo QR si sigue frente a la cámara.
     Nunca hace falta cerrar la cámara para ver un resultado.
2. **Campo manual**: arranca con `PMB-` y pone los guiones solo (`PMB-000-X`: 3 dígitos y 1 letra en mayúscula).
   Toma los primeros 3 dígitos y, después, la primera letra; ignora lo demás. Al borrar no vuelve a poner un guion que la persona quitó.
   Soporta la composición de los teclados Android (formatea en `compositionend`).
   La prueba de formato `/^PMB-\d{3}-[A-Z]$/` se mantiene: si no se cumple, aparece *"Formato: PMB-000-X"* y no se consulta.
   El botón ahora dice "Buscar código escrito", porque primero consulta y luego pide la tabla.
3. **Resumen**: el contador local se reemplazó por dos cifras arriba de la app, **"Asistentes registrados: N"** y
   **"Pendientes por ingresar: M"**, con la hora de la última actualización. Se actualizan al entrar, después de cada intento
   de registro y cada 30 s. Si una petición coincide con otra en curso, se repite al terminar. Si el PIN deja de ser válido,
   vuelve a la pantalla de PIN.
- El estilo del bingo no cambió: el fondo SVG, el `<header>` y el CSS base tienen el mismo SHA-256 que antes.
  Los elementos nuevos usan la misma paleta (borde dorado `#fbbf24`, paneles `#0f172a`/`#1e293b`, verde/amarillo/rojo de siempre).

## Resultados de las verificaciones

### `node --check`
- `Código.js`: OK
- `<script>` inline de `docs/index.html`: OK

### `verificar-html.js` — TODO OK
| Verificación | Resultado |
|---|---|
| Un solo `<script>` inline; `node --check` | OK |
| id duplicados | ninguno |
| id usados por el JS (22): app, asistentes, btnCamara, btnCancelar, btnEntendido, btnEntrar, btnRegistrar, capa, formTabla, foto, lector, lectorFoto, login, manual, manualError, pendientes, pin, pinError, resultado, resumenEstado, tabla, tablaError | todos existen |
| Handlers inline (entrar, alternarCamara, leerFoto, registrarManual, cerrarCapa, confirmarRegistro, entendido) | todos definidos |
| Textos "Asistentes registrados:" / "Pendientes por ingresar:"; el contador viejo ya no existe | OK |
| `#manual` arranca con `PMB-`; `#tabla` es `type=number step=1`; el panel tiene Tabla N°, Registrar, Cancelar y Entendido | OK |
| Llamadas `consultar`, `registrar` (codigo + tabla numérica) y `resumen`; intervalo de 30 s; cierre del verde a los 3 s | OK |
| `tabla_repetida` se muestra en rojo | OK |
| URLs externas: solo `html5-qrcode@2.3.8` | OK |
| Formateo `PMB-000-X` (14 casos, ejecutando la función real) y regex de formato | OK |

### `probar-codigo.js` — 60/60 OK
Datos de prueba: 9 filas. Incluyen un ya ingresado con tabla 7, otro con la tabla `"12"` como texto, un código duplicado
(`PMB-009-Z` / `pmb-009-z `) y una fila vacía.

| Caso | Resultado esperado | |
|---|---|---|
| PIN incorrecto en validarPin / consultar / registrar / resumen | `pin`, sin escrituras | OK |
| validarPin correcto · acción desconocida | `pin_ok` · `error` | OK |
| consultar: sin ingresar | `pendiente` con nombre; no expone el índice interno | OK |
| consultar: ya ingresado | `repetido`, hora `22/11/2026 14:05`, tabla `7` | OK |
| consultar: no existe · duplicado · vacío | `no_existe` · `duplicado` · `no_existe` | OK |
| consultar nunca escribe ni toma el lock | 0 escrituras | OK |
| **resumen** inicial | total 7, asistentes 2, pendientes 5 (ignora la fila vacía) | OK |
| resumen después de registrar · hoja solo con encabezados | 7/3/4 · 0/0/0 | OK |
| **registrar ok** | `ok`; H = TRUE, I = fecha, Tabla N° = 3 (número); un `flush`; lock tomado y liberado | OK |
| registrar lee la hoja solo dentro del lock | 0 lecturas fuera del lock | OK |
| **repetido** (registrar dos veces) | `repetido` con tabla 3; no cambia la tabla | OK |
| **no existe** · **duplicado** | sin escrituras | OK |
| El lock siempre se libera | 4 tomas / 4 liberaciones | OK |
| **Sin tabla / inválida**: undefined, null, "", "   ", 0, "0", -3, "-3", 1.5, "1.5", "1,2", "3, 4", [1,2], "abc", "7a", "1e3" | `tabla_invalida` (16 casos), sin escrituras | OK |
| Tabla `" 15 "` | se acepta como 15 | OK |
| **Sin columna "Tabla N°"** | `error` claro; no escribe; libera el lock | OK |
| Encabezados `TABLA Nº`, `  tabla   n° `, `Tábla N°` en la columna K | encontrados por nombre; escribe en K | OK |
| **Tabla repetida**: tabla 7 (la tiene PMB-002-B) | `tabla_repetida` con nombre "LUIS GÓMEZ" y código PMB-002-B; no registra | OK |
| Tabla repetida guardada como texto (`"12"`) | `tabla_repetida` (PMB-010-E) | OK |
| Con otro número · reutilizar la tabla recién asignada | `ok` · `tabla_repetida` | OK |
| Carrera: otro celular registra el código entre consultar y registrar | `repetido`, sin sobrescribir | OK |
| Carrera: otro celular toma la misma tabla entre consultar y registrar | `tabla_repetida`, sin escrituras | OK |

## Pendiente (no se hizo, por indicación)
- `clasp push` y una nueva versión del despliegue web de Apps Script. La página nueva usa `consultar` y `resumen`,
  así que no funcionará contra el despliegue actual.
- Agregar la columna **Tabla N°** en la fila 1 de la hoja.
- `git push`. Tampoco se probó en un celular real ni en un navegador (solo con las verificaciones de arriba).
