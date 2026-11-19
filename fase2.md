# Fase 2 — Juego de bingo

> Estado: API desplegada (versión 4). Falta configurar las claves de juego y TV, cargar las tablas y hacer `git push` de las páginas.

## URLs
| Página | URL | Clave |
|---|---|---|
| Ingreso (sin cambios, salvo ceros a la izquierda) | https://duwier97-ux.github.io/Bingo/ | CLAVE_INGRESO (o el PIN actual) |
| Control del juego | https://duwier97-ux.github.io/Bingo/bingo.html | CLAVE_JUEGO |
| Pantalla (TV) | https://duwier97-ux.github.io/Bingo/pantalla.html | CLAVE_TV |
| API (Apps Script, misma URL de siempre) | https://script.google.com/macros/s/AKfycbzcTqnNERnPIBRAhMR-EiVYOlCC26XBM8rbkWK1nVx3XOcs_ETcaOi1EVIS2GtLcl1m/exec | — |

Las páginas `bingo.html` y `pantalla.html` solo existirán en GitHub Pages después del `git push` (pendiente).

## Cómo configurar las claves
En el editor de Apps Script: **Configuración del proyecto (⚙) → Propiedades del script → Agregar propiedad**.

| Propiedad | Para qué | Notas |
|---|---|---|
| `CLAVE_INGRESO` | Página de ingreso | Opcional. Si no existe, se sigue usando la propiedad `PIN` actual, así que el ingreso no se rompe. Si la creas, el PIN viejo deja de valer. |
| `CLAVE_JUEGO` | Control del juego y carga de tablas | Obligatoria para jugar. Puede hacer todo. |
| `CLAVE_TV` | Pantalla | Solo lectura: ve el estado, no puede cambiar nada. |

- Usa claves distintas entre sí. Una clave vacía nunca se acepta.
- Las claves **no** están en el código ni en el repositorio.
- Para cambiar una clave, edita la propiedad: la TV pedirá la nueva clave sola.

## Datos (paso 1)
- `datos/tablas.json` y `datos/tablas.csv` (fuera de git): 300 filas. Se quitó la fila duplicada de 2135 (Documento.pdf p18; se dejó la p17).
- Columnas nuevas: `activa` y `nota`. **0762** queda con `activa = false` y la nota "Por confirmar". Se activa o desactiva desde el control (sección "Tablas por confirmar") o con la casilla de la columna Activa en la hoja.

## Carga de tablas (paso 2) — pendiente
Las tablas van a una hoja privada **"Tablas"** dentro del mismo Sheet, nunca al repositorio. Se suben con la acción `cargarTablas` y la `CLAVE_JUEGO`.

Script listo, fuera del repo: `C:UsersDuwier Mosqueraingo-ocrase2subir-tablas.js`. Lee la clave de la variable de entorno `CLAVE_JUEGO`, para que no quede escrita en ningún archivo:

```
CLAVE_JUEGO=tu-clave node ~/bingo-ocr/fase2/subir-tablas.js
```

Debe responder `tablas_cargadas` con 300 tablas (299 activas y 1 por confirmar). Se puede repetir: reemplaza la hoja completa.

## API (`Código.js`)
Acciones de ingreso sin cambios (`validarPin`, `consultar`, `registrar`, `resumen`). Acciones nuevas, todas con `clave`:

| Acción | Clave | Qué hace |
|---|---|---|
| `validarClave` | juego o TV | Devuelve el rol (`juego` o `tv`) |
| `estado` | juego o TV | Balotas, actual, figura, premio, ronda, anuncio |
| `marcarBalota` / `quitarBalota` / `deshacerBalota` | juego | Marca una balota (1–75), quita una concreta o la última |
| `fijarFigura` | juego | 12 figuras predefinidas o una personalizada de 5×5 |
| `fijarPremio` | juego | Premio en COP (acepta `1500000` o `$ 1.500.000`) |
| `ganadores` | juego | Calcula los ganadores **solo entre tablas con ingreso** y activas |
| `confirmarGanador` | juego | Lo verifica otra vez, lo anota en la hoja "Ganadores" y lo muestra en la TV |
| `ocultarAnuncio` / `nuevaRonda` | juego | Quita el anuncio de la TV / borra las balotas de la ronda |
| `listarTablas` / `activarTabla` / `cargarTablas` | juego | Administración de la hoja "Tablas" |

- El centro (N3) es el número de tabla y cuenta siempre como marcado.
- El estado se guarda en Propiedades del script (`ESTADO_JUEGO`), con una copia en caché para el polling de la TV.

## Despliegue (hecho el 2026-11-19)

1. Antes de subir se bajó el código remoto con `clasp clone`: era idéntico al de la fase 1b, así que no se sobrescribió ningún cambio hecho en el editor web.
2. `clasp push` (solo `Código.js` y `appsscript.json`, según `.claspignore`) → **versión 4**.
3. Se creó un despliegue temporal de prueba en la versión 4 y se verificó **antes** de tocar producción:
   - API: `doGet` responde `activo`; `validarPin`, `consultar`, `registrar` y `resumen` con un PIN incorrecto responden `pin`; las acciones del juego sin clave responden `clave`; un JSON roto da un `error` controlado.
   - Página de ingreso nueva en Chrome contra ese despliegue: con un PIN incorrecto muestra **"PIN incorrecto"**.
4. Como el ingreso funcionaba, se actualizó el despliegue de **producción** a la versión 4 (`clasp redeploy`). **La URL `/exec` no cambió.** Se repitieron las pruebas contra producción, incluida la página de ingreso que hoy está en GitHub Pages: responde "PIN incorrecto".
5. Se borró el despliegue temporal. Quedan solo `@HEAD` y producción `@4`.

**No se pudo probar con el PIN real** (no lo tengo): falta un registro real. Antes del evento, entra a la página de ingreso con el PIN y comprueba que el resumen carga.

**Latencia de Apps Script:** cada llamada tardó de 2,5 a 6 s, con picos de 14–26 s, y hubo respuestas vacías aisladas justo después del redeploy (al repetir, respondieron bien). La TV consulta cada 2 s, pero la siguiente consulta sale cuando termina la anterior, así que en la práctica se actualiza cada 3–6 s. El control marca las balotas en fila, en orden, aunque se toquen rápido.

## Pruebas
- Ingreso: 60/60 pruebas de la fase 1b siguen pasando con el código nuevo.
- Juego: 106/106 pruebas con servicios simulados y las 300 tablas reales.
- Páginas: sintaxis, ids y handlers verificados; capturas en Chrome headless de la TV (1920×1080 y 1366×768), con anuncio, y del control.

## Pendiente
- [x] Vista de celular del control: sin desbordes a 360 y 390 px (medido en iframes de ancho exacto). Leyenda de colores en los cartones de los ganadores y números de tabla siempre con 4 dígitos (0766).
- [x] `clasp push` + versión 4 en el despliegue `/exec` (verificado antes y después).
- [x] `CLAVE_JUEGO` y `CLAVE_TV` configuradas en Propiedades del script (lo confirmó quien administra el proyecto; no están en el repo).
- [ ] Cargar las tablas en la hoja "Tablas" (necesita la `CLAVE_JUEGO`; ver abajo).
- [x] `git push` de `Código.js`, `docs/index.html`, `docs/bingo.html`, `docs/pantalla.html`, `fase2.md` y `.gitignore`. Quedan fuera `datos/`, los PDFs y `revisar-tablas.md` (en `.gitignore`), además de los .md de análisis sin rastrear.
