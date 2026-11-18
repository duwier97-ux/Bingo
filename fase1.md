# Fase 1 — Cambios visuales en `docs/index.html`

Solo cambió lo visual. El `<script>` inline, los `id` y los `onclick`/`onchange` no se tocaron.

## Cómo se hizo
Con un script de Node (`/tmp/fase1-bingo.js`, en Windows `%TEMP%\fase1-bingo.js`) que reemplaza tres bloques:
`<title>`, `<style>` y `<header>`. Antes de escribir el archivo, el script compara el SHA-256 del `<script>` inline
de antes y de después, y se detiene si no son iguales.

## Qué cambió
1. **Título**
   - `<title>`: `Gran Bingo Municipal · 22 de noviembre de 2026`
   - Header: título "Gran Bingo Municipal", subtítulo "22 de noviembre de 2026" y una fila decorativa de bolitas B‑I‑N‑G‑O.
2. **Fondo de bingo, solo con CSS y SVG inline (sin imágenes externas)**
   - `body`: degradado oscuro (índigo/morado/vino) con resplandores radiales y una cuadrícula tipo cartón de 56 px, hecha con `linear-gradient`.
   - `<svg class="fondo-bingo">` fijo detrás del contenido (`aria-hidden`, `pointer-events:none`, opacidad 0.55):
     10 bolas de colores con letra y número según la columna (B 1‑15 azul, I 16‑30 rojo, N 31‑45 morado,
     G 46‑60 verde, O 61‑75 naranja) y 2 cartones 5×5 inclinados. Las posiciones están en % para que se repartan en cualquier pantalla.
3. **Legibilidad**
   - `main` es ahora un panel oscuro semiopaco (`rgba(15,23,42,.86)` + blur) con borde dorado, así el fondo no compite con el contenido.
   - Los botones conservan sus colores (azul `#2563eb` / gris `#334155`); se agregaron sombra, estado `:active` y `:disabled`.
   - Los resultados conservan sus colores: verde `#16a34a`, amarillo `#d97706`, rojo `#dc2626` e info `#334155`.
     Se agregaron borde claro, sombra y `text-shadow` sutil.
   - Los inputs tienen un foco dorado visible.
   - Móvil: el panel lleva 10 px de margen lateral. Desde 560 px de ancho se centra y el título sube a 30 px.

## Verificaciones (script `/tmp/fase1-verif.js`)
| Verificación | Resultado |
|---|---|
| `node --check` sobre el `<script>` inline extraído | OK |
| `<script>` inline idéntico al de `HEAD` (SHA-256, normalizando CRLF) | true |
| id usados por el JS (`$('…')`, `getElementById`, `new Html5Qrcode('…')`): pin, btnEntrar, pinError, login, app, btnCamara, foto, manual, contador, resultado, lector, lectorFoto | todos existen |
| id originales conservados | true |
| id duplicados | ninguno (único id nuevo: `bingoBrillo`, el degradado del SVG) |
| handlers inline `onclick`/`onchange` sin cambios | true |
| URLs externas | solo la que ya estaba (`html5-qrcode@2.3.8` en unpkg) |

También se revisó visualmente con Chrome headless a 360 px y 900 px de ancho: pantalla de PIN y pantalla de la app
con resultados verde/rojo/amarillo, usando una copia de prueba en `/tmp`. El título, los botones y los resultados se leen bien.
