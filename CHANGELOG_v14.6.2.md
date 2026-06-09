# CHANGELOG — WEP AI v14.6.2 (patch de colores de marca)

**Tipo:** patch puramente visual. Sin cambios funcionales.
**Fecha:** 2026-05-21
**Base:** v14.6.1 (todos los features intactos)

## Cambio

### Colores de marca alineados con el icon oficial

El icon oficial de WEP AI muestra las letras "WEP" con tres colores específicos: W azul (Word), E verde (Excel), P naranja (PowerPoint). El branding tricolor **ya estaba implementado** en el código desde versiones anteriores, en varios puntos de la UI:

- Logo principal del login (W·E·P gigante)
- Sub-tag "Word · Excel · PowerPoint" cada palabra coloreada
- Header del sidebar del chat
- Etiquetas de tipo de documento en cada conversación
- Selector de tipo de documento (3 botones outlined)
- Iconos de archivos en la lista
- Progress bar de generación tomando el color del tipo

**Lo que cambió en v14.6.2:** los tonos de las 3 constantes de color para que coincidan exactamente con el icon entregado por el usuario:

| Constante | Antes (v14.6.1) | Ahora (v14.6.2) | Origen |
|---|---|---|---|
| `WORD_C`  | `#2d7bf4` (iOS system blue) | `#2196F3` (Material Blue 500) | Más vibrante y brillante |
| `EXCEL_C` | `#30d158` (Apple system green) | `#34A853` (Google green) | Saturación adecuada para fondo dark |
| `PPT_C`   | `#ff9f0a` (Apple system orange) | `#FB8C00` (Material Orange 600) | Más cálido, menos amarillo |

### Archivo modificado

- **`ui/chat_window.py`** — 3 líneas (constantes `WORD_C`, `EXCEL_C`, `PPT_C`) con comentario explicando el alineamiento al icon oficial.

## Sin breaking changes

Todos los widgets ya usaban estas constantes (no hay hardcoding de colores en otros archivos). Cambiar las 3 constantes propaga automáticamente a:

- Login screen (logo grande + sub-tag)
- Register screen (sub-tag + plan card de Pro)
- Chat sidebar (header logo + etiquetas de tipo en cada fila)
- Selector de tipo de documento (3 botones)
- Lista de documentos generados (iconos por extensión)
- Progress bar de generación

13/13 tests siguen pasando (los tests no validan colores específicos, solo funcionalidad).
