# WEP AI — Asistente inteligente de Office para macOS
## Word · Excel · PowerPoint con Inteligencia Artificial
### Fase 1 MVP — Agente conversacional con control directo de Office

---

## ¿Qué es WEP AI?

WEP AI es un asistente de IA que controla Microsoft Office directamente en tu Mac:

1. **Conversas** con la IA describiendo lo que necesitas
2. **Abre Office** automáticamente (Word, Excel o PowerPoint)
3. **Crea el documento** completo con IA
4. **Te muestra el resultado** y pregunta si está bien
5. **Hace correcciones** hasta que quede perfecto

**WEP** = **W**ord · **E**xcel · **P**owerPoint

---

## Instalación rápida

### Requisitos
- macOS 12 (Monterey) o superior
- Microsoft Office instalado (Word, Excel, PowerPoint)
- Python 3.11+ → [Descargar](https://www.python.org/downloads/)
- API Key de Anthropic → [Obtener](https://console.anthropic.com)

### Pasos

```bash
# 1. Descomprime el ZIP en el Escritorio
cd ~/Desktop/wepai

# 2. Ejecuta el instalador
bash install.sh

# 3. Inicia WEP AI
bash run_wepai.sh
```

---

## Estructura del proyecto

```
wepai/
├── main.py                  # Punto de entrada
├── requirements.txt         # Dependencias Python
├── install.sh               # Instalador automático
├── run_wepai.sh             # Inicio rápido
│
├── ui/
│   └── chat_window.py       # Interfaz de usuario
│
├── agent/
│   ├── brain.py             # Motor de IA (Claude API)
│   └── vision.py            # Capturas + verificación visual
│
├── office/
│   ├── word_controller.py   # Control de Microsoft Word
│   ├── excel_controller.py  # Control de Microsoft Excel
│   └── ppt_controller.py    # Control de PowerPoint
│
└── storage/
    └── db.py                # Base de datos local (SQLite)
```

---

## Ejemplos de uso

### Word — Documento legal
```
"Necesito una demanda ante Tigo por incumplimiento de contrato.
Soy Juan Vásquez, llevo 3 meses sin internet y me siguen cobrando.
Estoy en Colombia, contrato número 4589234."
```

### Excel — Inventario + estadísticas
```
"Necesito una tabla de inventario para mis cremas para la cara
y las estadísticas de ventas de mayo con ganancias y gastos."
```

### PowerPoint — Presentación empresarial
```
"Presenta mi empresa SportFlow a distribuidores en Colombia.
Tenemos 4 años, vendemos en Bogotá y Medellín, 2000 piezas al mes,
15 tiendas activas. Necesito 10 slides, estilo ejecutivo."
```

---

## Documentos generados

Todos los documentos se guardan en:
```
~/Documents/WEP_AI/
```

Historial de conversaciones:
```
~/.wepai/wepai.db
```

---

## Configurar API Key

```bash
# Temporal (solo esta sesión)
export ANTHROPIC_API_KEY='sk-ant-...'

# Permanente (se aplica siempre)
echo "export ANTHROPIC_API_KEY='sk-ant-...'" >> ~/.zshrc
source ~/.zshrc
```

---

## Permisos macOS necesarios

Ve a **Preferencias del Sistema → Privacidad y Seguridad** y activa:
- **Grabación de pantalla** — para tomar capturas del documento
- **Accesibilidad** — para controlar Office
- **Automatización** — para AppleScript

---

## Versiones futuras

| Versión | Función |
|---------|---------|
| v1.1 | Templates corporativos con logo propio |
| v2.0 | Importar datos desde CSV/Excel real |
| v2.5 | Multi-usuario con roles y permisos |
| v3.0 | Reportes automáticos programados |
| v4.0 | Integración con ERP (SAP, Odoo) |
| Enterprise | Multi-empresa, auditoría, SSO |

---

© 2026 WEP AI — Word · Excel · PowerPoint con Inteligencia Artificial
Desarrollado con Claude (Anthropic)

---

## Planes y límites (v1.1)

| Plan | Precio | Docs/mes | Historial | Correcciones IA | Usuarios |
|------|--------|----------|-----------|-----------------|----------|
| Personal Free | $0 | 5 | 7 días | ✗ | 1 |
| Personal Pro | $4.99/mes | Ilimitado | 90 días | ✓ | 1 |
| Enterprise | $24.99/mes | Ilimitado | 1 año | ✓ | Hasta 10 |

El plan Free muestra un aviso cuando el usuario llega al límite y ofrece actualizar a Pro.

## Cambios en v1.1
- Login con autenticación real (SQLite)
- Registro de nuevo usuario en 3 pasos (Datos → Plan → Confirmación)
- 3 planes: Free, Pro, Enterprise
- Límites del plan Free programados y aplicados
- Bilingüe completo (ES/EN) en login, registro y app
- Menú de usuario con nombre, email y plan activo
- Historial de documentos por usuario
