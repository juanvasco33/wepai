# CHANGELOG v16.0 — Enfoque Colombia + Versión Web

Fecha: 29 de mayo de 2026
Tipo: cambio mayor de producto (enfoque de mercado + canal de distribución)

## Resumen

Tres cambios grandes en esta versión:

1. **Enfoque exclusivo en Colombia y español** (sin amputar el motor multipaís).
2. **Auditoría de los tres formatos** (Word/Excel/PowerPoint) — verificada con
   generación real, no con documentación.
3. **Versión web** (FastAPI) que reutiliza el motor existente, para que la app
   funcione en Windows, macOS, Linux y móvil sin instalación.

---

## 1. Enfoque Colombia + español

La arquitectura multipaís (20 países) sigue **intacta bajo el capó**. Solo cambió
el comportamiento por defecto y la presentación:

- `office/legal_config.py`:
  - `_get_legal_config()` ahora cae a Colombia (CO) en lugar de GENERIC.
  - `_detect_country()` usa Colombia como default cuando no se detecta otro país.
  - Los 20 países siguen disponibles: si el texto menciona otro país, se respeta.
- `agent/brain.py`:
  - El system prompt se reescribió para Colombia: ya no pregunta por país,
    aplica automáticamente CST, Código Civil/Comercio, DIAN, EPS, IVA 19%, COP.
  - Tras parsear la respuesta del LLM, se fuerza `idioma="es"` y `pais="Colombia"`
    en `gen_data`, sin importar lo que sugiera el modelo o la UI.

**Verificación:** documento generado sin especificar país → Colombia (IVA 19%,
COP, DIAN). México sigue accesible bajo el capó (IVA 16%). 20 países vivos.

## 2. Auditoría de los tres formatos

Se ejecutó generación real (modo estático, sin API key) para confirmar que los
builders cumplen lo que prometen:

- **Word:** 10/10 documentos generados (contrato servicios, NDA, renuncia,
  factura, poder notarial, arrendamiento, acta SAS, oferta laboral,
  amonestación, finiquito). Cita CST, sin contaminación de otros países.
- **Excel:** 8/8 (presupuesto anual, nómina, flujo de caja, estado de
  resultados, KPIs, ratios financieros, inventario, orden de compra).
  IVA 19% y COP correctos. Incluye los builders antes "huérfanos".
- **PowerPoint:** 4/4 (corporativa, informe trimestral, plan de negocios,
  propuesta comercial). 19 slides en el pitch de prueba.
- **Suite de tests:** 301/301 pasan tras los cambios.

Todos los archivos generados son ZIP válidos (formato Office correcto).

## 3. Versión web (nuevo)

Nueva carpeta `web/`:

- `web/server.py` — backend FastAPI que envuelve `agent.brain` y los controllers.
  - `POST /api/chat` — conversa y, si procede, genera el documento.
  - `GET /api/download/{token}` — descarga el documento generado.
  - `GET /api/health` — estado del servidor y del motor.
- `web/static/index.html` — frontend de una sola página, en español, con el chat,
  ejemplos clicables, tarjeta de documento con sello "✓ Verificado con leyes
  vigentes de Colombia" y botón de descarga.

El motor (brain + controllers + verificadores) **no se modificó para esto**: la
versión web y la de escritorio comparten exactamente la misma lógica de
generación. Migrar a web no implicó reescribir el núcleo.

**Cómo correr la versión web:**

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
cd web
python3 server.py            # o: uvicorn server:app --host 0.0.0.0 --port 8000
# abrir http://localhost:8000
```

**Verificación:** servidor levantado, `/api/health` responde OK
(engine_loaded=true), sirve el frontend, y el flujo completo
chat → generación → descarga produce un .docx válido de ~40KB.

---

## Notas / pendientes (no resueltos en esta versión)

- La validación de los verificadores con API key real (tarea 4.1) sigue
  pendiente — es independiente de estos cambios.
- `save_overrides` aún no valida rangos antes de persistir (riesgo señalado).
- La versión de escritorio (`ui/chat_window.py`) sigue presente y funcional;
  no se eliminó. Si se decide ir 100% web, puede retirarse más adelante.
- Autenticación, cobro y persistencia de sesiones NO están implementados en la
  versión web todavía (el `server.py` es el núcleo funcional mínimo).
