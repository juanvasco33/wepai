# CHANGELOG v15.16.3 — Mejoras estéticas y de UX en la interfaz

Fecha: 28 de mayo de 2026
Módulo: ui/chat_window.py
Tipo: mejoras de experiencia de usuario (no afecta generación ni verificadores)

## Cambios aplicados

### Mejora 1 — Ejemplos "Auto" en la bienvenida
Se añadió la clave "auto" a EXAMPLES_ES/EN con una mezcla de los documentos más
comunes (contrato, cotización, pitch, nómina). Ahora la pantalla de bienvenida
ofrece ejemplos clicables incluso sin tipo de documento preseleccionado.

### Mejora 2 — Selector de formato como filtro opcional
El selector pasó de "Crear:" (instrucción) a "Formato (opcional):" con una nueva
opción ✨ Auto por defecto, donde la IA decide el tipo por el texto. Elimina la
ambigüedad entre los botones y el chat. selected_doc_type ahora se inicializa en
"auto" (antes no se inicializaba — bug latente corregido).

### Mejora 3+5 — Barra lateral muestra valor, no dato técnico
Se reemplazó el bloque "Claude Sonnet / Modelo activo" (irrelevante para el
usuario y además decía Sonnet cuando se usa Haiku) por el contador de
"Documentos este mes" + plan del usuario.

### Mejora 4 — Sello de verificación legal (diferenciador clave)
Nuevo método _verif_badge(): la tarjeta de cada documento generado muestra un
sello verde "✓ Verificado con leyes vigentes de {país}" con los marcadores
legales usados (ej. LFT · IVA · SAT · IMSS) y el mes. Si el overlay persistente
fue aplicado, añade "· actualizado por IA". Hace visible el trabajo de
localización y auto-actualización que antes era invisible para el usuario.

## Verificación
- ui/chat_window.py compila e importa sin errores (validado con tkinter+xvfb).
- App instanciada y renderizada en entorno virtual: las 5 mejoras se ven OK.
- Lógica del sello probada en MX/CO/PE/US: marcadores legales correctos.

## Notas
- Se detectaron 2 bugs latentes del código original (no introducidos aquí):
  self._prefs nunca se asigna en __init__, y self.state="idle" pisa el método
  state() de customtkinter. No afectan el uso normal pero conviene corregirlos.
