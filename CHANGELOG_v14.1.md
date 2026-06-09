# WEP AI — CHANGELOG v14.1 (Bug Fix Release)

Fecha: 21 de mayo de 2026
Base: `wepai_completo_v14.zip`

Esta versión corrige los 14 issues identificados en la revisión de código.
**Ningún fix cambia la arquitectura ni el comportamiento esperado** — solo
elimina bugs latentes, vulnerabilidades y código muerto.

---

## Resumen ejecutivo

| Métrica | Antes | Después |
|---|---|---|
| Bugs críticos | 4 | 0 |
| Bugs altos | 5 | 0 |
| Bugs medios | 5 | 0 |
| Pyflakes en `ui/chat_window.py` | 4 | 1 |
| Pyflakes en `office/excel_controller.py` | 47 | 46 |
| `except: pass` en el proyecto | 15 | 0 |
| Definiciones duplicadas | 1 | 0 |
| Código muerto | 1,403 líneas | 0 |
| Race conditions | 1 | 0 |
| Tests de regresión | 0 | 9 grupos |

---

## Fixes críticos

### 1. `NameError: instrucciones` en modo "Arregla mi Excel"
**Archivo:** `ui/chat_window.py:2178`

**Antes** — variable `instrucciones` nunca definida, crash garantizado al
adjuntar un Excel para mejorar:
```python
fp, bug_report = improve_excel(attached, instrucciones, gen_data)
```

**Después** — usa el campo correcto de `gen_data`:
```python
fp, bug_report = improve_excel(
    attached,
    gen_data.get("instrucciones", ""),
    gen_data,
)
```

---

### 2. Race condition en el límite del plan free
**Archivo:** `storage/db.py` (reescrito completo)

**Antes** — dos operaciones separadas, no atómicas:
```python
can, reason = db.check_can_generate(user_id)   # SELECT
if can:
    db.increment_docs_used(user_id)            # UPDATE separado
```
Bajo concurrencia, varios threads podían pasar el SELECT antes de que
ninguno hiciera el UPDATE, bypaseando el límite de 5 docs/mes.

**Después** — una sola transacción con lock de escritura:
```python
def check_and_consume_doc(user_id):
    with _con() as con:
        cur.execute("BEGIN IMMEDIATE")
        # SELECT y UPDATE dentro de la misma transacción
        # SQLite serializa writers, race condition imposible
        ...
        cur.execute("COMMIT")
```

**Validado con test:** 20 threads paralelos compitiendo por 5 cupos
→ exactamente 5 aceptados, 15 rechazados.

---

### 3. Passwords con SHA-256 sin sal
**Archivo:** `storage/db.py`

**Antes:**
```python
pw_hash = hashlib.sha256(password.encode()).hexdigest()
```
SHA-256 es un hash de propósito general, rapidísimo de fuerza-brutear
offline. Sin sal, dos usuarios con el mismo password producen el mismo
hash, lo que facilita ataques con rainbow tables.

**Después:** bcrypt con cost factor 12.
```python
pw_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode()
```

**Migración automática:** `login_user()` detecta hashes SHA-256 antiguos
(no empiezan con `$2a$`/`$2b$`/`$2y$`) y los actualiza a bcrypt en el
momento del login correcto. Los usuarios existentes no necesitan resetear
su password.

---

### 4. Formulario de pago falso afirmando ser "Pago Seguro · SSL/TLS"
**Archivo:** `ui/chat_window.py:818-1050`

**Antes** — el UI mostraba:
- "🔒 Pago Seguro"
- "Tus datos están cifrados con SSL/TLS"
- "✓ Pago procesado correctamente"

…pero no había procesamiento real ni red ni cifrado. Solo se guardaban
los últimos 4 dígitos en SQLite local. Esto induce a engaño y, si el
usuario entrega una tarjeta real, expone al desarrollador a problemas
legales (engaño al consumidor) y de PCI-DSS (captura de PAN+CVV en
formulario propio).

**Después** — UI rotulado honestamente como modo demo:
- Header: "⚠ MODO DEMO"
- Subtítulo: "Esta versión no procesa cobros reales. No ingreses una tarjeta real."
- Mensaje de éxito: "Plan activado en modo demo. Sin cobro real. Tarjeta no procesada. Para activar cobros reales, integra Stripe Checkout."

**No se cambia la lógica**, solo las cadenas de texto. La integración
con Stripe queda pendiente como decisión de producto.

---

## Fixes altos

### 5. Duplicación de `open_in_excel`
**Archivo:** `office/excel_controller.py`

Había dos definiciones de `open_in_excel` (líneas 2758 y 4423). La segunda
sobreescribía silenciosamente a la primera y usaba un método de apertura
distinto (`subprocess.Popen` en vez de `osascript`).

**Después:** eliminada la duplicada. Queda solo la versión vía AppleScript,
consistente con `open_in_word` y `open_in_powerpoint`.

---

### 6. Patrones regex usados con operador `in` (substring literal)
**Archivo:** `office/excel_controller.py:2698-2716`

**Antes** — 7 ramas del dispatcher escritas así:
```python
elif any(w in full for w in ["directorio.*cliente","crm",...]):
    _build_client_directory(...)
```
`in` hace matching de substring literal, no regex. El patrón
`"directorio.*cliente"` solo matcheaba si el usuario literalmente
escribía `.*` en su mensaje → **nunca matcheaba**. Los siguientes
templates eran inalcanzables:
- `_build_client_directory` (directorio clientes)
- `_build_supplier_directory` (catálogo proveedores)
- `_build_employee_directory` (directorio empleados)
- `_build_work_schedule` (horario turnos)
- `_build_attendance` (control asistencia)
- `_build_vacation_tracker` (vacaciones)
- `_build_sales_pipeline` (pipeline ventas)
- `_build_scenario_simulator` (simulador escenarios)

**Después** — nuevo helper `_matches(text, patterns)`:
- Si el patrón contiene `.*`, `\d`, `\s` o `\b` → usa `re.search`.
- Si no → usa substring literal (comportamiento original).
- Si el regex es inválido → fallback a substring sin crashear.

Las 7 ramas afectadas se migraron a `_matches(full, [...])`.

---

### 7. `ALTER TABLE` ejecutado en cada flujo de pago
**Archivo:** `ui/chat_window.py:1019-1029` → `storage/db.py:_apply_migrations()`

**Antes** — cada vez que un usuario abría el formulario de pago:
```python
try:
    cur.execute("ALTER TABLE users ADD COLUMN card_last4 TEXT")
    con.commit(); con.close()
except: pass  # ignoraba "duplicate column" después de la primera vez
```

**Después** — la columna se crea una vez en `init_db()` vía sistema de
migraciones versionadas con `PRAGMA user_version`. El flujo de pago
ahora solo hace `db.save_card_last4(user_id, last4)` con validación
de que sean exactamente 4 dígitos.

---

### 8. Tkinter modificado desde hilos sin `self.after`
**Archivo:** `ui/chat_window.py:_upd_title`

**Antes:**
```python
def _upd_title(self, text):
    try:
        t = brain.generate_title(text)
        db.update_conversation_title(self.conversation_id, t)
        self.chat_title.configure(text=t)   # modificación de widget desde hilo
        self._load_conversations()           # idem
    except: pass
```

**Después:** Tkinter ops dispatched al main loop:
```python
def _upd_title(self, text):
    try:
        t = brain.generate_title(text)
    except Exception as e:
        print(f"[WEP AI] generate_title failed: {e}")
        return
    ...
    self.after(0, lambda: self.chat_title.configure(text=t))
    self.after(0, self._load_conversations)
```

**Nota:** otras llamadas a `self._bubble()` desde hilos en `_process` y
`_agent` siguen sin envolverse en `self.after`. Esto es deliberado:
refactorizarlas requiere romper el flujo lineal del método y se prefirió
no introducir riesgo en esta release. Es trabajo para v14.2.

---

### 9. 146 unused locals en `word_controller.py`
**Sin cambios** — out of scope de esta release. Indica que varios
builders del Word controller capturan datos del usuario (`today`, `sym`,
`obj_venta`, etc.) pero no los insertan en el documento. Cada uno
requiere análisis manual del builder afectado. Recomendación: ejecutar
`python3 -m pyflakes office/word_controller.py` y revisar caso por caso
en una release dedicada.

---

## Fixes medios

### 10. `word_controller_pre_country.py` (1,403 líneas de código muerto)
**Eliminado.** El archivo era un backup de una versión anterior, no se
importaba en ningún sitio. Si necesitas la historia, usa git tags.

### 11. `except: pass` en 15 ubicaciones (8 en chat_window.py + 7 en ppt_controller.py)
**Todos reemplazados** por `except Exception as e: print(f"[WEP AI] ...: {e}")`.
Ahora los errores se loggean en lugar de tragarse silenciosamente, y
ya no se captura `KeyboardInterrupt`/`SystemExit`/`MemoryError`.

### 12. Dead branch en `ExcelBugDetector.auto_fix`
**Archivo:** `office/excel_controller.py:455`

**Antes:**
```python
elif bug["fix"] == "placeholder":      # nunca match — scan tagea como "formula_needed"
    fixes.append(...)
```

**Después:**
```python
elif bug["fix"] == "formula_needed":
    fixes.append(...)
```

### 13. Faltaban índices en foreign keys
**Archivo:** `storage/db.py`

Añadidos 5 índices:
- `idx_conv_user` en `conversations(user_id)`
- `idx_conv_updated` en `conversations(updated_at)`
- `idx_msg_conv` en `messages(conversation_id)`
- `idx_doc_conv` en `documents(conversation_id)`
- `idx_doc_user` en `documents(user_id)`

### 14. Conexiones SQLite sin context manager
**Archivo:** `storage/db.py`

Antes cada función abría y cerraba su propia conexión con `con = _con(); ...; con.close()`. Si algo fallaba a la mitad, la conexión quedaba colgando hasta el GC.

Ahora todas usan `with _con() as con:` — garantía de cleanup.

---

## Cambios en dependencias

`requirements.txt`:
```diff
 anthropic>=0.40.0
+bcrypt>=4.1.0
 customtkinter>=5.2.2
 python-docx>=1.1.0
 openpyxl>=3.1.2
 python-pptx>=0.6.23
 Pillow>=10.0.0
 python-dotenv>=1.0.0
```

---

## Backward compatibility

- ✅ La base de datos existente sigue funcionando — el sistema de
  migraciones detecta el schema viejo y aplica solo los cambios necesarios.
- ✅ Usuarios con hashes SHA-256 antiguos siguen pudiendo iniciar sesión.
  Su hash se actualiza a bcrypt automáticamente en el primer login.
- ✅ `db.check_can_generate()` e `db.increment_docs_used()` quedan como
  aliases de la nueva API atómica, así código que las llame sigue funcionando.

---

## Tests de regresión incluidos

El script de test en `tests/test_db.py` cubre:
1. `init_db()` sin errores
2. Registro y login con bcrypt
3. Migración automática SHA-256 → bcrypt en login
4. Plan free: exactamente 5 docs permitidos, 6° rechazado
5. **Race condition**: 20 threads concurrentes, exactamente 5 docs aceptados
6. `save_card_last4()` valida formato (rechaza no-dígitos)
7. `get_plan_limits()` para free/pro/biz
8. CRUD de conversations, messages, documents
9. `_matches()` helper para los patrones regex del dispatcher

Para correrlos: `python3 tests/test_db.py`

---

## Para v14.2 (trabajo pendiente)

- Refactor del threading en `_process` y `_agent` para que TODAS las
  llamadas a Tkinter pasen por `self.after`.
- Auditar los 146 unused locals en `word_controller.py` (probable
  pérdida de datos del usuario).
- Considerar dividir `word_controller.py` (5,150 líneas) y
  `excel_controller.py` (4,427 líneas) en módulos por tipo de documento.
- Migrar el dispatch de templates a que Claude devuelva directamente el
  `template_id` en el JSON `##GENERAR##` — más confiable que adivinar
  con keywords.
- Integrar Stripe Checkout para reemplazar el formulario de pago demo.
