# Reporte de Auditoría Jurídica — WEP AI v16.9

**Proyecto:** WEP AI · Asistente de Office con IA para Colombia
**Versión auditada:** v16.8 → corregida a **v16.9**
**Fecha:** 30 de mayo de 2026
**Alcance:** revisión de exactitud jurídica del contenido laboral (Ley 2466 de 2025) — porcentajes escalonados, umbrales y excepciones por tamaño de empresa — y corrección de los errores detectados.

> **Advertencia:** esta revisión se hizo con criterio jurídico y verificación
> contra fuentes oficiales (texto de la Ley 2466 en Función Pública y análisis de
> firmas y entidades), pero NO sustituye el concepto de un abogado colombiano
> colegiado que responda por él. Es un control de exactitud técnica, no un aval
> profesional formal.

---

## 1. Resumen ejecutivo

Se revisó la exactitud jurídica del contenido laboral introducido en v16.7–16.8.
Se detectaron **tres errores** —dos graves y uno menor— y se corrigieron todos.
La causa raíz común: el código trataba como **valores fijos** datos que la ley
define de forma **escalonada en el tiempo**. La corrección no solo arregla los
datos de hoy, sino que los hace **calcularse según la fecha vigente**, evitando
que vuelvan a quedar obsoletos en julio de 2026 y de 2027.

| Métrica | Resultado |
|---|---|
| Errores jurídicos detectados | 3 (2 graves, 1 menor) |
| Errores corregidos | 3 / 3 ✅ |
| Pruebas (antes / después) | 358 → 363 ✅ |
| Regresiones | 0 |
| Verificación contra fuente oficial | Sí (Función Pública, texto de la Ley 2466) |

---

## 2. Errores detectados y corregidos

### E1 · GRAVE — Falsa excepción de microempresa en jornada nocturna ✅ Corregido
El código eximía a las microempresas del inicio de la jornada nocturna a las
7:00 p.m., dejándolas en las 9:00 p.m. **Falso.** El Art. 11 de la Ley 2466
define el trabajo nocturno de 7:00 p.m. a 6:00 a.m. para TODAS las empresas sin
excepción por tamaño (vigente desde el 25-dic-2025). El error habría hecho que un
microempresario dejara de pagar el recargo nocturno (35%) entre las 7 y 9 p.m.,
acumulando deuda laboral. **Corrección:** eliminada la excepción; la regla ahora
alerta a empresas de cualquier tamaño. (La única excepción real por tamaño está
en el procedimiento disciplinario — ver E-extra.)

### E2 · GRAVE — Recargo dominical impreciso (no escalonado por fecha) ✅ Corregido
El código describía el recargo como "hacia el 100%" y la regla alertaba siempre
que fuera menor a 100%, generando falsos positivos. **Corrección:** se implementó
el escalonamiento exacto del Art. 14 (parágrafo transitorio): 80% (jul-2025 a
jun-2026), 90% (jul-2026 a jun-2027), 100% (desde jul-2027). La regla ahora
compara contra el porcentaje **vigente según la fecha**. Hoy (may-2026) exige 80%.

### E3 · MENOR — Jornada de 42h presentada como plena ✅ Corregido
El código fijaba 42h/semana, pero la reducción de la Ley 2101/2021 es gradual:
44h hasta el 14-jul-2026, 42h desde el 15-jul-2026. La regla generaba falsos
positivos para empresas con 44h (que hoy cumplen). **Corrección:** la regla
compara contra el máximo vigente por fecha (44h hoy, 42h tras jul-2026).

### E-extra · Matiz añadido — Procedimiento disciplinario en empresas pequeñas ✅
Se añadió el régimen simplificado para empresas con menos de 10 empleados (basta
escuchar previamente al trabajador), que faltaba.

---

## 3. Contenido verificado como CORRECTO (no requirió cambios)

- Límite de 4 años al contrato a término fijo, incluidas prórrogas. ✅
- Contrato indefinido como regla general. ✅
- Cuota de aprendices (1 por cada 20, empresas de 15+, monetización 1,5 SMMLV). ✅
- Recargo nocturno del 35% y jornada nocturna desde el 25-dic-2025. ✅
- Exigencia de forma escrita para fijo/obra-labor. ✅

---

## 4. Mejora de arquitectura (criterio, no solo dato)

La corrección introduce dos funciones —`recargo_dominical_vigente()` y
`jornada_maxima_vigente()`— que calculan el valor correcto según la fecha. Esto
convierte tres "cambios estructurales futuros" (jul-2026, jul-2027) en valores
que el sistema resuelve solo, **sin requerir intervención de un desarrollador en
esas fechas**. Es la forma robusta de manejar legislación escalonada y reduce la
carga del mecanismo de alertas estructurales de v16.8 para estos casos.

---

## 5. Pendientes (heredados)

1. **(Legal — crítico)** Aval de un abogado colombiano colegiado antes de publicar.
2. **(Calidad)** Validar la auto-verificación con una `ANTHROPIC_API_KEY` real.
3. **(Seguridad)** Forzar fallo de arranque del servidor si falta `WEPAI_SECRET_KEY`.
4. **(Producto)** Aclarar que la "factura" no es comprobante DIAN.
5. **(Vigilancia)** Confirmar periódicamente las etapas y fechas contra el texto
   vigente; las fechas implementadas reflejan la norma a mayo de 2026.

---

*Auditoría de exactitud jurídica sobre el código entregado en v16.9. Las
correcciones están incluidas en el ZIP. El aval profesional sigue siendo un paso
humano posterior e indispensable.*
