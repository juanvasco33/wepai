# CHANGELOG v16.9 — Corrección jurídica del contenido laboral (Ley 2466)

Fecha: 30 de mayo de 2026
Módulos: office/legal_config.py, agent/labor_reform_diagnostic.py,
tests/test_labor_reform_diagnostic.py

## Resumen

Tras una revisión de exactitud jurídica contra fuentes oficiales (texto de la
Ley 2466 en Función Pública), se corrigieron 3 errores en el contenido laboral
introducido en v16.7–16.8. Causa raíz común: se trataban como valores fijos
datos que la ley define escalonados en el tiempo. La corrección los hace
calcularse según la fecha vigente.

## Correcciones

### 1. Jornada nocturna — eliminada falsa excepción de microempresa (GRAVE)
La jornada nocturna a las 7:00 p.m. aplica a TODAS las empresas sin excepción por
tamaño (Art. 11 Ley 2466, vigente desde 25-dic-2025). Se eliminó la excepción
incorrecta que eximía a microempresas. La regla ahora alerta a cualquier tamaño.

### 2. Recargo dominical — escalonamiento exacto por fecha (GRAVE)
Implementado el cronograma del Art. 14: 80% (jul-2025–jun-2026), 90%
(jul-2026–jun-2027), 100% (desde jul-2027). Nueva función
`recargo_dominical_vigente(fecha)`. La regla compara contra el % vigente hoy
(80% en may-2026), no contra un genérico 100% (que daba falsos positivos).

### 3. Jornada semanal — transición 44h→42h por fecha (MENOR)
44h hasta el 14-jul-2026, 42h desde el 15-jul-2026 (Ley 2101/2021). Nueva función
`jornada_maxima_vigente(fecha)`. Elimina los falsos positivos para empresas con
44h, que hoy cumplen.

### Extra — procedimiento disciplinario en empresas pequeñas
Añadido el régimen simplificado para empresas con menos de 10 empleados (basta
escuchar previamente al trabajador).

## Config legal (office/legal_config.py, bloque CO)

- `jornada`, `jornada_nocturna_inicio`, `recargo_dominical` y
  `procedimiento_disciplinario` reescritos con los datos verificados y fechas.

## Pruebas

- Eliminado el test que codificaba el bug de la excepción de microempresa;
  reemplazado por uno que verifica el comportamiento correcto.
- Añadidos 6 tests de la lógica escalonada por fecha (valores vigentes en
  distintas fechas, ausencia de falsos positivos en el % vigente, régimen
  simplificado en empresas pequeñas).
- Verificación funcional con fecha real (may-2026): recargo vigente 80%, jornada
  44h, microempresa recibe alerta nocturna.
- Suite total: 363/363 ✅, sin regresiones.

## Nota de arquitectura

Las funciones por fecha convierten dos cambios futuros (jul-2026, jul-2027) en
valores que el sistema resuelve solo, sin requerir intervención de un
desarrollador en esas fechas.
