# WEP AI v16.9.1 — Notas de despliegue

## 1) Arreglo del registro web ("No se pudo crear la cuenta")

### Qué cambió
- `storage/db.py`: la ruta de la base de datos ahora es configurable por la
  variable de entorno `WEPAI_DB_PATH`. Antes estaba fija en `~/.wepai/wepai.db`,
  lo que en producción (servidor) puede caer en un HOME no escribible y romper el
  registro con un error 500 genérico.
- `web/server.py`:
  - El endpoint `/api/register` ahora devuelve la causa real del fallo en vez del
    mensaje genérico.
  - El endpoint `/api/health` ahora incluye un bloque `"db"` que prueba la
    escritura real a la base de datos.

### Cómo diagnosticar TU servidor
1. Vuelve a desplegar con estos archivos.
2. Abre en el navegador: `https://wepai.maclabsmedellin.com/api/health`
3. Mira el bloque `"db"`:
   - `"write_test": "ok"`  → la BD funciona; el problema es otro (ver punto 4).
   - `"write_test": "failed"` con error tipo `unable to open database file` o
     `readonly database` → es PERMISOS / HOME no escribible.
   - `"dir_writable": false` → el usuario del servidor no puede escribir ahí.

### La solución recomendada (independiente del diagnóstico)
Define una ruta de BD garantizada-escribible mediante variable de entorno antes
de arrancar el servidor. Ejemplo:

    export WEPAI_DB_PATH=/var/lib/wepai/wepai.db
    mkdir -p /var/lib/wepai
    chown <usuario-del-servidor> /var/lib/wepai     # el usuario que corre uvicorn

(Reemplaza la ruta y el usuario según tu hosting.)

### 4) Si write_test sale "ok" pero igual no entra
Entonces es la SESIÓN sobre HTTP. La barra del navegador muestra "Not Secure":
el sitio se sirve sin HTTPS. Si la cookie de sesión está marcada `Secure`, el
navegador la descarta sobre HTTP y la sesión no persiste aunque el registro
funcione. Solución correcta: servir el sitio por HTTPS (certificado TLS, p. ej.
Let's Encrypt) en el subdominio `wepai.maclabsmedellin.com`.

---

## 2) Botón de descarga para Mac

### Estado actual
La sección de descarga YA está en `web/static/index.html`, debajo del formulario
de registro. El botón está DESACTIVADO con la etiqueta "Próximamente" para que no
quede un enlace roto mientras no exista el archivo .dmg.

### Paso A — Empaquetar la app (en una Mac)
El binario de Mac debe generarse EN una Mac. Usa el script incluido:

    chmod +x build_macos.sh
    ./build_macos.sh

Esto produce `dist/WEP-AI.dmg`. (Lee las notas de firma/notarización al final
del script: sin firmar, el usuario verá "desarrollador no identificado".)

### Paso B — Subir el .dmg
Sube `WEP-AI.dmg` a tu servidor, por ejemplo a `web/static/downloads/WEP-AI.dmg`,
de modo que quede accesible en:
`https://wepai.maclabsmedellin.com/downloads/WEP-AI.dmg`

### Paso C — Activar el botón
En `web/static/index.html`, busca el bloque con `id="dlMacBtn"` y reemplázalo:

DESACTIVADO (actual):

    <a class="dl-btn disabled" id="dlMacBtn" role="button" aria-disabled="true">
      <svg ...></svg>
      Descargar para Mac
      <span class="dl-soon">Próximamente</span>
    </a>

ACTIVO (cuando el .dmg ya esté subido):

    <a class="dl-btn" id="dlMacBtn" href="/downloads/WEP-AI.dmg" download>
      <svg ...></svg>
      Descargar para Mac
    </a>

(Conserva el `<svg>...</svg>` tal cual; solo quita la clase `disabled`, los
atributos `role`/`aria-disabled`, añade `href` + `download`, y borra el
`<span class="dl-soon">`.)
