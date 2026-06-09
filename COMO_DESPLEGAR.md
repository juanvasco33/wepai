# Cómo poner WEP AI a funcionar de verdad (registro + login)

El frontend ya se ve bien, pero los usuarios no pueden crear cuenta porque el
backend Python (server.py) no está corriendo. Netlify NO puede correr Python.
Esta guía lo resuelve desplegando todo en Render.com (gratis).

## Resultado final
Toda la app (frontend + backend) correrá desde una sola URL de Render, p. ej.
`https://wepai.onrender.com`. El registro, login y generación de documentos
funcionarán. Luego apuntas tu dominio `wepai.maclabsmedellin.com` a esa URL.

---

## Pasos

### 1. Sube el proyecto a GitHub
Si ya lo tienes en GitHub, asegúrate de subir esta versión nueva (incluye
`render.yaml`). Si no:
    cd wepai_v1
    git init && git add . && git commit -m "WEP AI v16.9.1"
    git remote add origin https://github.com/TU_USUARIO/wepai.git
    git push -u origin main

### 2. Crea el servicio en Render
1. Entra a https://render.com y crea una cuenta (puedes usar tu GitHub).
2. Clic en **New → Blueprint**.
3. Conecta tu repositorio de WEP AI.
4. Render detecta el archivo `render.yaml` y propone crear el servicio "wepai".
5. Acepta. Empezará a construir (tarda unos minutos la primera vez).

### 3. Agrega tu clave de Anthropic
En el dashboard del servicio → **Environment** → agrega:
    ANTHROPIC_API_KEY = (tu clave de https://console.anthropic.com)
Sin esta clave, el login/registro SÍ funcionan, pero la generación de
documentos con IA no.

### 4. Prueba que funciona
1. Render te da una URL: `https://wepai-xxxx.onrender.com`.
2. Ábrela. Verás tu página de WEP AI.
3. Abre `https://wepai-xxxx.onrender.com/api/health` y confirma que el bloque
   `"db"` diga `"write_test": "ok"`.
4. Crea una cuenta de prueba. Ahora SÍ debe entrar.

### 5. Conecta tu dominio wepai.maclabsmedellin.com
Tienes dos opciones:

**Opción A (recomendada) — apuntar el subdominio a Render:**
En Render → Settings → Custom Domains → agrega `wepai.maclabsmedellin.com`.
Render te dará un registro CNAME. En el panel DNS de maclabsmedellin (donde
manejas el dominio) crea ese CNAME apuntando a la URL de Render. Render emite el
certificado HTTPS automáticamente. Quita la versión vieja de Netlify para este
subdominio para que no haya conflicto.

**Opción B — dejar el frontend en Netlify y solo el backend en Render:**
Más complejo (hay que manejar CORS y apuntar las llamadas /api a otra URL). No
lo recomiendo; la Opción A es más simple porque server.py ya sirve el frontend.

---

## Nota sobre el plan gratis de Render
El plan free "duerme" tras ~15 min sin uso; la primera visita después tarda
~30 segundos en despertar. Para producción real, el plan Starter (US$7/mes)
lo mantiene siempre activo. Para validar y conseguir tus primeros usuarios, el
free sirve.

## Alternativa: Railway.app
Funciona igual de bien. En vez de render.yaml, Railway detecta el proyecto
Python solo. Define las mismas variables (WEPAI_DB_PATH, ANTHROPIC_API_KEY) y
el start command: `cd web && uvicorn server:app --host 0.0.0.0 --port $PORT`.
