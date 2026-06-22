# Evaluación Piloto · Entidades Patrocinadoras (CORFO)

Informe interactivo (D3.js) con los resultados agregados del piloto de evaluación
de entidades patrocinadoras. Es un sitio **estático**: todo corre en el navegador.

## Ver en línea (GitHub Pages)

1. Sube este proyecto a un repositorio de GitHub.
2. En el repo: **Settings → Pages**.
3. En *Build and deployment* → *Source*: **Deploy from a branch**.
4. Branch: **main** (o `master`) y carpeta **/ (root)**. Guardar.
5. A los pocos minutos el sitio queda disponible en
   `https://<usuario>.github.io/<repositorio>/`.

No requiere servidor ni configuración adicional: `index.html` ya trae los datos
incrustados y carga D3 desde `vendor/` (con respaldo por CDN).

## Estructura

- `index.html` — informe completo (HTML + CSS + JS + datos incrustados).
- `vendor/d3.v7.min.js` — librería de gráficos.
- `logo corfo.png` — logo institucional.
- `data/` — copias de los datos en JSON (respaldo; el sitio usa los datos incrustados).
- `procesar_datos.py` — script que regenera los datos desde el Excel.

## Actualizar los datos

1. Edita el Excel `Analisis Piloto Ev Patrocinadores.xlsx` y **ciérralo**
   (mientras está abierto queda bloqueado y el script no puede leerlo).
2. Ejecuta:
   ```
   py procesar_datos.py
   ```
   Esto regenera la carpeta `data/` e **inyecta los datos en `index.html`**.
3. Sube los cambios a GitHub.

> ⚠️ **Privacidad:** el Excel contiene correos electrónicos y está excluido del
> repositorio mediante `.gitignore`. Los archivos publicados (`index.html`, `data/`)
> contienen solo datos agregados y anonimizados. Si el repositorio es **público**,
> no vuelvas a incluir el Excel.
