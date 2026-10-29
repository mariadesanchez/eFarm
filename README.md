# eFarm

App estática (un solo `index.html`) para cargar recetas por voz y enviarlas a Google Sheets.

## 1. Apps Script (una vez)
1. Abrí la hoja → Extensiones → Apps Script.
2. Pegá el contenido de `Code.gs` y guardá.
3. Implementar → Nueva implementación → tipo "Aplicación web". Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**.
4. Autorizá los permisos y copiá la URL que termina en `/exec`.
5. Pegala en `index.html`, en `var APPS_SCRIPT_URL=''`.

## 2. GitHub + Vercel
1. Subí esta carpeta a un repositorio nuevo.
2. En Vercel: Add New → Project → importá el repositorio. Framework: Other, sin build. Deploy.
3. Abrí la URL https://… del proyecto (el micrófono solo funciona con https).

Cada envío agrega las filas debajo de la última fila de la pestaña `datos`.

Columnas de la hoja (igual que el listado de psicotrópicos): `Nº | Mat. Médico | Ape. Médico | Fecha | Cant. | Proveedor | Producto`.
El formulario no cambia; `Code.gs` acomoda cada dato en su columna. Si la hoja tiene el orden viejo, la primera vez que se usa la reordena sola y deja una copia en la pestaña `datos_respaldo_…`.

## Subir a GitHub
```
git init
git add .
git commit -m "eFarm: carga por voz a Google Sheets"
git branch -M main
git remote add origin https://github.com/mariadesanchez/eFarm.git
git push -u origin main
```

## Actualizar el script después de cambiar Code.gs
En Apps Script: Implementar → Administrar implementaciones → lápiz (Editar) → Versión: "Nueva versión" → Implementar.
La URL `/exec` no cambia.
