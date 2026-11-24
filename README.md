# Segunda Vuelta Regional - Prerregistro Módulo Especializado

Formulario de prerregistro para el módulo especializado de la Segunda Elección Regional (6 de diciembre de 2026).

Esta carpeta es una copia independiente del formulario de ERM2026 (primera vuelta), adaptada para este nuevo proceso. No comparte Google Sheet ni carpeta de Drive con la primera vuelta.

## Qué cambió respecto a ERM2026 (primera vuelta)

- Banner y título actualizados a "Segunda Vuelta Regional".
- Registro abierto desde ya, cierra el 15 de diciembre de 2026 (23:59, hora Perú).
- Ahora se permite **más de una persona registrada por la misma organización o institución** (antes, una institución que ya tenía un registro no podía volver a registrarse). Esto se cambió quitando la validación de duplicado en `Code.gs` (función `manejarCheckInstitution`).
- **Catálogos reconstruidos a partir de `Base_habilitacion_ser_2026.xlsx`** que compartiste: `instituciones_publicas.json`, `misiones_observacion.json`, `partidos_nacionales.json` y `movimientos_regionales.json` ahora reflejan solo las entidades de esa base (lista reducida para esta segunda vuelta).
- **"Misión de observación / veedor"** ahora es un desplegable con catálogo (antes era texto libre + país). Se quitó el campo de país.
- **Se quitó la categoría "Encuestadoras vigentes"** del formulario (no había base de datos para esta segunda vuelta).
- **Opción "OTRO"** agregada a los cuatro desplegables de entidad (institución pública, partido, movimiento regional, misión de observación): si la entidad no está en la lista, se elige "OTRO" y aparece un campo de texto para escribirla a mano. El tipo de entidad (el radio button) sigue siendo obligatorio escogerlo igual.
- Acepta PDF, JPG o PNG como documento firmado (antes solo PDF).
- **Token básico de seguridad (`APP_TOKEN`)**: tanto `index.html` como `Code.gs` ahora incluyen una misma clave compartida (`APP_TOKEN`) que se envía con cada solicitud y se valida en `doPost`. Esto no es autenticación real (cualquiera puede verla en el código fuente de la página), pero bloquea automáticamente la mayoría de solicitudes de bots/scripts que no pasen por el formulario. Si alguna vez quieres cambiarla, debe coincidir exactamente en ambos archivos.
- `API_URL` (en `index.html`) y `DRIVE_FOLDER_ID` y `APP_TOKEN` (en `Code.gs`) están en blanco o con un valor de ejemplo como placeholders — hay que crear una Google Sheet y una carpeta de Drive nuevas para este proceso y completarlos (ver instrucciones dentro de `Code.gs`).

## Pendiente antes de publicar

1. Crear una Google Sheet nueva para la Segunda Vuelta Regional.
2. Pegar el contenido de `Code.gs` en el editor de Apps Script de esa Sheet.
3. (Opcional pero recomendado) Cambiar el valor de `APP_TOKEN` en `Code.gs` por uno propio, y copiar ese mismo valor en la constante `APP_TOKEN` de `index.html` — deben ser idénticos.
4. Crear (o reutilizar) una carpeta de Drive para los documentos firmados, y poner su ID en `DRIVE_FOLDER_ID`.
5. Implementar como aplicación web (Ejecutar como: Yo / Acceso: Cualquier usuario) y copiar la URL `/exec`.
6. Pegar esa URL en la constante `API_URL` de `index.html`.
7. Si esto va a publicarse en GitHub Pages, hay que crear el repositorio nuevo y subir estos archivos (por ahora esta carpeta es solo local).
