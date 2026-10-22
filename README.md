# Incidencias Puntonet

Registro y control de las incidencias de Puntonet con el servicio de Metrored.

- **Panel de control:** https://metroredec.github.io/incidencias/
- **Formulario de queja:** https://metroredec.github.io/incidencias/queja.html

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | Panel: indicadores, gráficos, filtros, edición, exportación a CSV |
| `queja.html` | Formulario para registrar una queja nueva |
| `assets/incidencias.js` | Lógica compartida: almacenamiento, categorías e histórico 2026 |
| `apps-script/incidencias.gs` | Backend en Google Sheets |
| `apps-script/cargar-reporte-jul-oct.gs` | Carga única de las 37 quejas de julio al 21 de octubre de 2026 (instrucciones dentro del archivo) |

## Activar el registro compartido

**En producción:** `ENDPOINT` (en `assets/incidencias.js`) apunta al Apps Script publicado. Si se vacía, las páginas vuelven al **modo demo** (cada navegador guarda sus propias incidencias).

1. Crear una hoja de Google → Extensiones → Apps Script → pegar `apps-script/incidencias.gs`.
2. Cambiar `CLAVE` por una clave propia y compartirla con el equipo de Puntonet.
3. Ejecutar `cargarHistorico()` una vez para cargar las incidencias de 2026.
4. Implementar → Nueva implementación → Aplicación web (ejecutar como: yo · acceso: cualquier persona).
5. Pegar la URL `/exec` en `ENDPOINT` dentro de `assets/incidencias.js`.

Los nombres de pacientes del histórico se reemplazaron por `[NOMBRE RESERVADO]` porque este repositorio es público. Los datos reales viven en la hoja de Google, protegida por la clave.
