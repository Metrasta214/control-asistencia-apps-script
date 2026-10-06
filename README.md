# Control de asistencia

Web App para controlar la asistencia de 30 participantes durante 10 sesiones del curso **Inteligencia artificial aplicada a las TIC**.

## Objetivo

Registrar asistencia por sesion usando Google Apps Script como backend y alojamiento de la interfaz, con Google Sheets como almacenamiento. No requiere servidores externos ni otra base de datos.

## Arquitectura

- `Code.gs`: expone `doGet()` y sirve la interfaz `index.html`.
- `Config.gs`: centraliza el ID de Google Sheets, nombres de hojas, encabezados y estados.
- `Setup.gs`: crea y valida las hojas `Participantes`, `Sesiones` y `Asistencias`; carga participantes `P01` a `P30` y sesiones `S01` a `S10` sin duplicar datos.
- `Asistencia.gs`: consulta sesiones, programa fecha/hora y guarda estados de pase de lista actualizando registros existentes o insertando nuevos.
- `index.html`: interfaz responsive para teléfono, tableta y escritorio, con tarjetas por alumno, estados de presente, ausente, retardo, falta justificada y suspensión de clases.
- `appsscript.json`: manifiesto del proyecto Apps Script con V8 y zona horaria `America/Mexico_City`.

## Configuracion

El proyecto usa la hoja de calculo configurada en `Config.gs`:

```javascript
const SPREADSHEET_ID = '1GBuWaHHUpZ4xOI9gIfxR-RRJlCF9fhIn4SjzVZwlqdM';
```

Las hojas esperadas son:

- `Participantes`: `id_participante`, `nombre`, `correo`, `activo` (la preparación inicial carga 30 nombres y correos escolares de ejemplo)
- `Sesiones`: `id_sesion`, `numero`, `fecha`, `tema`, `estado`
- `Asistencias`: `id_sesion`, `id_participante`, `estado`, `hora_registro`

Ejecuta `prepararBaseDatos()` desde Apps Script para inicializar la hoja. La funcion es idempotente: puede ejecutarse varias veces sin borrar datos ni duplicar participantes o sesiones.

## Sincronizacion con clasp

Verifica autenticacion:

```bash
clasp login
```

Descargar cambios remotos:

```bash
clasp pull
```

Subir cambios locales cuando ya hayan sido revisados:

```bash
clasp push
```

Este repositorio ignora `.clasp.json` y `.clasprc.json` para evitar publicar identificadores locales o credenciales.

## Publicacion

En Apps Script:

1. Ejecuta `prepararBaseDatos()`.
2. Revisa que las hojas y encabezados se hayan creado correctamente.
3. Despliega como Web App.
4. Configura permisos de ejecucion segun la practica.
5. Abre la URL de la Web App en telefono, tableta y escritorio.

## Pruebas sugeridas

- Ejecutar `prepararBaseDatos()` dos veces y confirmar que no duplica `P01` a `P30` ni `S01` a `S10`.
- Cambiar encabezados manualmente en una hoja de prueba y confirmar que la validacion rechaza la estructura incorrecta.
- Cargar una sesion desde la Web App y verificar que aparecen 30 participantes activos.
- Marcar algunos participantes, guardar y confirmar que se crean registros en `Asistencias`.
- Cambiar la misma sesion y guardar de nuevo para confirmar que actualiza registros existentes sin borrar la tabla.
- Cambiar un estado a retardo, falta justificada o suspensión y confirmar que se conserva en `Asistencias`.
- Seleccionar una fecha y hora distintas, guardar el horario y comprobar que se conserva al recargar.
- Confirmar que la sesion cambia a `REGISTRADA` tras guardar.
- Abrir dos ventanas, guardar casi al mismo tiempo y verificar que no se generan duplicados.
