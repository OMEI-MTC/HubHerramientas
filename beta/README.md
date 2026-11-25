# Módulo beta · Registro del corte de Intervenciones MTC

Módulo **en prueba**, paralelo al Hub. No figura en el menú, no modifica ninguna página activa y no reemplaza todavía el proceso vigente (matrices Excel y consolidador). Se integrará al Hub cuando se haya probado.

- Página: `beta/registro-intervenciones.html` → <https://omei-mtc.github.io/HubHerramientas/beta/registro-intervenciones.html>
- Backend: `backend/IntervencionesMTC_backend.gs` (Apps Script del Sheet «Base de Intervenciones MTC»)
- Modo demo: `beta/demo/` (simulador de Sheets, datos de ejemplo y arranque)

## Qué hace

Un solo **universo** de intervenciones (pestaña Universo) es la fuente de datos. La columna **Tipo** (Proyecto, Servicio, Política / Normativa, Supervisión, Acción) define el alcance de cada módulo: Intervenciones MTC = Proyecto o Servicio vigentes; Hoja de Ruta = Hoja de Ruta SI; Cartera Priorizada = Cartera Priorizada SI.

El corte sigue los pasos de la directiva v2, pero dentro del Hub:

| Paso | Quién | En el módulo |
|---|---|---|
| 1. Solicitud | Dirección OMEI o administrador | Abre el corte (fecha de corte y plazo) |
| 2. Reporte | Punto focal de la unidad | Ve solo sus intervenciones con el corte anterior precargado; actualiza el estado situacional y los datos del corte; puede cambiar código/CUI (varios con ＋) o crédito suplementario con motivo; registra altas; solicita bajas y correcciones; envía. También puede descargar y cargar la matriz en Excel |
| 3. Revisión | Especialista OMEI | Valida, o **corrige y valida** con sustento (sin devolver a la unidad); devolver es la excepción. Aprueba altas (COD INTERNO sugerido), bajas y correcciones. Puede validar en lote |
| 4. Financiera | Administrador | Descarga la lista de CUI vigentes y carga el archivo de la consulta masiva MEF (pestaña `base`) con vista previa |
| 5–8. Publicación | Administrador o Dirección | Solo con el 100 % validado, sin solicitudes pendientes y con financiera: respalda Actual como «dd-mm», escribe Actual (criterio único del monto), CUI vigentes y KPIs |

Reglas que aplica el sistema (en la página y otra vez en el servidor): desplegables encadenados (Tipo → tipos de código y bloques; Modalidad → solo sus fases, Anexo 04; Crédito SI → monto del crédito; Hoja de Ruta SI → Cartera Priorizada, Relevancia sectorial e Incluida en), CUI de 7 dígitos, monto obligatorio sin CUI, aviso de nombres parecidos y de CUI compartidos. Los valores de las listas vienen de las pestañas `Cat_*`: para agregar una opción se agrega una fila, sin tocar el código.

Asignaciones: Carpeta Departamental por unidad reportante; Hoja de Ruta y Cartera Priorizada por intervención. Lo que no tiene especialista lo ven todos los especialistas. Cada cambio guarda el historial.

## Probar sin instalar nada (modo demo)

Mientras `BACKEND_URL` esté vacío en la página, esta funciona en **modo demo**: el mismo backend corre en el navegador sobre un simulador de Sheets, con ~30 intervenciones del universo (montos, fases y estados situacionales ficticios). Los datos quedan solo en ese navegador. Usuarios: `admin`, `direccion`, `esp.omei`, `esp.ogame` y `pf.<sigla>` (p. ej. `pf.pvn`); contraseña `demo`. «Reiniciar demo» vuelve al inicio.

## Instalar con datos reales (Google)

1. En la cuenta OGAME, crear un Google Sheet nuevo: **Base de Intervenciones MTC**.
2. **Extensiones → Apps Script**: pegar `backend/IntervencionesMTC_backend.gs`, guardar y recargar el Sheet. Aparece el menú **Intervenciones**.
3. **Intervenciones → 1. Preparar base**: crea las pestañas y los catálogos (Modalidad-Fase del Anexo 04, departamentos, tipos).
4. Pegar como valores, con su fila de encabezados:
   - `Carga_Universo` ← pestaña **IntervencionesMTC** de la Hoja de Ruta (universo).
   - `Carga_Actual` ← pestaña **Actual** del Sheet Carpeta_Departamental (último corte; da estado situacional, fase, monto y órgano ejecutor).
   - `Carga_Maestro` ← pestaña **Maestro_interv** de la Hoja de Ruta (Relevancia sectorial, Incluida en, Responsable OGAME; Cartera = periodicidad SEMANAL).
   - `Carga_Financiera` ← tabla **Financiera** del consolidador (CUI, Costo de Inversión, PIA, PIM, Devengado Anual).
   - `Carga_Credito` ← hoja del **crédito suplementario** (COD INTERNO y/o CUI, Monto asignado). Opcional.
5. **Intervenciones → 2. Importar carga inicial**: muestra un resumen (universo, último corte, CUI, asignaciones y lo pendiente de alinear entre la Hoja de Ruta y el universo). Revisar `Cat_Listas` (tipos de infraestructura y de órgano ejecutor salen de la data) y `Cat_Unidades` (nombre de cada sigla).
6. Llenar **Accesos**: Usuario · Contraseña · Rol (`Punto focal`, `Especialista`, `Administrador`, `Dirección`) · Unidad reportante (sigla, solo para punto focal) · Nombre · Activo. Para que se conserven las asignaciones de la Hoja de Ruta, el usuario del especialista debe ser igual al de «Responsable OGAME» del maestro (p. ej. `C.VIGO`).
7. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como: yo · Quién tiene acceso: cualquier usuario). Copiar la URL `/exec` en `BACKEND_URL` de `beta/registro-intervenciones.html`. Cambios posteriores del script: **Gestionar implementaciones → lápiz → Nueva versión** (no «Nueva implementación»).
8. En la página, como administrador: asignar especialistas (pestaña Asignaciones) y abrir un corte de prueba.

**Seguridad.** No publicar el Sheet completo en la web: la pestaña Accesos tiene las contraseñas, que se verifican solo en el Apps Script. Si se quiere comparar el resultado con el panel del Hub, publicar **solo la pestaña Actual** (Archivo → Compartir → Publicar en la web → elegir «Actual»).

## Pruebas

El backend se probó de punta a punta sobre el simulador (login, reglas, reportes, devolución, corrección con sustento, altas con COD sugerido, bajas, correcciones, asignaciones, financiera, publicación y precarga del corte siguiente) y la página con Playwright en los tres perfiles.
