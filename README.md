# Hub de Herramientas · OGAME

Portal web de la Oficina General de Articulación, Monitoreo y Evaluación de Impacto (OGAME) del MTC. Reúne en un solo lugar las herramientas de seguimiento de las intervenciones del sector.

Se publica con GitHub Pages en: <https://omei-mtc.github.io/HubHerramientas/>

## Cómo funciona

Las páginas son archivos HTML que no guardan datos: los leen de hojas de cálculo de Google publicadas en la web. Cuando una página necesita guardar información (por ejemplo, un registro de avance), la envía a un Apps Script, que la escribe en la hoja correspondiente. Los datos y los scripts viven en la cuenta de Google de OGAME; este repositorio solo guarda las páginas y una copia de referencia de los scripts.

## Estructura

```
├── index.html                 Portada del Hub (menú de herramientas, indicadores y publicaciones)
├── hoja-de-ruta.html          Panel de seguimiento de hitos de la Hoja de Ruta
├── intervenciones-mtc.html    Panel integral de intervenciones del MTC
├── departamentales.html       Carpetas departamentales (Ayuda Memoria y listados)
├── puentes.html               Puentes del MTC (en prueba, aún no figura en el menú)
├── assets/
│   ├── img/                   Logo y foto de portada
│   └── iconos/                Íconos de los indicadores de la portada
├── backend/                   Copia de referencia de los Apps Script
└── archivo/                   Páginas retiradas e imágenes sin uso (no borrar sin revisar)
```

Las páginas activas están en la raíz a propósito: su dirección web es la que usan las oficinas y la hoja de configuración del Hub. **No moverlas ni renombrarlas**, porque se romperían esos enlaces.

## De dónde lee cada página

| Página | Fuente de datos | Escribe en |
| --- | --- | --- |
| `index.html` | Google Sheet "Configuración_Hub_OMEI" (pestañas Herramientas, KPIS y Publicaciones) | — |
| `hoja-de-ruta.html` | Google Sheet "Hoja de Ruta - Intervenciones + Normativa" | Apps Script de la Hoja de Ruta |
| `intervenciones-mtc.html` | Google Sheet "Carpeta_Departamental" (pestaña Actual) y la Hoja de Ruta | — |
| `departamentales.html` | Google Sheets "Carpeta_Departamental" y "Datos_departamentales" | — |
| `puentes.html` | Google Sheet "Puentes – Consolidado" | — (se actualiza con su Apps Script) |

Las direcciones de cada hoja están al inicio del código de cada página, en un bloque de configuración.

## Apps Script (carpeta `backend/`)

Los archivos de esta carpeta son **copias de referencia**. El código que realmente se ejecuta está en los proyectos de Apps Script de la cuenta de OGAME. Cuando se modifique un script, conviene actualizar también su copia aquí.

- `HojaDeRuta_backend.gs`: recibe los registros y validaciones de la Hoja de Ruta. Después de cambiarlo hay que publicar una nueva versión: **Implementar → Gestionar implementaciones → lápiz → Nueva versión**. No usar "Nueva implementación", porque cambia la dirección y la página deja de conectarse.
- `Puentes_Consolidado.gs`: está dentro del Sheet "Puentes – Consolidado". Lee los Excel que los órganos dejan en la carpeta de Drive "Puentes – Fuentes" (subcarpetas PVN, DGPPT y PVD) y se ejecuta desde el menú **Puentes → Actualizar desde Drive**. No requiere implementación.

## Íconos de la portada

La pestaña KPIS de la hoja de configuración indica el ícono de cada indicador solo con el nombre del archivo (por ejemplo, `hito.png`). La portada lo busca dentro de `assets/iconos/`. Para usar un ícono nuevo, basta con subirlo a esa carpeta y escribir su nombre en la hoja.

## Cómo publicar un cambio

1. Subir el archivo modificado con **Add file → Upload files**, con el mismo nombre y en la misma carpeta.
2. Revisar en **Actions** que la tarea "pages build and deployment" termine con un check verde.
3. Abrir la página con Ctrl+F5 para no ver la versión guardada en caché.

Para mover o reorganizar archivos, usar el editor web de GitHub: presionar la tecla punto (`.`) en la página del repositorio.
