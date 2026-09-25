# Sistema de Alerta Temprana (SAT) Triprovincial - Cuenca Río V

Este repositorio contiene el sistema automatizado de monitoreo hidrometeorológico en tiempo real para la Cuenca del Río V (San Luis, Córdoba y La Pampa).

## Estructura de Archivos
* `sat_unificado_hidro_meteo.py`: Script principal que consulta los datos en vivo, aplica el modelo de onda y genera los informes.
* `requirements.txt`: Dependencias de Python necesarias.
* `.github/workflows/actualizar_sat.yml`: Tarea programada (GitHub Actions) que corre cada 30 minutos y actualiza la web.
* `mapa_sat_rio_v_triprovincial.html` / `index.html`: Visor cartográfico Leaflet con capas satelitales, radares y botones de descarga.
* `sat_unificado_rio_v_triprovincial.xlsx`: Informe ejecutivo multisábana descargable directamente desde la web.
* `resumen_cruce_rio_v_triprovincial.csv`: Matriz de cruce hidrológico-meteorológico.

## Puesta en marcha en GitHub (Paso a Paso)
1. Crea un repositorio nuevo en GitHub (ej: `sat-riov-tiempo-real`).
2. Sube todos los archivos de esta carpeta manteniendo la estructura (especialmente la carpeta `.github/workflows/`).
3. Ve a **Settings** > **Pages**:
   - En **Source**, selecciona `Deploy from a branch`.
   - Elige rama `main` y carpeta `/ (root)`. Guarda los cambios.
4. En **Settings** > **Actions** > **General**:
   - Asegúrate de que en **Workflow permissions** esté seleccionado **Read and write permissions**.
5. ¡Listo! El sistema se ejecutará automáticamente cada 30 minutos actualizando los datos en vivo y permitiendo la descarga de informes.
