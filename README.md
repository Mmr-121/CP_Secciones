# CP_Secciones

readme_content = """# Pipeline de Análisis de Renta y Urbanizaciones por Código Postal (Alicante)

Este proyecto proporciona un pipeline automatizado en Python para procesar, agrupar y analizar datos socioeconómicos y urbanísticos de la provincia de Alicante (Código INE `03`) a nivel de **Código Postal (CP)**. 

A partir de los microdatos oficiales del **Instituto Nacional de Estadística (INE)** (secciones censales, tramos de vía y datos de renta por hogar), el sistema calcula promedios ponderados de renta, tasas de evolución interanual, porcentaje de edificación tipo urbanización y niveles de fiabilidad estadística, volcando finalmente los resultados en un ranking profesional listo para decisiones de segmentación, marketing (SEO/Ads) e inversión.

---

## Estructura del Proyecto

```text
.
├── caj_esp_XXXXXX/                        # Carpeta con microdatos INE (VIAS / TRAM)
├── 30833.xlsx                             # Excel oficial del INE con datos de renta por CUSEC
├── Plantilla_Profesional_IPJ_Alicante.xlsx # Plantilla Excel de destino (Hoja: Ranking_CP)
├── clean.py                               # Script de limpieza de salidas anteriores
├── generar_relacion_secciones_cp.py       # Asignador de CUSEC a Código Postal mayoritario
├── Renta_segun_CP.py                      # Agregador de renta y urbanizaciones por CP
├── actualizar_ranking_plantilla.py        # Volcado de datos a la plantilla profesional
├── actualizar_todo.py                     # Orquestador principal del pipeline
└── README.md                              # Documentación del proyecto
```
---

## Datos de Entrada Necesarios
Para ejecutar el pipeline correctamente, asegúrate de situar en la raíz del proyecto los siguientes archivos oficiales:

1. Directorio de Datos del INE (caj_esp_*):

    Extraído de los microdatos del Callejero del Censo Electoral del INE - (https://www.ine.es/dyngs/DAB/es/index.htm?cid=1390).

        - VIAS_*: Fichero de catálogo de vías (necesario).
        - TRAM_*: Fichero de tramos de vía con CUSEC y Código Postal (necesario).
        - SECC_*, PSEU_*, UP_*: vienen en la misma descarga del INE pero el pipeline actual NO los usa. SECC es útil para comprobar a mano la cobertura (cuántas secciones censales oficiales quedan fuera del cruce), pero no hace falta para ejecutar el pipeline.

2. Excel de Renta del INE (30833.xlsx o similar): https://www.ine.es/dynt3/inebase/es/index.htm?padre=12385&capsel=5685

3. Plantilla Excel Profesional (Plantilla_Profesional_IPJ_Alicante.xlsx).
    
---
## Instrucciones de Uso
Coloca los ficheros de datos del INE (caj_esp_* y 30833.xlsx) en el directorio raíz.

Abre una terminal en el directorio del proyecto.

Ejecuta el orquestador principal: python actualizar_todo.py

Revisa los logs en la terminal. Al finalizar, la plantilla Excel se habrá actualizado automáticamente.

---
## Archivos Generados

Relacion_Secciones_CP.xlsx: Tabla de equivalencia entre Secciones Censales (10 dígitos, con municipio/distrito/sección desglosados) y Códigos Postales.

Informe_Renta_Urbanizacion_por_CP.xlsx: Informe de trabajo detallado con métricas numéricas agregadas por CP.

Plantilla_Profesional_IPJ_Alicante.xlsx: Documento final listo para presentación y análisis estratégico.

---
## Limitaciones conocidas

- **Asignación de renta por CP "mayoritario":** el INE publica renta por sección censal, no por código postal, y no existe una tabla oficial de traducción gratuita entre ambos. Cuando una sección tiene calles repartidas entre varios CP, toda su renta se asigna al CP con más tramos de calle. Con la edición de datos usada al validar este pipeline: ~12% de las secciones tienen tramos en más de un CP, y ~4% de los tramos totales de la provincia quedan en la parte "minoritaria" de su sección (pureza media 97,6%). Es una muy buena aproximación para segmentación de marketing (Google Ads ya trabaja con un nivel de precisión de CP similar), pero no una asignación exacta al 100%.
- **% de urbanizaciones:** depende de que el Callejero del INE clasifique bien el tipo de vía (`URB`/`URBAT` sobre el resto de tipos). El cálculo en sí está verificado contra los ficheros reales, pero la calidad de esa clasificación es responsabilidad del INE, no del pipeline.
- **Secreto estadístico:** algunas secciones censales muy pequeñas no tienen renta publicada por el INE (protección de privacidad). El cruce sección→CP no depende del Excel de renta y sí las incluye; el CP resultante puede tener menos cobertura de renta de la que le correspondería. Los CP con menos del 50% de cobertura de renta se marcan como "Fiabilidad_Renta: Baja" y se anotan en la columna Observaciones de `Ranking_CP`.
- **Orden de `Ranking_CP`:** actualmente ordena solo por renta bruta descendente. Combinar renta y % de urbanizaciones (u otras métricas como Habitantes, % Unifamiliares o Competencia) en un único criterio de prioridad queda pendiente de definir con más datos.