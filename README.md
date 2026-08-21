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

    Extraído de los microdatos del Callejero/Tramos del INE - (https://www.ine.es/dyngs/DAB/es/index.htm?cid=1390).

        - VIAS*: Fichero de catálogo de vías.
        - TRAM*: Fichero de tramos de vía con CUSEC y Código Postal.

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

Relacion_Secciones_CP.xlsx: Tabla de equivalencia entre Secciones Censales (10 dígitos), Municipios y Códigos Postales.

Informe_Renta_Urbanizacion_por_CP.xlsx: Informe de trabajo detallado con métricas numéricas agregadas por CP.

Plantilla_Profesional_IPJ_Alicante.xlsx: Documento final listo para presentación y análisis estratégico.