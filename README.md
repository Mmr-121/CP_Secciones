# CP_Secciones

readme_content = """# Pipeline de Análisis de Renta y Urbanizaciones por Código Postal (Alicante)

Este proyecto proporciona un pipeline automatizado en Python para procesar, agrupar y analizar datos socioeconómicos y urbanísticos de la provincia de Alicante (Código INE `03`) a nivel de **Código Postal (CP)**. 

A partir de los microdatos oficiales del **Instituto Nacional de Estadística (INE)** (secciones censales, tramos de vía y datos de renta por hogar), el sistema calcula promedios ponderados de renta, tasas de evolución interanual, porcentaje de edificación tipo urbanización y niveles de fiabilidad estadística, volcando finalmente los resultados en un ranking profesional listo para decisiones de segmentación, marketing (SEO/Ads) e inversión.

---

## Tabla de Contenidos

1. [Características Principales](#-características-principales)
2. [Estructura del Proyecto](#-estructura-del-proyecto)
3. [Requisitos y Dependencias](#-requisitos-y-dependencias)
4. [Estructura de Datos de Entrada](#-estructura-de-datos-de-entrada)
5. [Flujo de Trabajo y Scripts](#-flujo-de-trabajo-y-scripts)
6. [Metodología de Cálculo](#-metodología-de-cálculo)
7. [Guía de Uso](#-guía-de-uso)
8. [Archivos Generados](#-archivos-generados)

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

## Flujo de Trabajo y Pipeline

El script orquestador actualizar_todo.py gestiona la ejecución secuencial del pipeline:

graph TD
    A[actualizar_todo.py] --> B[1. clean.py]
    A --> C[2. generar_relacion_secciones_cp.py]
    A --> D[3. Renta_segun_CP.py]
    A --> E[4. actualizar_ranking_plantilla.py]
    
    B -->|Elimina outputs antiguos| B1[Limpia workspace]
    C -->|Mapea CUSEC ➔ CP| C1[Relacion_Secciones_CP.xlsx]
    D -->|Pondera Renta y Urba| D1[Informe_Renta_Urbanizacion_por_CP.xlsx]
    E -->|Vuelca datos| E1[Plantilla_Profesional_IPJ_Alicante.xlsx]