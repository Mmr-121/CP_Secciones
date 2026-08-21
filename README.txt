========================================================================
LEEME: GUÍA DE ACTUALIZACIÓN ANUAL - RENTAS Y CÓDIGOS POSTALES ALICANTE
========================================================================

Este proyecto consta de dos scripts de Python diseñados para ejecutarse de forma secuencial cada vez que el INE (Instituto Nacional de Estadística) publique nuevos datos.

1. PREPARACIÓN DE CARPETAS Y ARCHIVOS
------------------------------------------------------------------------
Antes de ejecutar nada, asegúrate de tener en esta misma carpeta:
* Seccion_Urba_mapea.py
* Renta_por_CP.py
* El archivo de rentas anuales del INE (ej: 30833.xlsx).
* La carpeta descomprimida del callejero del INE (ej: caj_esp_072026) que contiene los ficheros TRAM y PSEU.

Nota: Si los nombres de los ficheros descargados del INE cambian el año que viene, debes abrir los archivos .py con un bloc de notas y actualizar el nombre en las variables "FICHERO_TRAM", "FICHERO_PSEU" y "ARCHIVO_RENTA".

2. EJECUCIÓN (PASO A PASO)
------------------------------------------------------------------------
PASO 1: Ejecutar el mapeo base.
Abre tu terminal o consola de comandos y ejecuta:
> python Seccion_Urba_mapea.py

Resultado: Se generará un archivo llamado "mapeo_cusec_cp_urba.xlsx". 
(No lo abras ni lo modifiques, el siguiente script lo necesita cerrado).

PASO 2: Cruzar los datos y generar el Excel final.
En la misma consola, ejecuta:
> python Renta_por_CP.py

Resultado: Se generará un archivo llamado "Datos_Exportar_Plantilla.xlsx".

3. INCORPORACIÓN A TU PLANTILLA
------------------------------------------------------------------------
Abre "Datos_Exportar_Plantilla.xlsx". Copia los datos ordenados y pégalos directamente en tu archivo "Plantilla_Profesional_IPJ_Alicante_2.xlsx" para tener todo el cuadro de mandos y el ranking actualizados.