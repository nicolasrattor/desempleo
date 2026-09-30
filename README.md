# desempleo

Caracterización de las personas desocupadas de la Encuesta Nacional de Empleo (ENE, INE Chile) en el trimestre móvil junio-julio-agosto 2026, con un módulo de panel que enlaza ese trimestre con JJA 2025, y un dashboard con la serie histórica 2010-2026.

**Dashboard:** [Desempleo en Chile, más allá del 9,6%](https://nicolasrattor.github.io/desempleo/)

## Datos

Los microdatos de la ENE **no se versionan**: pesan alrededor de 1,2 GB y son de descarga pública desde el INE (<https://www.ine.gob.cl/estadisticas/sociales/mercado-laboral/ocupacion-y-desocupacion>). Los scripts los esperan en `data/`, conservando la partición original `ano=<año>/mes_central=<mes>/part-0.parquet`. Si están en otro lugar, basta con apuntar la variable de entorno `ENE_DATA` a esa carpeta:

```bash
export ENE_DATA=/ruta/a/los/microdatos
```

## Cómo reproducirlo

```bash
# 1. tabulados -> desocupados_jja2026.xlsx
ENE_ANIO=2026 ENE_MES=7 python3 build_xlsx.py

# 2. dashboard -> dashboard_desempleo.html
python3 build_dashboard.py
```

`ENE_MES` es el mes central del trimestre móvil (1 = DEF, 6 = MJJ, 7 = JJA…); sin esas variables `build_xlsx.py` usa el trimestre por defecto que trae escrito. `build_dashboard.py` toma solo el `desocupados_*.xlsx` más reciente de la carpeta y deriva de él todos los titulares y textos, así que no hay que editarlo al cambiar de trimestre; la única constante que sí se ajusta a mano es `FECHA_PUB`, la fecha en que el INE publicó el trimestre.

R necesita `arrow`, `dplyr` y `openxlsx`; Python necesita `pandas`, `pyarrow` y `openpyxl`. Todos los comandos se ejecutan desde la raíz del repositorio.

## Contenido del repositorio

`desocupados_mjj2026.R` es el script de referencia, escrito en R con `arrow`, `dplyr` y `openxlsx`. Quedó fijado en el trimestre MJJ 2026, que fue el corte para el que se escribió; documenta la lógica de tabulación pero ya no es el que produce el archivo vigente.

`desocupados_jja2026.xlsx` es el producto vigente (se conserva también `desocupados_mjj2026.xlsx`, el corte anterior): 24 hojas con todos los tabulados, con los porcentajes y las tasas escritos como fórmulas vivas para que recalculen si se edita algún insumo.

`build_xlsx.py` es un espejo en Python de la misma lógica. Existe porque el entorno donde se generó el archivo no tenía R instalado, y es lo que efectivamente construye los `.xlsx` versionados aquí. Está parametrizado por trimestre —`ENE_ANIO` y `ENE_MES` fijan el corte, el enlace panel con la ola de doce meses antes y todas las glosas del archivo—, de modo que actualizar a un trimestre nuevo no requiere tocar el código.

## Universo y definiciones

Los datos provienen de la partición `ano=2026/mes_central=7`, que corresponde al trimestre móvil junio-julio-agosto 2026 (julio es el mes central). La muestra tiene 97.529 personas.

Se considera persona desocupada a quien tiene `activ == 2`: sin empleo, que buscó trabajo en las últimas cuatro semanas y está disponible para trabajar, según el marco conceptual de la OIT adoptado en la 19ª CIET. Dentro de ese grupo, `cae_general == 4` identifica a las personas cesantes (con experiencia laboral previa) y `cae_general == 5` a quienes buscan trabajo por primera vez. La fuerza de trabajo es `ft == 1` y la tasa de desocupación es el cociente entre ambas, multiplicado por cien.

Todas las cifras de personas están expandidas con el factor trimestral `fact_cal`, construido sobre proyecciones de población base Censo 2017. Las columnas de casos muestrales se reportan siempre junto a las expandidas, porque en las desagregaciones más finas el número de observaciones es pequeño: conviene ser cauto con cualquier categoría bajo 50 casos.

## Resultados principales del corte transversal

En JJA 2026 hay 989.400 personas desocupadas (4.415 casos muestrales) sobre una fuerza de trabajo de 10.263.145, lo que da una tasa de desocupación de 9,64%. De ellas, 901.354 son cesantes (91,1%) y 88.046 buscan trabajo por primera vez (8,9%). Estas cifras coinciden con las publicadas por el INE para el mismo trimestre: 9,6% nacional, 9,1% en hombres y 10,3% en mujeres.

Entre las personas cesantes, el motivo de término del último empleo más frecuente es el fin del contrato, proyecto, faena, temporada o reemplazo (57,4%), seguido del despido (21,0%) y la renuncia (14,2%). Dentro de los despidos, la reducción de personal explica tres cuartos de los casos (75,6%). Dentro de las renuncias el cuadro es mucho más repartido: cuidado de niños o de personas enfermas (19,2%), otra razón (17,2%), cansancio del trabajo (14,1%) y motivos de salud (12,8%).

## Módulo panel JJA 2025 – JJA 2026

La ENE es un panel rotativo: cada persona seleccionada se entrevista durante tres meses consecutivos, sale de la muestra y vuelve por tres meses más, doce meses después de la primera visita. El identificador de persona `idrph` permite enlazar ambas olas. El enlace se hace siempre contra el mismo mes calendario (junio con junio, julio con julio, agosto con agosto), de modo que la distancia entre observaciones es exactamente de doce meses, y se valida exigiendo mismo sexo y una diferencia de edad de cero a dos años entre las dos olas. Con ese criterio quedan 24.955 pares válidos, de los cuales 1.032 corresponden a personas que están desocupadas en 2026: un 23,4% del total de desocupados del trimestre.

La respuesta a la pregunta de dónde vienen: un 44,6% de quienes hoy están desocupados estaba ocupado un año antes, un 23,1% ya estaba desocupado (casi todos cesantes) y un 32,3% estaba fuera de la fuerza de trabajo, repartido entre 14,4% que ya formaba parte de la fuerza de trabajo potencial y 17,9% que estaba plenamente inactivo. Es decir, menos de la mitad de la desocupación actual proviene de una pérdida reciente de empleo; el resto es desocupación persistente o entrada (y reentrada) al mercado laboral.

Entre quienes sí estaban ocupados en 2025 (456 casos), los sectores de origen se concentran en comercio (19,9%), construcción (14,2%), industria manufacturera (11,2%), servicios administrativos y de apoyo (8,2%) y enseñanza (6,7%). Un 68,2% eran asalariados del sector privado y un 21,7% trabajaba por cuenta propia. En cuanto al contrato, un 39,9% tenía contrato escrito indefinido, un 24,1% contrato escrito a plazo fijo o de duración definida, un 11,7% trabajaba sin contrato escrito y un 23,7% no era asalariado, por lo que la pregunta no aplica. El 32,8% tenía un empleo informal, por encima del 26-27% que caracteriza al total de ocupados del país.

La lectura de estas cifras debe hacerse con una advertencia importante: la submuestra panel no tiene diseño muestral ni factores de expansión propios. Las cifras expandidas de las hojas `Panel_*` se calculan con `fact_cal` de 2026 y son solo indicativas; la distribución porcentual es la lectura recomendada, y no corresponde presentarlas como estimaciones oficiales del INE.

## Supuestos y validaciones documentadas

Las etiquetas de `cine11_1d` (nivel educacional, CINE 2011) no venían en los metadatos de los archivos parquet y se validaron empíricamente cruzando esa variable con `nivel` y `termino_nivel`.

Las glosas de `e22` (motivo de término del último empleo, pregunta E19 del cuestionario), `e23` (motivo del despido) y `e24` (motivo de la renuncia) provienen del libro de códigos del INE. Las dos últimas se incorporaron a partir del trimestre JJA 2020, recomendadas por la OIT en el contexto del impacto de la pandemia de COVID-19 sobre la población no ocupada.

El tipo de contrato se construyó combinando `b8` (¿tiene contrato escrito?) y `b9` (duración del contrato o acuerdo). La dirección de los códigos de `b9` se verificó contra `b10`, que solo se pregunta a quienes declaran duración definida: `b9 == 1` es definida o a plazo fijo y `b9 == 2` es indefinida.

La clasificación de `cae_general` 6 a 8 como fuerza de trabajo potencial se validó expandiendo esos códigos en el total de JJA 2025, lo que arroja 986.365 personas, consistente con la cifra que publica el INE para ese agregado.

La duración de la búsqueda y de la cesantía se calcula como la diferencia en meses entre el mes de la entrevista y la fecha declarada en `e6` y `e21` respectivamente; los valores negativos se truncan en cero y los códigos 88, 99 y 9999 se clasifican como no declarado. La variable `e21_tramo` que trae la base se descartó por ser inconsistente con las duraciones calculadas.

La serie de larga duración del dashboard usa una regla distinta, porque la no respuesta del mes crece mucho en el tiempo: en `e21_mes` el código 88 pasa de cerca del 10% a un 69% entre fines de 2023 y fines de 2024, de modo que descartar esos casos deforma la serie. La regla se apoya en el año, que casi siempre viene declarado: si el episodio empezó en el mismo año de la entrevista lleva menos de doce meses, si empezó dos o más años antes lleva doce o más, y si empezó el año anterior se compara el mes cuando está declarado y se imputa la probabilidad `mes_encuesta/12` cuando no lo está. Esa celda ambigua es minoritaria —el 93,4% de los casos de cesantía sin mes declarado tiene una brecha de dos años o más, es decir se resuelve sin imputar— y el resultado es una serie estable. Los porcentajes de esa serie corren por encima de los del tabulado transversal (18,8% frente a 16,1% en búsqueda, 29,3% frente a 21,3% en cesantía para JJA 2026) porque los casos sin mes declarado se concentran en episodios largos y el tabulado los deja fuera. La línea de cesantía parte en JJA 2020 porque `e21` no existe en los archivos anteriores.

Rama de actividad, ocupación, categoría ocupacional y formalidad del empleo anterior no están disponibles para las personas desocupadas en el corte transversal: `b1`, `b13`/`b14_rev4cl_caenes`, `ciso1`, `ciso2`, `ocup_form`, `sector` y `categoria_ocupacion` vienen nulas cuando `activ == 2`. Esa es precisamente la información que recupera el módulo panel a partir de la ola de 2025.

## Hojas del Excel

El archivo abre con `Notas` (fuente, definiciones, variables y supuestos) y `Resumen` (PET, fuerza de trabajo, ocupados, desocupados y tasas). Siguen las hojas de caracterización transversal: `Sexo`, `Edad`, `Region`, `Provincia`, `Educacion`, `Nacionalidad`, `Pueblo_indigena`, `Jefatura_hogar`, `Condicion`, `Duracion_busqueda`, `Duracion_cesantia`, `Motivo_termino`, `Motivo_despido`, `Motivo_renuncia`, `Jornada_buscada` y `Metodos_busqueda`. Cierran las seis hojas del panel: `Panel_situacion_2025`, `Panel_rama_2025`, `Panel_contrato_2025`, `Panel_categoria_2025`, `Panel_formalidad_2025` y `Panel_sector_inst_2025`.

Las hojas que se apoyan en una variable que particiona toda la fuerza de trabajo incluyen además la columna de tasa de desocupación. `Metodos_busqueda` es de respuesta múltiple, por lo que sus porcentajes no suman cien y el denominador es el total de personas desocupadas.

La hoja `Provincia` merece una advertencia aparte. La ENE tiene representatividad regional, no provincial: la muestra no está diseñada para producir estimaciones a ese nivel y en más de la mitad de las provincias el número de casos de personas desocupadas es inferior a cincuenta, lo que deja tasas con un error muestral alto. Son cifras de elaboración propia, útiles como exploración pero no como estimaciones oficiales. Aparecen 52 de las 56 provincias del país: Isla de Pascua y Palena no forman parte de la muestra del trimestre, y otras dos no registran personas desocupadas en ella.

## Dashboard

`dashboard_desempleo.html` es un dashboard autocontenido titulado "Desempleo en Chile, más allá del 9,6%". Se abre en cualquier navegador con doble clic; los datos van embebidos como JSON dentro del propio archivo y lo único que carga desde internet es la librería de gráficos Chart.js.

La portada muestra la evolución conjunta de la tasa de desocupación y del número de personas desocupadas en los 198 trimestres móviles disponibles, desde EFM 2010 hasta JJA 2026, con selectores de rango temporal. Bajo la portada hay una barra de pestañas: el contenido se organiza en cinco pestañas y solo una está visible a la vez, sin índice. Las cuatro primeras recorren los mismos tabulados del Excel —quiénes están desocupados (sexo, edad en tasa y en distribución, región, provincia, educación, nacionalidad, pueblo indígena, jefatura de hogar y condición), cuánto llevan buscando (una serie de larga duración desde 2020 más la duración de la búsqueda y de la cesantía, los métodos y la jornada buscada), por qué terminó su último empleo (motivo de término, de despido y de renuncia) y de dónde vienen según el módulo panel (situación en 2025, rama, contrato, categoría ocupacional, formalidad y sector institucional)— y la quinta es la nota metodológica.

En la primera pestaña hay además una serie desde 2020 sobre las personas con educación profesional universitaria (CINE 2011, `cine11_1d` = 4, que excluye magíster y doctorado): una línea con su peso dentro del total de personas desocupadas y otra con la tasa de desocupación del propio grupo. Las dos juntas permiten distinguir el efecto de que el grupo haya crecido dentro de la fuerza de trabajo del efecto de que se desocupe más. Se calcula directamente desde los microdatos, expandido con `fact_cal`.

La serie de larga duración es un gráfico de dos líneas con el porcentaje de personas desocupadas que llevan doce meses o más buscando trabajo y el porcentaje de cesantes que llevan doce meses o más sin empleo, trimestre a trimestre desde 2020. Se calcula directamente desde los microdatos con la regla de brecha de años descrita más arriba, no desde el Excel.

`build_dashboard.py` es el script que lo construye. Se ejecuta desde la raíz del repositorio con `python3 build_dashboard.py` y necesita `openpyxl`, `pandas` y `pyarrow`. Lee los tabulados desde el `desocupados_*.xlsx` más reciente de la carpeta y recalcula la serie histórica recorriendo todas las particiones `ano=*/mes_central=*`. La serie no usa la variable `ft` porque solo existe en los archivos recientes: la fuerza de trabajo se reconstruye como la suma de personas ocupadas y desocupadas (`activ` 1 y 2), que es exactamente su definición. Los porcentajes que en el Excel son fórmulas se recalculan en Python desde los valores, de modo que el dashboard no depende de que el archivo haya sido abierto por Excel o LibreOffice.

Las cifras del dashboard fueron verificadas una a una contra las hojas del Excel (casos muestrales y personas expandidas idénticos en las 21 hojas), los porcentajes de cada hoja suman cien, y la serie reproduce los valores publicados por el INE en los puntos de control: 9,64% en JJA 2026, 8,56% en JJA 2025 y el máximo de la pandemia, 13,09% en MJJ 2020.

El archivo `index.html` es solo una redirección al dashboard, para que GitHub Pages lo sirva como portada del repositorio.

## Fuente y uso

Elaboración propia a partir de microdatos públicos de la Encuesta Nacional de Empleo del Instituto Nacional de Estadísticas de Chile. Las cifras del corte transversal reproducen las publicadas por el INE para JJA 2026; las del módulo panel son elaboración propia y no constituyen estadísticas oficiales.
