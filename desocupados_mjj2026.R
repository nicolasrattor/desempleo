# =============================================================================
# Caracterización de las personas DESOCUPADAS - ENE (INE Chile)
# Trimestre móvil mayo-junio-julio 2026 (MJJ 2026)
#
# Fuente: microdatos ENE en input/data (parquet particionado ano/mes_central)
# Salida: ene_empleo_microdatos/desocupados_mjj2026.xlsx
#
# Incluye un módulo de PANEL: la ENE es un panel rotativo y cada persona (idrph)
# vuelve a ser entrevistada 12 meses después, en el mismo mes calendario. Eso
# permite enlazar MJJ 2026 con MJJ 2025 y observar de dónde vienen las personas
# que hoy están desocupadas.
#
# Definiciones (INE, marco OIT / 19a CIET):
#   activ == 1 -> Ocupado/a ; activ == 2 -> Desocupado/a ; activ == 3 -> Fuera de la FT
#   ft == 1    -> Fuerza de trabajo (ocupados + desocupados)
#   cae_general == 4 -> Cesante ; cae_general == 5 -> Busca trabajo por primera vez
#   Tasa de desocupación = desocupados / fuerza de trabajo * 100
#   Todas las cifras poblacionales usan el factor de expansión trimestral fact_cal.
# =============================================================================

library(arrow)
library(dplyr)
library(tidyr)
library(openxlsx)

ruta_datos  <- Sys.getenv("ENE_DATA", "data")  # microdatos ENE, partición ano=/mes_central=
# Los microdatos no viajan en el repositorio (~1,2 GB): descargarlos del INE y dejarlos en ./data,
# o apuntar ENE_DATA a la carpeta donde estén. Ejecutar desde la raíz del repositorio.
ruta_salida <- "desocupados_mjj2026.xlsx"

ANIO        <- 2026
MES_CENTRAL <- 6   # junio = trimestre móvil mayo-junio-julio
ANIO_PREV   <- ANIO - 1   # ola anterior del panel: MJJ 2025

# ---------------------------------------------------------------------------
# 1. Lectura del trimestre
# ---------------------------------------------------------------------------
ds <- open_dataset(ruta_datos)

ene <- ds %>%
  filter(ano == ANIO, mes_central == MES_CENTRAL) %>%
  collect()

stopifnot(nrow(ene) > 0)

# ---------------------------------------------------------------------------
# 2. Etiquetas
# ---------------------------------------------------------------------------
lab_sexo <- c("1" = "Hombre", "2" = "Mujer")

lab_tramo_edad <- c(
  "1" = "15 a 19 años",  "2" = "20 a 24 años",  "3" = "25 a 29 años",
  "4" = "30 a 34 años",  "5" = "35 a 39 años",  "6" = "40 a 44 años",
  "7" = "45 a 49 años",  "8" = "50 a 54 años",  "9" = "55 a 59 años",
  "10" = "60 a 64 años", "11" = "65 a 69 años", "12" = "70 años o más"
)

lab_region <- c(
  "15" = "Arica y Parinacota", "1" = "Tarapacá",   "2" = "Antofagasta",
  "3" = "Atacama",             "4" = "Coquimbo",   "5" = "Valparaíso",
  "13" = "Metropolitana de Santiago", "6" = "O'Higgins", "7" = "Maule",
  "16" = "Ñuble", "8" = "Biobío", "9" = "La Araucanía", "14" = "Los Ríos",
  "10" = "Los Lagos", "11" = "Aysén", "12" = "Magallanes"
)
orden_region <- c(15, 1, 2, 3, 4, 5, 13, 6, 7, 16, 8, 9, 14, 10, 11, 12)

# Catálogo de provincias (código INE). Isla de Pascua (52) y Palena (104) no están
# en la muestra del trimestre, por lo que no aparecen en el tabulado.
lab_provincia <- c(
  "151" = "Arica (Arica y Parinacota)",
  "152" = "Parinacota (Arica y Parinacota)",
  "11" = "Iquique (Tarapacá)",
  "14" = "Tamarugal (Tarapacá)",
  "21" = "Antofagasta (Antofagasta)",
  "22" = "El Loa (Antofagasta)",
  "23" = "Tocopilla (Antofagasta)",
  "31" = "Copiapó (Atacama)",
  "32" = "Chañaral (Atacama)",
  "33" = "Huasco (Atacama)",
  "41" = "Elqui (Coquimbo)",
  "42" = "Choapa (Coquimbo)",
  "43" = "Limarí (Coquimbo)",
  "51" = "Valparaíso (Valparaíso)",
  "53" = "Los Andes (Valparaíso)",
  "54" = "Petorca (Valparaíso)",
  "55" = "Quillota (Valparaíso)",
  "56" = "San Antonio (Valparaíso)",
  "57" = "San Felipe de Aconcagua (Valparaíso)",
  "58" = "Marga Marga (Valparaíso)",
  "131" = "Santiago (Metropolitana)",
  "132" = "Cordillera (Metropolitana)",
  "133" = "Chacabuco (Metropolitana)",
  "134" = "Maipo (Metropolitana)",
  "135" = "Melipilla (Metropolitana)",
  "136" = "Talagante (Metropolitana)",
  "61" = "Cachapoal (O'Higgins)",
  "62" = "Cardenal Caro (O'Higgins)",
  "63" = "Colchagua (O'Higgins)",
  "71" = "Talca (Maule)",
  "72" = "Cauquenes (Maule)",
  "73" = "Curicó (Maule)",
  "74" = "Linares (Maule)",
  "161" = "Diguillín (Ñuble)",
  "162" = "Itata (Ñuble)",
  "163" = "Punilla (Ñuble)",
  "81" = "Concepción (Biobío)",
  "82" = "Arauco (Biobío)",
  "83" = "Biobío (Biobío)",
  "91" = "Cautín (La Araucanía)",
  "92" = "Malleco (La Araucanía)",
  "141" = "Valdivia (Los Ríos)",
  "142" = "Ranco (Los Ríos)",
  "101" = "Llanquihue (Los Lagos)",
  "102" = "Chiloé (Los Lagos)",
  "103" = "Osorno (Los Lagos)",
  "111" = "Coyhaique (Aysén)",
  "112" = "Aysén (Aysén)",
  "113" = "Capitán Prat (Aysén)",
  "114" = "General Carrera (Aysén)",
  "121" = "Magallanes (Magallanes)",
  "122" = "Antártica Chilena (Magallanes)",
  "123" = "Tierra del Fuego (Magallanes)",
  "124" = "Última Esperanza (Magallanes)"
)
orden_provincia <- c(151, 152, 11, 14, 21, 22, 23, 31, 32, 33, 41, 42, 43, 51, 53, 54, 55, 56, 57, 58, 131, 132, 133, 134, 135, 136, 61, 62, 63, 71, 72, 73, 74, 161, 162, 163, 81, 82, 83, 91, 92, 141, 142, 101, 102, 103, 111, 112, 113, 114, 121, 122, 123, 124)

# cine11_1d = nivel educacional alcanzado (CINE 2011, recodificación INE).
# Etiquetas validadas cruzando cine11_1d con `nivel` y `termino_nivel` de la misma base.
lab_educ <- c(
  "0" = "Sin educación formal o básica incompleta",
  "1" = "Básica completa o media incompleta",
  "2" = "Media completa",
  "3" = "Técnica de nivel superior",
  "4" = "Profesional (universitaria)",
  "5" = "Magíster",
  "6" = "Doctorado",
  "9" = "Nivel no declarado"
)

lab_condicion <- c("4" = "Cesante", "5" = "Busca trabajo por primera vez")

lab_jornada <- c("1" = "Jornada completa", "2" = "Jornada parcial",
                 "3" = "Indistinto", "88" = "No sabe", "99" = "No responde")

# e22 (pregunta E19): "¿Por qué razón ya no tiene ese empleo, negocio o actividad
# por cuenta propia?". Glosas de los códigos 3 a 9, 88 y 99 según el libro de códigos
# del INE. Los códigos 1 y 2 son los que habilitan e23 (motivo de despido) y e24
# (motivo de renuncia).
lab_motivo <- c(
  "1"  = "Despido",
  "2"  = "Renuncia",
  "3"  = "Fin del contrato, proyecto, faena, temporada o reemplazo",
  "4"  = "Jubilación",
  "5"  = "Término del ejercicio de la actividad por cuenta propia",
  "6"  = "Quiebra o cierre de negocio por razones económicas",
  "7"  = "Quiebra o cierre de negocio por un fenómeno natural o siniestro",
  "8"  = "Quiebra o cierre de negocio por la pandemia del COVID-19",
  "9"  = "Imposibilidad de realizar actividad por cuenta propia por la pandemia del COVID-19",
  "88" = "No sabe", "99" = "No responde"
)

# e23: "¿Cuál fue el motivo de su despido?" (solo si e22 = 1).
# e24: "¿Cuál fue el motivo de su renuncia?" (solo si e22 = 2).
# Ambas preguntas se incorporaron a partir del trimestre JJA 2020, recomendadas por
# la OIT en el contexto del impacto de la pandemia de COVID-19 (INE, 2020c).
lab_despido <- c(
  "1"  = "Por razones de edad (muy joven o edad avanzada)",
  "2"  = "Conflicto con su jefe o superior",
  "3"  = "Falta de calificación o capacitación",
  "4"  = "Ya no hubo más trabajo",
  "5"  = "Discriminación por su aspecto físico",
  "6"  = "Incumplimiento de funciones",
  "7"  = "Enfermedad o incapacidad propia",
  "8"  = "Embarazo y/o incompatibilidad con cuidado de personas en el hogar",
  "9"  = "Reducción de personal",
  "10" = "Pandemia del COVID-19",
  "11" = "Otra razón",
  "88" = "No sabe", "99" = "No responde"
)

lab_renuncia <- c(
  "1"  = "Realizar estudios o recibir formación",
  "2"  = "Cuidado de niños o de adultos enfermos o incapacitados",
  "3"  = "Motivos de salud",
  "4"  = "Deseaba un trabajo con mayores ingresos",
  "5"  = "Para mejorar su calidad de vida",
  "6"  = "Acoso o falta de respeto a su persona",
  "7"  = "Embarazo",
  "8"  = "Se cansó de ese trabajo",
  "9"  = "Pandemia del COVID-19",
  "10" = "Otra razón",
  "88" = "No sabe", "99" = "No responde"
)

# --- Etiquetas del módulo panel (variables medidas en MJJ 2025) --------------
# CAENES (CIIU Rev. 4 CL), secciones A a U.
lab_rama <- c(
  "1"  = "A. Agricultura, ganadería, silvicultura y pesca",
  "2"  = "B. Explotación de minas y canteras",
  "3"  = "C. Industrias manufactureras",
  "4"  = "D. Suministro de electricidad, gas, vapor y aire acondicionado",
  "5"  = "E. Suministro de agua y gestión de desechos",
  "6"  = "F. Construcción",
  "7"  = "G. Comercio al por mayor y al por menor; reparación de vehículos",
  "8"  = "H. Transporte y almacenamiento",
  "9"  = "I. Alojamiento y servicio de comidas",
  "10" = "J. Información y comunicaciones",
  "11" = "K. Actividades financieras y de seguros",
  "12" = "L. Actividades inmobiliarias",
  "13" = "M. Actividades profesionales, científicas y técnicas",
  "14" = "N. Actividades de servicios administrativos y de apoyo",
  "15" = "O. Administración pública y defensa",
  "16" = "P. Enseñanza",
  "17" = "Q. Salud humana y asistencia social",
  "18" = "R. Actividades artísticas, de entretenimiento y recreativas",
  "19" = "S. Otras actividades de servicios",
  "20" = "T. Hogares como empleadores de personal doméstico",
  "21" = "U. Organizaciones y órganos extraterritoriales",
  "999" = "No declarado"
)
orden_rama <- c(1:21, 999)

lab_categ <- c(
  "1" = "Empleador/a",
  "2" = "Trabajador/a por cuenta propia",
  "3" = "Asalariado/a del sector privado",
  "4" = "Asalariado/a del sector público",
  "5" = "Personal de servicio doméstico puertas afuera",
  "6" = "Personal de servicio doméstico puertas adentro",
  "7" = "Familiar o personal no remunerado"
)

lab_formal   <- c("1" = "Empleo formal", "2" = "Empleo informal")
lab_sectinst <- c("1" = "Sector formal", "2" = "Sector informal", "3" = "Hogares")

orden_situacion <- c(
  "Ocupado/a",
  "Desocupado/a: cesante",
  "Desocupado/a: busca trabajo por primera vez",
  "Fuera de la fuerza de trabajo: fuerza de trabajo potencial",
  "Fuera de la fuerza de trabajo: resto"
)

# b8 = "¿Tiene contrato escrito?" (1 sí, 2 no); b9 = "¿La duración de ese contrato o
# acuerdo de trabajo es...?" (1 definida o a plazo fijo, 2 indefinida). Ambas se
# preguntan solo a asalariados/as y personal de servicio doméstico.
orden_contrato <- c(
  "Con contrato escrito, de duración indefinida",
  "Con contrato escrito, de duración definida o a plazo fijo",
  "Sin contrato escrito, acuerdo de duración indefinida",
  "Sin contrato escrito, acuerdo de duración definida",
  "No sabe / no responde",
  "No aplica (empleador, cuenta propia o familiar no remunerado)"
)

lab_metodos <- c(
  e3_1  = "Envió CV a empresas o instituciones",
  e3_2  = "Consultó directamente con empleadores",
  e3_3  = "Pidió a conocidos o familiares que le avisaran",
  e3_4  = "Revisó y contestó anuncios",
  e3_5  = "Se inscribió o revisó anuncios en la OMIL",
  e3_6  = "Realizó gestiones para establecerse por su cuenta",
  e3_7  = "Estuvo buscando clientes o pedidos",
  e3_8  = "Puso anuncios",
  e3_9  = "Participó en una prueba o entrevista",
  e3_10 = "Consultó con agencias de empleo",
  e3_11 = "Actualizó su CV publicado en internet",
  e3_12 = "Ninguna de las anteriores"
)

# ---------------------------------------------------------------------------
# 3. Variables derivadas
# ---------------------------------------------------------------------------
meses_desde <- function(mes, anio, mes_ref, anio_ref) {
  mes  <- ifelse(mes %in% 1:12, mes, NA_real_)
  anio <- ifelse(!is.na(anio) & anio >= 1900 & anio <= anio_ref, anio, NA_real_)
  pmax((anio_ref - anio) * 12 + (mes_ref - mes), 0)
}

tramos_duracion <- function(d) {
  cut(d, breaks = c(-Inf, 2.99, 5.99, 11.99, 23.99, Inf),
      labels = c("Menos de 3 meses", "3 a 5 meses", "6 a 11 meses",
                 "12 a 23 meses", "24 meses o más"))
}

ene <- ene %>%
  mutate(
    desocupado     = as.integer(activ == 2),
    en_ft          = as.integer(ft == 1),
    dur_busqueda   = meses_desde(e6_mes,  e6_ano,  mes_encuesta, ANIO),
    dur_cesantia   = meses_desde(e21_mes, e21_ano, mes_encuesta, ANIO),
    tramo_busqueda = as.character(tramos_duracion(dur_busqueda)),
    tramo_cesantia = as.character(tramos_duracion(dur_cesantia)),
    tramo_busqueda = ifelse(is.na(tramo_busqueda), "No declarado", tramo_busqueda),
    tramo_cesantia = ifelse(is.na(tramo_cesantia), "No declarado", tramo_cesantia),
    nacional       = ifelse(nacionalidad == 152, "Chilena", "Extranjera"),
    indigena       = dplyr::recode(as.character(orig1), "1" = "Sí", "2" = "No",
                                   .default = "No declarado"),
    jefatura       = ifelse(parentesco == 1, "Jefe/a de hogar", "Otro parentesco")
  )

desoc <- ene %>% filter(desocupado == 1)
ftrab <- ene %>% filter(en_ft == 1)

# ---------------------------------------------------------------------------
# 4. Funciones de tabulación
# ---------------------------------------------------------------------------

# Variable que particiona a toda la fuerza de trabajo -> permite tasa de desocupación
tabla_con_tasa <- function(var) {
  d <- desoc %>% group_by(cat = .data[[var]]) %>%
    summarise(casos = n(), desocupados = sum(fact_cal, na.rm = TRUE), .groups = "drop")
  f <- ftrab %>% group_by(cat = .data[[var]]) %>%
    summarise(fuerza_trabajo = sum(fact_cal, na.rm = TRUE), .groups = "drop")
  left_join(d, f, by = "cat")
}

# Variable definida solo dentro del universo de personas desocupadas
tabla_simple <- function(var, base = desoc) {
  base %>% group_by(cat = .data[[var]]) %>%
    summarise(casos = n(), desocupados = sum(fact_cal, na.rm = TRUE), .groups = "drop")
}

etiquetar <- function(tab, labels, orden) {
  tab <- tab[match(orden, tab$cat), ]
  tab <- tab[!is.na(tab$cat), ]
  tab$Categoria <- unname(labels[as.character(tab$cat)])
  tab$cat <- NULL
  dplyr::relocate(tab, Categoria)
}

ordenar <- function(tab, orden) {
  tab <- dplyr::rename(tab, Categoria = cat)
  tab <- tab[match(orden, tab$Categoria), ]
  tab[!is.na(tab$Categoria), ]
}

# ---------------------------------------------------------------------------
# 5. Tablas
# ---------------------------------------------------------------------------
t_sexo   <- etiquetar(tabla_con_tasa("sexo"),       lab_sexo,       c(1, 2))
t_edad   <- etiquetar(tabla_con_tasa("tramo_edad"), lab_tramo_edad, 1:12)
t_region <- etiquetar(tabla_con_tasa("region"),     lab_region,     orden_region)
t_provincia <- etiquetar(tabla_con_tasa("provincia"), lab_provincia, orden_provincia)
t_educ   <- etiquetar(tabla_con_tasa("cine11_1d"),  lab_educ,       c(0, 1, 2, 3, 4, 5, 6, 9))
t_nac    <- ordenar(tabla_con_tasa("nacional"), c("Chilena", "Extranjera"))
t_indig  <- ordenar(tabla_con_tasa("indigena"), c("Sí", "No", "No declarado"))
t_jefe   <- ordenar(tabla_con_tasa("jefatura"), c("Jefe/a de hogar", "Otro parentesco"))

t_cond <- etiquetar(tabla_simple("cae_general"), lab_condicion, c(4, 5))
t_jorn <- etiquetar(tabla_simple("e7"),          lab_jornada,   c(1, 2, 3, 88, 99))

orden_dur <- c("Menos de 3 meses", "3 a 5 meses", "6 a 11 meses",
               "12 a 23 meses", "24 meses o más", "No declarado")
t_busq <- ordenar(tabla_simple("tramo_busqueda"), orden_dur)

cesantes <- desoc %>% filter(cae_general == 4)
t_cesa   <- ordenar(tabla_simple("tramo_cesantia", base = cesantes), orden_dur)
t_motivo <- etiquetar(tabla_simple("e22", base = cesantes), lab_motivo,
                      c(1, 2, 3, 4, 5, 6, 7, 8, 9, 88, 99))

# Detalle del motivo: e23 solo se pregunta a quienes fueron despedidos (e22 = 1)
# y e24 solo a quienes renunciaron (e22 = 2).
despedidos  <- cesantes %>% filter(e22 == 1)
renunciados <- cesantes %>% filter(e22 == 2)
t_despido  <- etiquetar(tabla_simple("e23", base = despedidos),  lab_despido,
                        c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 88, 99))
t_renuncia <- etiquetar(tabla_simple("e24", base = renunciados), lab_renuncia,
                        c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 88, 99))

# ---------------------------------------------------------------------------
# 5b. Módulo panel: enlace MJJ 2026 <-> MJJ 2025
# ---------------------------------------------------------------------------
# La ENE es un panel rotativo: cada persona seleccionada es entrevistada tres meses
# consecutivos, sale de la muestra y vuelve tres meses más, 12 meses después de la
# primera visita. El identificador de persona idrph permite enlazar ambas olas,
# siempre comparando el mismo mes calendario (mayo con mayo, junio con junio, etc.).
vars_prev <- c("idrph", "activ", "sexo", "edad", "cae_general", "ocup_form", "sector",
               "categoria_ocupacion", "r_p_rev4cl_caenes", "b8", "b9", "mes_encuesta")

prev <- ds %>%
  filter(ano == ANIO_PREV, mes_central == MES_CENTRAL) %>%
  select(all_of(vars_prev)) %>%
  collect() %>%
  rename_with(~ paste0(.x, "_25"), -idrph)

# Validación del enlace: mismo mes calendario, mismo sexo y diferencia de edad 0-2 años.
par <- ene %>%
  inner_join(prev, by = "idrph") %>%
  filter(mes_encuesta == mes_encuesta_25,
         sexo == sexo_25,
         edad - edad_25 >= 0, edad - edad_25 <= 2) %>%
  mutate(
    situacion_25 = case_when(
      activ_25 == 1                 ~ "Ocupado/a",
      cae_general_25 == 4           ~ "Desocupado/a: cesante",
      cae_general_25 == 5           ~ "Desocupado/a: busca trabajo por primera vez",
      cae_general_25 %in% c(6,7,8)  ~ "Fuera de la fuerza de trabajo: fuerza de trabajo potencial",
      cae_general_25 == 9           ~ "Fuera de la fuerza de trabajo: resto",
      TRUE                          ~ "No clasificado"),
    contrato_25 = case_when(
      b8_25 == 1 & b9_25 == 2 ~ orden_contrato[1],
      b8_25 == 1 & b9_25 == 1 ~ orden_contrato[2],
      b8_25 == 2 & b9_25 == 2 ~ orden_contrato[3],
      b8_25 == 2 & b9_25 == 1 ~ orden_contrato[4],
      !is.na(b8_25)           ~ orden_contrato[5],
      TRUE                    ~ orden_contrato[6])
  )

pan_desoc <- par %>% filter(desocupado == 1)          # desocupados/as 2026 con dato 2025
pan_ocup  <- pan_desoc %>% filter(activ_25 == 1)      # ...que estaban ocupados/as en 2025

t_p_sit      <- ordenar(tabla_simple("situacion_25", base = pan_desoc), orden_situacion)
t_p_rama     <- etiquetar(tabla_simple("r_p_rev4cl_caenes_25", base = pan_ocup),
                          lab_rama, orden_rama)
t_p_contrato <- ordenar(tabla_simple("contrato_25", base = pan_ocup), orden_contrato)
t_p_categ    <- etiquetar(tabla_simple("categoria_ocupacion_25", base = pan_ocup),
                          lab_categ, 1:7)
t_p_formal   <- etiquetar(tabla_simple("ocup_form_25", base = pan_ocup), lab_formal, c(1, 2))
t_p_sectinst <- etiquetar(tabla_simple("sector_25",    base = pan_ocup), lab_sectinst, c(1, 2, 3))

pan <- list(pares = nrow(par), casos = nrow(pan_desoc), casos_ocup = nrow(pan_ocup))

# Métodos de búsqueda: respuesta múltiple, los porcentajes no suman 100
t_metodos <- bind_rows(lapply(names(lab_metodos), function(v) {
  sel <- !is.na(desoc[[v]]) & desoc[[v]] == 1
  data.frame(Categoria   = unname(lab_metodos[v]),
             casos       = sum(sel),
             desocupados = sum(desoc$fact_cal[sel], na.rm = TRUE))
})) %>% arrange(desc(desocupados))

# ---------------------------------------------------------------------------
# 6. Totales de referencia
# ---------------------------------------------------------------------------
tot <- list(
  pet           = sum(ene$fact_cal[ene$pet == 1], na.rm = TRUE),
  ft            = sum(ftrab$fact_cal, na.rm = TRUE),
  ocupados      = sum(ene$fact_cal[ene$activ == 1], na.rm = TRUE),
  desocupados   = sum(desoc$fact_cal, na.rm = TRUE),
  cesantes      = sum(desoc$fact_cal[desoc$cae_general == 4], na.rm = TRUE),
  primera_vez   = sum(desoc$fact_cal[desoc$cae_general == 5], na.rm = TRUE),
  casos_desoc   = nrow(desoc),
  casos_muestra = nrow(ene)
)
cat("Tasa de desocupación MJJ 2026:", round(tot$desocupados / tot$ft * 100, 2), "%\n")
cat("Personas desocupadas:", format(round(tot$desocupados), big.mark = "."), "\n")

# ---------------------------------------------------------------------------
# 7. Exportación a Excel
# ---------------------------------------------------------------------------
wb <- createWorkbook()

est_titulo <- createStyle(fontName = "Arial", fontSize = 12, textDecoration = "bold")
est_header <- createStyle(fontName = "Arial", fontSize = 10, textDecoration = "bold",
                          fgFill = "#1F3864", fontColour = "white", halign = "center",
                          valign = "center", wrapText = TRUE, border = "TopBottomLeftRight")
est_texto  <- createStyle(fontName = "Arial", fontSize = 10)
est_num    <- createStyle(fontName = "Arial", fontSize = 10, numFmt = "#,##0")
est_pct    <- createStyle(fontName = "Arial", fontSize = 10, numFmt = "0.0")
est_total  <- createStyle(fontName = "Arial", fontSize = 10, textDecoration = "bold", border = "top")
est_nota   <- createStyle(fontName = "Arial", fontSize = 9, textDecoration = "italic")

escribir_hoja <- function(wb, hoja, titulo, tab, con_tasa, nota = NULL, multiple = FALSE,
                          pct = "% del total de desocupados") {
  addWorksheet(wb, hoja)
  writeData(wb, hoja, titulo, startRow = 1, startCol = 1)
  addStyle(wb, hoja, est_titulo, rows = 1, cols = 1)

  if (con_tasa) {
    encabezados <- c("Categoría", "Casos muestrales", "Personas desocupadas",
                     pct, "Fuerza de trabajo",
                     "Tasa de desocupación (%)")
    df <- data.frame(Categoria = tab$Categoria, casos = tab$casos,
                     desocupados = round(tab$desocupados), pct = NA_real_,
                     ft = round(tab$fuerza_trabajo), tasa = NA_real_)
  } else {
    encabezados <- c("Categoría", "Casos muestrales", "Personas desocupadas", pct)
    df <- data.frame(Categoria = tab$Categoria, casos = tab$casos,
                     desocupados = round(tab$desocupados), pct = NA_real_)
  }

  writeData(wb, hoja, t(encabezados), startRow = 3, startCol = 1, colNames = FALSE)
  addStyle(wb, hoja, est_header, rows = 3, cols = seq_along(encabezados), gridExpand = TRUE)
  writeData(wb, hoja, df, startRow = 4, startCol = 1, colNames = FALSE)

  n <- nrow(df); r0 <- 4; r1 <- 3 + n; rt <- r1 + 1

  for (i in seq_len(n)) {
    r <- r0 + i - 1
    writeFormula(wb, hoja, sprintf('=IFERROR(C%d/$C$%d*100,"")', r, rt), startRow = r, startCol = 4)
    if (con_tasa)
      writeFormula(wb, hoja, sprintf('=IFERROR(C%d/E%d*100,"")', r, r), startRow = r, startCol = 6)
  }

  if (multiple) {
    # Respuesta múltiple: el denominador es el total de personas desocupadas y no
    # la suma de las filas, porque una persona puede aparecer en varias categorías.
    writeData(wb, hoja, "Total de personas desocupadas", startRow = rt, startCol = 1)
    writeData(wb, hoja, tot$casos_desoc, startRow = rt, startCol = 2)
    writeFormula(wb, hoja, "=Resumen!B7", startRow = rt, startCol = 3)
    writeFormula(wb, hoja, sprintf('=IFERROR(C%d/$C$%d*100,"")', rt, rt), startRow = rt, startCol = 4)
  } else {
  writeData(wb, hoja, "Total", startRow = rt, startCol = 1)
  writeFormula(wb, hoja, sprintf("=SUM(B%d:B%d)", r0, r1), startRow = rt, startCol = 2)
  writeFormula(wb, hoja, sprintf("=SUM(C%d:C%d)", r0, r1), startRow = rt, startCol = 3)
  writeFormula(wb, hoja, sprintf("=SUM(D%d:D%d)", r0, r1), startRow = rt, startCol = 4)
  }
  if (con_tasa) {
    writeFormula(wb, hoja, sprintf("=SUM(E%d:E%d)", r0, r1), startRow = rt, startCol = 5)
    writeFormula(wb, hoja, sprintf('=IFERROR(C%d/E%d*100,"")', rt, rt), startRow = rt, startCol = 6)
  }

  ncols <- length(encabezados)
  addStyle(wb, hoja, est_texto, rows = r0:rt, cols = 1, gridExpand = TRUE)
  addStyle(wb, hoja, est_num, rows = r0:rt, cols = c(2, 3, if (con_tasa) 5), gridExpand = TRUE)
  addStyle(wb, hoja, est_pct, rows = r0:rt, cols = c(4, if (con_tasa) 6), gridExpand = TRUE)
  addStyle(wb, hoja, est_total, rows = rt, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, hoja, cols = 1, widths = 42)
  setColWidths(wb, hoja, cols = 2:ncols, widths = 18)

  if (!is.null(nota)) {
    writeData(wb, hoja, paste("Nota:", nota), startRow = rt + 2, startCol = 1)
    addStyle(wb, hoja, est_nota, rows = rt + 2, cols = 1)
  }
  freezePanes(wb, hoja, firstActiveRow = 4)
}

# --- Notas metodológicas
addWorksheet(wb, "Notas")
notas <- data.frame(x = c(
  "Caracterización de las personas desocupadas - ENE, trimestre móvil mayo-junio-julio 2026",
  "",
  "Fuente: Instituto Nacional de Estadísticas (INE), Encuesta Nacional de Empleo (ENE), microdatos.",
  "Elaboración propia. Trimestre móvil MJJ 2026 (mes central: junio).",
  "",
  "DEFINICIONES",
  "Población en edad de trabajar (PET): personas de 15 años y más.",
  "Fuerza de trabajo: personas ocupadas más personas desocupadas (ft = 1).",
  "Persona desocupada (activ = 2): sin empleo, que buscó trabajo en las últimas cuatro semanas y",
  "   está disponible para trabajar (marco conceptual OIT, 19a CIET).",
  "Cesante (cae_general = 4): persona desocupada con experiencia laboral previa.",
  "Busca trabajo por primera vez (cae_general = 5): persona desocupada sin empleo anterior.",
  "Tasa de desocupación = personas desocupadas / fuerza de trabajo * 100.",
  "",
  "EXPANSIÓN",
  "Todas las cifras de personas están expandidas con el factor trimestral fact_cal (proyecciones",
  "   de población base Censo 2017). Los 'casos muestrales' corresponden a personas encuestadas.",
  "Se recomienda cautela al interpretar categorías con menos de 50 casos muestrales.",
  "",
  "VARIABLES UTILIZADAS",
  "sexo; tramo_edad; region; cine11_1d (nivel educacional); nacionalidad (ISO 3166, 152 = Chile);",
  "orig1 (pueblo indígena); parentesco (jefatura de hogar); cae_general (condición de desocupación);",
  "e7 (jornada buscada); e3_1 a e3_12 (gestiones de búsqueda); e6_mes/e6_ano (inicio de la búsqueda);",
  "e21_mes/e21_ano (término del último empleo); e22 (motivo de término del último empleo);",
  "e23 (motivo del despido, solo si e22 = 1); e24 (motivo de la renuncia, solo si e22 = 2).",
  "Enlace panel con MJJ 2025 por idrph; de 2025 se usan activ, cae_general, r_p_rev4cl_caenes,",
  "   categoria_ocupacion, ocup_form, sector, b8 y b9.",
  "",
  "SUPUESTOS",
  "La duración de la búsqueda y de la cesantía se calcula como diferencia en meses entre el mes de",
  "   la entrevista (mes_encuesta) y la fecha declarada; los valores negativos se truncan en 0 y los",
  "   códigos 88 / 99 / 9999 se clasifican como 'No declarado'.",
  "Las etiquetas de cine11_1d se validaron cruzando esa variable con 'nivel' y 'termino_nivel'.",
  "Las glosas de e22 (pregunta E19) para los códigos 3 a 9, 88 y 99 provienen del libro de códigos del INE.",
  "   Los códigos 1 y 2 se etiquetan como despido y renuncia porque son los que habilitan las preguntas",
  "   e23 (motivo de despido) y e24 (motivo de renuncia).",
  "Las glosas de e23 y e24 provienen del libro de códigos del INE; ambas preguntas se incorporaron",
  "   a partir del trimestre JJA 2020 por recomendación de la OIT (INE, 2020c).",
  "",
  "",
  "ESTRUCTURA PANEL (hojas Panel_*)",
  "La ENE es una encuesta de panel rotativo: cada persona seleccionada es entrevistada durante tres",
  "   meses consecutivos, sale de la muestra y vuelve a ser entrevistada tres meses más, 12 meses",
  "   después de la primera visita. Eso permite enlazar MJJ 2026 con MJJ 2025 por el identificador de",
  "   persona idrph, siempre en el mismo mes calendario (mayo con mayo, junio con junio, julio con julio).",
  paste0("Pares enlazados y validados: ", format(pan$pares, big.mark = "."),
         " personas entrevistadas en ambos trimestres móviles."),
  "Validación del enlace: se exige mismo sexo y una diferencia de edad de 0 a 2 años entre ambas olas;",
  "   los pares que no cumplen se descartan.",
  sprintf("Cobertura: %s de las %s personas desocupadas de MJJ 2026 (%.1f%%) tienen dato de 2025.",
          format(pan$casos, big.mark = "."), format(tot$casos_desoc, big.mark = "."),
          pan$casos / tot$casos_desoc * 100),
  "ADVERTENCIA: la submuestra panel no tiene diseño muestral ni factores de expansión propios. Las cifras",
  "   expandidas de las hojas Panel_* se calculan con fact_cal de 2026 y son solo indicativas; la lectura",
  "   recomendada es la distribución porcentual, y no deben presentarse como estimaciones oficiales del INE.",
  "Variables de 2025 utilizadas: activ y cae_general (situación laboral); r_p_rev4cl_caenes (rama CAENES);",
  "   categoria_ocupacion; ocup_form (formalidad); sector (sector institucional); b8 y b9 (contrato).",
  "",
  "Script que genera este archivo: ene_empleo_microdatos/desocupados_mjj2026.R"
))
writeData(wb, "Notas", notas, colNames = FALSE)
addStyle(wb, "Notas", createStyle(fontName = "Arial", fontSize = 10),
         rows = 1:nrow(notas), cols = 1, gridExpand = TRUE)
addStyle(wb, "Notas", est_titulo, rows = 1, cols = 1)
setColWidths(wb, "Notas", cols = 1, widths = 105)

# --- Resumen
addWorksheet(wb, "Resumen")
res <- data.frame(
  Indicador = c("Población en edad de trabajar (PET)", "Fuerza de trabajo",
                "Personas ocupadas", "Personas desocupadas",
                "   Cesantes", "   Buscan trabajo por primera vez",
                "Tasa de desocupación (%)", "Tasa de cesantía (%)",
                "Casos muestrales (total encuesta)", "Casos muestrales (personas desocupadas)"),
  Valor = c(round(tot$pet), round(tot$ft), round(tot$ocupados), round(tot$desocupados),
            round(tot$cesantes), round(tot$primera_vez), NA, NA,
            tot$casos_muestra, tot$casos_desoc)
)
writeData(wb, "Resumen", "Resumen - personas desocupadas, ENE MJJ 2026", startRow = 1, startCol = 1)
addStyle(wb, "Resumen", est_titulo, rows = 1, cols = 1)
writeData(wb, "Resumen", t(c("Indicador", "Valor")), startRow = 3, startCol = 1, colNames = FALSE)
addStyle(wb, "Resumen", est_header, rows = 3, cols = 1:2, gridExpand = TRUE)
writeData(wb, "Resumen", res, startRow = 4, startCol = 1, colNames = FALSE)
writeFormula(wb, "Resumen", "=B7/B5*100", startRow = 10, startCol = 2)  # tasa de desocupación
writeFormula(wb, "Resumen", "=B8/B5*100", startRow = 11, startCol = 2)  # tasa de cesantía
addStyle(wb, "Resumen", est_texto, rows = 4:13, cols = 1, gridExpand = TRUE)
addStyle(wb, "Resumen", est_num,   rows = 4:13, cols = 2, gridExpand = TRUE)
addStyle(wb, "Resumen", est_pct,   rows = 10:11, cols = 2, gridExpand = TRUE)
setColWidths(wb, "Resumen", cols = 1, widths = 42)
setColWidths(wb, "Resumen", cols = 2, widths = 18)

# --- Hojas de caracterización
escribir_hoja(wb, "Sexo",   "Personas desocupadas según sexo - MJJ 2026", t_sexo, TRUE)
escribir_hoja(wb, "Edad",   "Personas desocupadas según tramo de edad - MJJ 2026", t_edad, TRUE)
escribir_hoja(wb, "Region", "Personas desocupadas según región - MJJ 2026", t_region, TRUE)
escribir_hoja(wb, "Provincia", "Personas desocupadas según provincia - MJJ 2026", t_provincia, TRUE,
              nota = 'Nota: la ENE tiene representatividad regional, no provincial. Estas cifras son elaboración propia y no constituyen estimaciones oficiales del INE: en varias provincias el número de casos muestrales de personas desocupadas es muy bajo y la tasa resultante tiene un error muestral alto. Las provincias de Isla de Pascua y Palena no aparecen porque no forman parte de la muestra del trimestre. El nombre de cada provincia va acompañado de su región.')
escribir_hoja(wb, "Educacion", "Personas desocupadas según nivel educacional alcanzado - MJJ 2026",
              t_educ, TRUE, nota = "Clasificación CINE 2011 (variable cine11_1d de la ENE).")
escribir_hoja(wb, "Nacionalidad", "Personas desocupadas según nacionalidad - MJJ 2026", t_nac, TRUE)
escribir_hoja(wb, "Pueblo_indigena",
              "Personas desocupadas según pertenencia a pueblo indígena u originario - MJJ 2026",
              t_indig, TRUE)
escribir_hoja(wb, "Jefatura_hogar", "Personas desocupadas según jefatura de hogar - MJJ 2026",
              t_jefe, TRUE)
escribir_hoja(wb, "Condicion",
              "Personas desocupadas según condición: cesante o busca trabajo por primera vez - MJJ 2026",
              t_cond, FALSE)
escribir_hoja(wb, "Duracion_busqueda",
              "Personas desocupadas según duración de la búsqueda de empleo - MJJ 2026", t_busq, FALSE,
              nota = "Meses transcurridos desde que declara haber comenzado a buscar trabajo (e6).")
escribir_hoja(wb, "Duracion_cesantia",
              "Personas cesantes según tiempo transcurrido desde su último empleo - MJJ 2026", t_cesa, FALSE,
              nota = "Universo: solo personas cesantes (cae_general = 4). Base: e21.")
escribir_hoja(wb, "Motivo_termino",
              "Personas cesantes según motivo de término de su último empleo - MJJ 2026", t_motivo, FALSE,
              nota = paste("Universo: solo personas cesantes (e22, pregunta E19 del cuestionario). Quienes responden",
                           "'despido' detallan el motivo en e23 y quienes responden 'renuncia' lo detallan en e24."))
escribir_hoja(wb, "Motivo_despido",
              "Personas cesantes despedidas según motivo del despido - MJJ 2026", t_despido, FALSE,
              nota = paste("Universo: personas cesantes cuyo último empleo terminó por despido (e22 = 1).",
                           "Variable e23. Pregunta incorporada a partir del trimestre JJA 2020, recomendada",
                           "por la OIT en el contexto del impacto de la pandemia de COVID-19 sobre la",
                           "población no ocupada (INE, 2020c)."))
escribir_hoja(wb, "Motivo_renuncia",
              "Personas cesantes que renunciaron según motivo de la renuncia - MJJ 2026", t_renuncia, FALSE,
              nota = paste("Universo: personas cesantes cuyo último empleo terminó por renuncia (e22 = 2).",
                           "Variable e24. Pregunta incorporada a partir del trimestre JJA 2020, recomendada",
                           "por la OIT en el contexto del impacto de la pandemia de COVID-19 sobre la",
                           "población no ocupada (INE, 2020c)."))
escribir_hoja(wb, "Jornada_buscada",
              "Personas desocupadas según tipo de jornada buscada - MJJ 2026", t_jorn, FALSE)
escribir_hoja(wb, "Metodos_busqueda",
              "Personas desocupadas según gestiones de búsqueda realizadas - MJJ 2026", t_metodos, FALSE,
              nota = paste("Respuesta múltiple: una persona puede declarar varias gestiones, por lo que",
                           "los porcentajes no suman 100. El denominador es el total de personas desocupadas."),
              multiple = TRUE)

# --- Hojas del módulo panel
nota_panel <- sprintf(
  paste("Panel ENE: %s de las %s personas desocupadas de MJJ 2026 (%.1f%%) también fueron",
        "entrevistadas en MJJ 2025. La submuestra panel no tiene diseño ni factores de expansión",
        "propios, por lo que las cifras expandidas son indicativas y no constituyen estimaciones",
        "oficiales; el porcentaje es la lectura recomendada."),
  format(pan$casos, big.mark = "."), format(tot$casos_desoc, big.mark = "."),
  pan$casos / tot$casos_desoc * 100)

nota_ocup <- sprintf(
  "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 (%d casos). ",
  pan$casos_ocup)

escribir_hoja(wb, "Panel_situacion_2025",
              "¿En qué situación estaban en MJJ 2025 las personas desocupadas de MJJ 2026?",
              t_p_sit, FALSE, pct = "% de las personas enlazadas",
              nota = paste(nota_panel, "Fuerza de trabajo potencial: personas fuera de la fuerza de",
                           "trabajo disponibles para trabajar o que buscaron sin estar disponibles",
                           "(cae_general 6 a 8)."))
escribir_hoja(wb, "Panel_rama_2025",
              "Rama de actividad en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
              t_p_rama, FALSE, pct = "% de quienes estaban ocupados/as",
              nota = paste0(nota_ocup, "Clasificación CAENES (CIIU Rev. 4 CL), variable ",
                            "r_p_rev4cl_caenes de 2025. ", nota_panel))
escribir_hoja(wb, "Panel_contrato_2025",
              "Tipo de contrato en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
              t_p_contrato, FALSE, pct = "% de quienes estaban ocupados/as",
              nota = paste0(nota_ocup, "Construido con b8 (¿tiene contrato escrito?) y b9 ",
                            "(duración del contrato o acuerdo). No aplica a empleadores, cuenta ",
                            "propia ni familiares no remunerados. ", nota_panel))
escribir_hoja(wb, "Panel_categoria_2025",
              "Categoría ocupacional en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
              t_p_categ, FALSE, pct = "% de quienes estaban ocupados/as",
              nota = paste0(nota_ocup, nota_panel))
escribir_hoja(wb, "Panel_formalidad_2025",
              "Formalidad del empleo en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
              t_p_formal, FALSE, pct = "% de quienes estaban ocupados/as",
              nota = paste0(nota_ocup, "Variable ocup_form (condición de formalidad del empleo, ",
                            "marco OIT 17a CIET). ", nota_panel))
escribir_hoja(wb, "Panel_sector_inst_2025",
              "Sector institucional del empleo en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
              t_p_sectinst, FALSE, pct = "% de quienes estaban ocupados/as",
              nota = paste0(nota_ocup, "Variable sector: unidad productiva formal, informal u ",
                            "hogares. ", nota_panel))

dir.create(dirname(ruta_salida), recursive = TRUE, showWarnings = FALSE)
saveWorkbook(wb, ruta_salida, overwrite = TRUE)
cat("Archivo generado:", ruta_salida, "\n")
