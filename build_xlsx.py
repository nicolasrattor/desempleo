# Espejo en Python de desocupados_mjj2026.R. Existe porque el entorno donde se genero el
# archivo no tenia R instalado; ambas rutas producen exactamente el mismo .xlsx.
# Uso, desde la raiz del repositorio:  python3 build_xlsx.py
import os
import numpy as np, pandas as pd, pyarrow.parquet as pq
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Los microdatos no viajan en el repositorio: se esperan en ./data con la particion
# original ano=/mes_central=, o en la ruta que indique la variable de entorno ENE_DATA.
BASE = os.environ.get("ENE_DATA", "data")
ANIO, MES = 2026, 6
ene = pq.read_table(f"{BASE}/ano={ANIO}/mes_central={MES}/part-0.parquet").to_pandas()

lab_sexo = {1: "Hombre", 2: "Mujer"}
lab_edad = {1:"15 a 19 años",2:"20 a 24 años",3:"25 a 29 años",4:"30 a 34 años",5:"35 a 39 años",
            6:"40 a 44 años",7:"45 a 49 años",8:"50 a 54 años",9:"55 a 59 años",10:"60 a 64 años",
            11:"65 a 69 años",12:"70 años o más"}
lab_region = {15:"Arica y Parinacota",1:"Tarapacá",2:"Antofagasta",3:"Atacama",4:"Coquimbo",
              5:"Valparaíso",13:"Metropolitana de Santiago",6:"O'Higgins",7:"Maule",16:"Ñuble",
              8:"Biobío",9:"La Araucanía",14:"Los Ríos",10:"Los Lagos",11:"Aysén",12:"Magallanes"}
orden_region = [15,1,2,3,4,5,13,6,7,16,8,9,14,10,11,12]
lab_provincia = {
    151: "Arica (Arica y Parinacota)",
    152: "Parinacota (Arica y Parinacota)",
    11: "Iquique (Tarapacá)",
    14: "Tamarugal (Tarapacá)",
    21: "Antofagasta (Antofagasta)",
    22: "El Loa (Antofagasta)",
    23: "Tocopilla (Antofagasta)",
    31: "Copiapó (Atacama)",
    32: "Chañaral (Atacama)",
    33: "Huasco (Atacama)",
    41: "Elqui (Coquimbo)",
    42: "Choapa (Coquimbo)",
    43: "Limarí (Coquimbo)",
    51: "Valparaíso (Valparaíso)",
    53: "Los Andes (Valparaíso)",
    54: "Petorca (Valparaíso)",
    55: "Quillota (Valparaíso)",
    56: "San Antonio (Valparaíso)",
    57: "San Felipe de Aconcagua (Valparaíso)",
    58: "Marga Marga (Valparaíso)",
    131: "Santiago (Metropolitana)",
    132: "Cordillera (Metropolitana)",
    133: "Chacabuco (Metropolitana)",
    134: "Maipo (Metropolitana)",
    135: "Melipilla (Metropolitana)",
    136: "Talagante (Metropolitana)",
    61: "Cachapoal (O'Higgins)",
    62: "Cardenal Caro (O'Higgins)",
    63: "Colchagua (O'Higgins)",
    71: "Talca (Maule)",
    72: "Cauquenes (Maule)",
    73: "Curicó (Maule)",
    74: "Linares (Maule)",
    161: "Diguillín (Ñuble)",
    162: "Itata (Ñuble)",
    163: "Punilla (Ñuble)",
    81: "Concepción (Biobío)",
    82: "Arauco (Biobío)",
    83: "Biobío (Biobío)",
    91: "Cautín (La Araucanía)",
    92: "Malleco (La Araucanía)",
    141: "Valdivia (Los Ríos)",
    142: "Ranco (Los Ríos)",
    101: "Llanquihue (Los Lagos)",
    102: "Chiloé (Los Lagos)",
    103: "Osorno (Los Lagos)",
    111: "Coyhaique (Aysén)",
    112: "Aysén (Aysén)",
    113: "Capitán Prat (Aysén)",
    114: "General Carrera (Aysén)",
    121: "Magallanes (Magallanes)",
    122: "Antártica Chilena (Magallanes)",
    123: "Tierra del Fuego (Magallanes)",
    124: "Última Esperanza (Magallanes)",
}
orden_provincia = [151, 152, 11, 14, 21, 22, 23, 31, 32, 33, 41, 42, 43, 51, 53, 54, 55, 56, 57, 58, 131, 132, 133, 134, 135, 136, 61, 62, 63, 71, 72, 73, 74, 161, 162, 163, 81, 82, 83, 91, 92, 141, 142, 101, 102, 103, 111, 112, 113, 114, 121, 122, 123, 124]
lab_educ = {0:"Sin educación formal o básica incompleta",1:"Básica completa o media incompleta",
            2:"Media completa",3:"Técnica de nivel superior",4:"Profesional (universitaria)",
            5:"Magíster",6:"Doctorado",9:"Nivel no declarado"}
lab_cond = {4:"Cesante",5:"Busca trabajo por primera vez"}
lab_jorn = {1:"Jornada completa",2:"Jornada parcial",3:"Indistinto",88:"No sabe",99:"No responde"}
lab_motivo = {
    1: "Despido",
    2: "Renuncia",
    3: "Fin del contrato, proyecto, faena, temporada o reemplazo",
    4: "Jubilación",
    5: "Término del ejercicio de la actividad por cuenta propia",
    6: "Quiebra o cierre de negocio por razones económicas",
    7: "Quiebra o cierre de negocio por un fenómeno natural o siniestro",
    8: "Quiebra o cierre de negocio por la pandemia del COVID-19",
    9: "Imposibilidad de realizar actividad por cuenta propia por la pandemia del COVID-19",
    88: "No sabe", 99: "No responde"}
lab_despido = {
    1: "Por razones de edad (muy joven o edad avanzada)",
    2: "Conflicto con su jefe o superior",
    3: "Falta de calificación o capacitación",
    4: "Ya no hubo más trabajo",
    5: "Discriminación por su aspecto físico",
    6: "Incumplimiento de funciones",
    7: "Enfermedad o incapacidad propia",
    8: "Embarazo y/o incompatibilidad con cuidado de personas en el hogar",
    9: "Reducción de personal",
    10: "Pandemia del COVID-19",
    11: "Otra razón",
    88: "No sabe", 99: "No responde"}

lab_renuncia = {
    1: "Realizar estudios o recibir formación",
    2: "Cuidado de niños o de adultos enfermos o incapacitados",
    3: "Motivos de salud",
    4: "Deseaba un trabajo con mayores ingresos",
    5: "Para mejorar su calidad de vida",
    6: "Acoso o falta de respeto a su persona",
    7: "Embarazo",
    8: "Se cansó de ese trabajo",
    9: "Pandemia del COVID-19",
    10: "Otra razón",
    88: "No sabe", 99: "No responde"}

lab_metodos = {"e3_1":"Envió CV a empresas o instituciones","e3_2":"Consultó directamente con empleadores",
    "e3_3":"Pidió a conocidos o familiares que le avisaran","e3_4":"Revisó y contestó anuncios",
    "e3_5":"Se inscribió o revisó anuncios en la OMIL","e3_6":"Realizó gestiones para establecerse por su cuenta",
    "e3_7":"Estuvo buscando clientes o pedidos","e3_8":"Puso anuncios",
    "e3_9":"Participó en una prueba o entrevista","e3_10":"Consultó con agencias de empleo",
    "e3_11":"Actualizó su CV publicado en internet","e3_12":"Ninguna de las anteriores"}

def meses_desde(mes, anio, mes_ref, anio_ref=ANIO):
    m = mes.where(mes.between(1, 12))
    a = anio.where((anio >= 1900) & (anio <= anio_ref))
    return ((anio_ref - a) * 12 + (mes_ref - m)).clip(lower=0)

ORD_DUR = ["Menos de 3 meses","3 a 5 meses","6 a 11 meses","12 a 23 meses","24 meses o más","No declarado"]
def tramos(d):
    return pd.cut(d, [-np.inf,2.99,5.99,11.99,23.99,np.inf], labels=ORD_DUR[:5]).astype(object)

ene["dur_busqueda"] = meses_desde(ene.e6_mes, ene.e6_ano, ene.mes_encuesta)
ene["dur_cesantia"] = meses_desde(ene.e21_mes, ene.e21_ano, ene.mes_encuesta)
ene["tramo_busqueda"] = tramos(ene.dur_busqueda).fillna("No declarado")
ene["tramo_cesantia"] = tramos(ene.dur_cesantia).fillna("No declarado")
ene["nacional"] = np.where(ene.nacionalidad == 152, "Chilena", "Extranjera")
ene["indigena"] = ene.orig1.map({1: "Sí", 2: "No"}).fillna("No declarado")
ene["jefatura"] = np.where(ene.parentesco == 1, "Jefe/a de hogar", "Otro parentesco")

desoc = ene[ene.activ == 2].copy()
ftrab = ene[ene.ft == 1].copy()
cesantes = desoc[desoc.cae_general == 4].copy()

def con_tasa(var, labels=None, orden=None):
    d = desoc.groupby(var).agg(casos=("fact_cal","size"), desocupados=("fact_cal","sum"))
    f = ftrab.groupby(var)["fact_cal"].sum().rename("ft")
    t = d.join(f).reset_index().rename(columns={var:"cat"})
    return _fmt(t, labels, orden)

def simple(var, base=None, labels=None, orden=None):
    b = desoc if base is None else base
    t = b.groupby(var).agg(casos=("fact_cal","size"), desocupados=("fact_cal","sum")).reset_index()
    t = t.rename(columns={var:"cat"})
    return _fmt(t, labels, orden)

def _fmt(t, labels, orden):
    # Las columnas provenientes del merge con 2025 son float64 (tienen NA), por lo que
    # el orden y las etiquetas se normalizan a float antes de reindexar / mapear.
    num = pd.api.types.is_numeric_dtype(t["cat"])
    if orden is not None:
        if num:
            t["cat"] = t["cat"].astype(float)
            orden = [float(o) for o in orden]
        t = t.set_index("cat").reindex(orden).dropna(how="all").reset_index()
    if labels:
        L = {(float(k) if num else k): v for k, v in labels.items()}
        t["Categoria"] = t["cat"].map(L)
    else:
        t["Categoria"] = t["cat"]
    return t.drop(columns="cat")

t_sexo   = con_tasa("sexo", lab_sexo, [1,2])
t_edad   = con_tasa("tramo_edad", lab_edad, list(range(1,13)))
t_region = con_tasa("region", lab_region, orden_region)
t_provincia = con_tasa("provincia", lab_provincia, orden_provincia)
t_educ   = con_tasa("cine11_1d", lab_educ, [0,1,2,3,4,5,6,9])
t_nac    = con_tasa("nacional", None, ["Chilena","Extranjera"])
t_indig  = con_tasa("indigena", None, ["Sí","No","No declarado"])
t_jefe   = con_tasa("jefatura", None, ["Jefe/a de hogar","Otro parentesco"])
t_cond   = simple("cae_general", labels=lab_cond, orden=[4,5])
t_jorn   = simple("e7", labels=lab_jorn, orden=[1,2,3,88,99])
t_busq   = simple("tramo_busqueda", orden=ORD_DUR)
t_cesa   = simple("tramo_cesantia", base=cesantes, orden=ORD_DUR)
t_motivo = simple("e22", base=cesantes, labels=lab_motivo, orden=[1,2,3,4,5,6,7,8,9,88,99])

despedidos  = cesantes[cesantes.e22 == 1].copy()
renunciados = cesantes[cesantes.e22 == 2].copy()
t_despido  = simple("e23", base=despedidos,  labels=lab_despido,
                    orden=[1,2,3,4,5,6,7,8,9,10,11,88,99])
t_renuncia = simple("e24", base=renunciados, labels=lab_renuncia,
                    orden=[1,2,3,4,5,6,7,8,9,10,88,99])

# ---------------------------------------------------------------- PANEL 2025-2026
# La ENE es un panel rotativo: cada persona (idrph) vuelve a ser entrevistada 12 meses
# después, en el mismo mes calendario. Se enlaza MJJ 2026 con MJJ 2025 y se valida el
# par exigiendo mismo sexo y diferencia de edad entre 0 y 2 años.
c25 = ["idrph","activ","sexo","edad","cae_general","ocup_form","sector",
       "categoria_ocupacion","r_p_rev4cl_caenes","b8","b9","fact_cal","mes_encuesta"]
prev = pq.read_table(f"{BASE}/ano={ANIO-1}/mes_central={MES}/part-0.parquet",
                     columns=c25).to_pandas()
# Renombrar explícitamente: casi todas estas columnas existen también en 2026 y confiar en
# los sufijos del merge haría que se leyeran los valores de 2026 (nulos para desocupados).
prev = prev.rename(columns={c: c + "_25" for c in prev.columns if c != "idrph"})
par = ene.merge(prev, on="idrph")
par = par[par.mes_encuesta == par.mes_encuesta_25]
par = par[(par.sexo == par.sexo_25) & (par.edad - par.edad_25).between(0, 2)].copy()

par["sit25"] = np.select(
    [par.activ_25 == 1,
     par.cae_general_25 == 4,
     par.cae_general_25 == 5,
     par.cae_general_25.isin([6, 7, 8]),
     par.cae_general_25 == 9],
    ["Ocupado/a",
     "Desocupado/a: cesante",
     "Desocupado/a: busca trabajo por primera vez",
     "Fuera de la fuerza de trabajo: fuerza de trabajo potencial",
     "Fuera de la fuerza de trabajo: resto"], default="No clasificado")
ORD_SIT = ["Ocupado/a","Desocupado/a: cesante","Desocupado/a: busca trabajo por primera vez",
           "Fuera de la fuerza de trabajo: fuerza de trabajo potencial",
           "Fuera de la fuerza de trabajo: resto"]

lab_rama = {1:"A. Agricultura, ganadería, silvicultura y pesca",
    2:"B. Explotación de minas y canteras", 3:"C. Industrias manufactureras",
    4:"D. Suministro de electricidad, gas, vapor y aire acondicionado",
    5:"E. Suministro de agua y gestión de desechos", 6:"F. Construcción",
    7:"G. Comercio al por mayor y al por menor; reparación de vehículos",
    8:"H. Transporte y almacenamiento", 9:"I. Alojamiento y servicio de comidas",
    10:"J. Información y comunicaciones", 11:"K. Actividades financieras y de seguros",
    12:"L. Actividades inmobiliarias", 13:"M. Actividades profesionales, científicas y técnicas",
    14:"N. Actividades de servicios administrativos y de apoyo",
    15:"O. Administración pública y defensa", 16:"P. Enseñanza",
    17:"Q. Salud humana y asistencia social",
    18:"R. Actividades artísticas, de entretenimiento y recreativas",
    19:"S. Otras actividades de servicios",
    20:"T. Hogares como empleadores de personal doméstico",
    21:"U. Organizaciones y órganos extraterritoriales", 999:"No declarado"}
ORD_RAMA = list(range(1, 22)) + [999]

lab_categ = {1:"Empleador/a", 2:"Trabajador/a por cuenta propia",
    3:"Asalariado/a del sector privado", 4:"Asalariado/a del sector público",
    5:"Personal de servicio doméstico puertas afuera",
    6:"Personal de servicio doméstico puertas adentro",
    7:"Familiar o personal no remunerado"}

lab_formal = {1:"Empleo formal", 2:"Empleo informal"}
lab_sectinst = {1:"Sector formal", 2:"Sector informal", 3:"Hogares"}

ORD_CONTRATO = ["Con contrato escrito, de duración indefinida",
                "Con contrato escrito, de duración definida o a plazo fijo",
                "Sin contrato escrito, acuerdo de duración indefinida",
                "Sin contrato escrito, acuerdo de duración definida",
                "No sabe / no responde",
                "No aplica (empleador, cuenta propia o familiar no remunerado)"]
def tipo_contrato(b8, b9):
    esc = b8.map({1: 1, 2: 2})          # 1 = con contrato escrito, 2 = sin contrato escrito
    dur = b9.map({1: 1, 2: 2})          # 1 = definida o plazo fijo, 2 = indefinida
    out = pd.Series(ORD_CONTRATO[5], index=b8.index, dtype=object)   # no aplica
    out[b8.notna()] = ORD_CONTRATO[4]                                # no sabe / no responde
    ok = esc.notna() & dur.notna()
    out[ok] = np.where(esc[ok] == 1,
                       np.where(dur[ok] == 2, ORD_CONTRATO[0], ORD_CONTRATO[1]),
                       np.where(dur[ok] == 2, ORD_CONTRATO[2], ORD_CONTRATO[3]))
    return out

par["contrato25"] = tipo_contrato(par.b8_25, par.b9_25)
pan_desoc = par[par.activ == 2].copy()          # desocupados/as en MJJ 2026, con dato 2025
pan_ocup  = pan_desoc[pan_desoc.activ_25 == 1]  # ...que estaban ocupados/as en MJJ 2025

t_p_sit      = simple("sit25", base=pan_desoc, orden=ORD_SIT)
t_p_rama     = simple("r_p_rev4cl_caenes_25", base=pan_ocup, labels=lab_rama, orden=ORD_RAMA)
t_p_categ    = simple("categoria_ocupacion_25", base=pan_ocup, labels=lab_categ, orden=[1,2,3,4,5,6,7])
t_p_contrato = simple("contrato25", base=pan_ocup, orden=ORD_CONTRATO)
t_p_formal   = simple("ocup_form_25", base=pan_ocup, labels=lab_formal, orden=[1,2])
t_p_sectinst = simple("sector_25", base=pan_ocup, labels=lab_sectinst, orden=[1,2,3])

pan = dict(pares=len(par), casos=len(pan_desoc), exp=pan_desoc.fact_cal.sum(),
           ocup=len(pan_ocup), exp_ocup=pan_ocup.fact_cal.sum())

rows = []
for v, lab in lab_metodos.items():
    sel = desoc[v] == 1
    rows.append({"casos": int(sel.sum()), "desocupados": desoc.fact_cal[sel].sum(), "Categoria": lab})
t_metodos = pd.DataFrame(rows).sort_values("desocupados", ascending=False).reset_index(drop=True)

tot = dict(
    pet=ene.fact_cal[ene.pet == 1].sum(), ft=ftrab.fact_cal.sum(),
    ocupados=ene.fact_cal[ene.activ == 1].sum(), desocupados=desoc.fact_cal.sum(),
    cesantes=cesantes.fact_cal.sum(), primera=desoc.fact_cal[desoc.cae_general == 5].sum(),
    casos_desoc=len(desoc), casos_muestra=len(ene))
print("TD:", round(tot["desocupados"]/tot["ft"]*100, 2), "| desocupados:", round(tot["desocupados"]))

# ---------------------------------------------------------------- Excel
wb = Workbook(); wb.remove(wb.active)
A = "Arial"
F_TIT = Font(name=A, size=12, bold=True)
F_TXT = Font(name=A, size=10)
F_HDR = Font(name=A, size=10, bold=True, color="FFFFFF")
F_TOT = Font(name=A, size=10, bold=True)
F_NOT = Font(name=A, size=9, italic=True)
FILL = PatternFill("solid", fgColor="1F3864")
BRD = Border(*[Side(style="thin")]*4)
TOPB = Border(top=Side(style="thin"))

def hoja(nombre, titulo, tab, tasa, nota=None, multiple=False, pct="% del total de desocupados"):
    ws = wb.create_sheet(nombre)
    ws["A1"] = titulo; ws["A1"].font = F_TIT
    hdr = ["Categoría","Casos muestrales","Personas desocupadas", pct]
    if tasa: hdr += ["Fuerza de trabajo","Tasa de desocupación (%)"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(3, j, h); c.font = F_HDR; c.fill = FILL; c.border = BRD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    n = len(tab); r0, r1 = 4, 3 + n; rt = r1 + 1
    den = f"$C${rt}"
    for i, (_, row) in enumerate(tab.iterrows()):
        r = r0 + i
        ws.cell(r, 1, row["Categoria"]).font = F_TXT
        ws.cell(r, 2, int(row["casos"])).font = F_TXT
        ws.cell(r, 3, int(round(row["desocupados"]))).font = F_TXT
        ws.cell(r, 4, f'=IFERROR(C{r}/{den}*100,"")').font = F_TXT
        if tasa:
            ws.cell(r, 5, int(round(row["ft"]))).font = F_TXT
            ws.cell(r, 6, f'=IFERROR(C{r}/E{r}*100,"")').font = F_TXT
    if multiple:
        # Respuesta múltiple: el denominador es el total de personas desocupadas,
        # no la suma de las filas (una persona puede aparecer en varias).
        ws.cell(rt, 1, "Total de personas desocupadas")
        ws.cell(rt, 2, tot["casos_desoc"])
        ws.cell(rt, 3, "=Resumen!B7")
        ws.cell(rt, 4, f'=IFERROR(C{rt}/{den}*100,"")')
    else:
        ws.cell(rt, 1, "Total")
        ws.cell(rt, 2, f"=SUM(B{r0}:B{r1})")
        ws.cell(rt, 3, f"=SUM(C{r0}:C{r1})")
        ws.cell(rt, 4, f"=SUM(D{r0}:D{r1})")
    if tasa:
        ws.cell(rt, 5, f"=SUM(E{r0}:E{r1})")
        ws.cell(rt, 6, f'=IFERROR(C{rt}/E{rt}*100,"")')
    for j in range(1, len(hdr)+1):
        c = ws.cell(rt, j); c.font = F_TOT; c.border = TOPB
    for r in range(r0, rt+1):
        for j in ([2,3,5] if tasa else [2,3]):
            ws.cell(r, j).number_format = "#,##0"
        for j in ([4,6] if tasa else [4]):
            ws.cell(r, j).number_format = "0.0"
    ws.column_dimensions["A"].width = 46
    for j in range(2, len(hdr)+1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    if nota:
        c = ws.cell(rt+2, 1, "Nota: " + nota); c.font = F_NOT
    ws.freeze_panes = "A4"

ws = wb.create_sheet("Notas")
notas = ["Caracterización de las personas desocupadas - ENE, trimestre móvil mayo-junio-julio 2026","",
"Fuente: Instituto Nacional de Estadísticas (INE), Encuesta Nacional de Empleo (ENE), microdatos.",
"Elaboración propia. Trimestre móvil MJJ 2026 (mes central: junio). Muestra: 97.946 personas.","",
"DEFINICIONES",
"Población en edad de trabajar (PET): personas de 15 años y más.",
"Fuerza de trabajo: personas ocupadas más personas desocupadas (ft = 1).",
"Persona desocupada (activ = 2): sin empleo, que buscó trabajo en las últimas cuatro semanas y",
"   está disponible para trabajar (marco conceptual OIT, 19a CIET).",
"Cesante (cae_general = 4): persona desocupada con experiencia laboral previa.",
"Busca trabajo por primera vez (cae_general = 5): persona desocupada sin empleo anterior.",
"Tasa de desocupación = personas desocupadas / fuerza de trabajo * 100.","",
"EXPANSIÓN",
"Todas las cifras de personas están expandidas con el factor trimestral fact_cal (proyecciones de",
"   población base Censo 2017). Los 'casos muestrales' corresponden a personas encuestadas.",
"Se recomienda cautela al interpretar categorías con menos de 50 casos muestrales.","",
"VARIABLES UTILIZADAS",
"sexo; tramo_edad; region; cine11_1d (nivel educacional); nacionalidad (ISO 3166, 152 = Chile);",
"orig1 (pueblo indígena); parentesco (jefatura de hogar); cae_general (condición de desocupación);",
"e7 (jornada buscada); e3_1 a e3_12 (gestiones de búsqueda); e6_mes/e6_ano (inicio de la búsqueda);",
"e21_mes/e21_ano (término del último empleo); e22 (motivo de término del último empleo);",
"e23 (motivo del despido, solo si e22 = 1); e24 (motivo de la renuncia, solo si e22 = 2).","",
"SUPUESTOS",
"La duración de la búsqueda y de la cesantía se calcula como diferencia en meses entre el mes de la",
"   entrevista (mes_encuesta) y la fecha declarada; los valores negativos se truncan en 0 y los",
"   códigos 88 / 99 / 9999 se clasifican como 'No declarado'.",
"Las etiquetas de cine11_1d se validaron cruzando esa variable con 'nivel' y 'termino_nivel'.",
"Las glosas de e22 (pregunta E19) para los códigos 3 a 9, 88 y 99 provienen del libro de códigos del INE.",
"   Los códigos 1 y 2 se etiquetan como despido y renuncia porque son los que habilitan las preguntas",
"   e23 (motivo de despido) y e24 (motivo de renuncia).",
"Las glosas de e23 y e24 provienen del libro de códigos del INE; ambas preguntas se incorporaron a",
"   partir del trimestre JJA 2020 por recomendación de la OIT (INE, 2020c).","",
"",
"ESTRUCTURA PANEL (hojas Panel_*)",
"La ENE es una encuesta de panel rotativo: cada persona seleccionada es entrevistada durante tres",
"   meses consecutivos, sale de la muestra y vuelve a ser entrevistada tres meses más, 12 meses",
"   después de la primera visita. Eso permite enlazar MJJ 2026 con MJJ 2025 por el identificador de",
"   persona idrph, siempre en el mismo mes calendario (mayo con mayo, junio con junio, julio con julio).",
f"Pares enlazados y validados: {pan['pares']:,}".replace(",", ".") + " personas entrevistadas en ambos trimestres móviles.",
"Validación del enlace: se exige mismo sexo y una diferencia de edad de 0 a 2 años entre ambas olas;",
"   los pares que no cumplen se descartan.",
f"Cobertura: {pan['casos']:,} de las {tot['casos_desoc']:,} personas desocupadas de MJJ 2026 ".replace(",", ".") +
f"({pan['casos']/tot['casos_desoc']*100:.1f}%) tienen dato de 2025.",
"ADVERTENCIA: la submuestra panel no tiene diseño muestral ni factores de expansión propios. Las cifras",
"   expandidas de las hojas Panel_* se calculan con fact_cal de 2026 y son solo indicativas; la lectura",
"   recomendada es la distribución porcentual, y no deben presentarse como estimaciones oficiales del INE.",
"Variables de 2025 utilizadas: activ y cae_general (situación laboral); r_p_rev4cl_caenes (rama CAENES);",
"   categoria_ocupacion; ocup_form (formalidad); sector (sector institucional); b8 y b9 (contrato).",
"",
"Script que genera este archivo: ene_empleo_microdatos/desocupados_mjj2026.R"]
for i, t in enumerate(notas, 1):
    c = ws.cell(i, 1, t); c.font = F_TXT
ws["A1"].font = F_TIT
ws.column_dimensions["A"].width = 105

ws = wb.create_sheet("Resumen")
ws["A1"] = "Resumen - personas desocupadas, ENE MJJ 2026"; ws["A1"].font = F_TIT
for j, h in enumerate(["Indicador","Valor"], 1):
    c = ws.cell(3, j, h); c.font = F_HDR; c.fill = FILL; c.border = BRD
    c.alignment = Alignment(horizontal="center", vertical="center")
res = [("Población en edad de trabajar (PET)", round(tot["pet"])),
       ("Fuerza de trabajo", round(tot["ft"])),
       ("Personas ocupadas", round(tot["ocupados"])),
       ("Personas desocupadas", round(tot["desocupados"])),
       ("   Cesantes", round(tot["cesantes"])),
       ("   Buscan trabajo por primera vez", round(tot["primera"])),
       ("Tasa de desocupación (%)", "=B7/B5*100"),
       ("Tasa de cesantía (%)", "=B8/B5*100"),
       ("Casos muestrales (total encuesta)", tot["casos_muestra"]),
       ("Casos muestrales (personas desocupadas)", tot["casos_desoc"])]
for i, (k, v) in enumerate(res, 4):
    ws.cell(i, 1, k).font = F_TXT
    c = ws.cell(i, 2, v); c.font = F_TXT
    c.number_format = "0.0" if i in (10, 11) else "#,##0"
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 18

hoja("Sexo", "Personas desocupadas según sexo - MJJ 2026", t_sexo, True)
hoja("Edad", "Personas desocupadas según tramo de edad - MJJ 2026", t_edad, True)
hoja("Region", "Personas desocupadas según región - MJJ 2026", t_region, True)
hoja("Provincia", "Personas desocupadas según provincia - MJJ 2026", t_provincia, True,
     nota='Nota: la ENE tiene representatividad regional, no provincial. Estas cifras son elaboración propia y no constituyen estimaciones oficiales del INE: en varias provincias el número de casos muestrales de personas desocupadas es muy bajo y la tasa resultante tiene un error muestral alto. Las provincias de Isla de Pascua y Palena no aparecen porque no forman parte de la muestra del trimestre. El nombre de cada provincia va acompañado de su región.')
hoja("Educacion", "Personas desocupadas según nivel educacional alcanzado - MJJ 2026", t_educ, True,
     "Clasificación CINE 2011 (variable cine11_1d de la ENE).")
hoja("Nacionalidad", "Personas desocupadas según nacionalidad - MJJ 2026", t_nac, True)
hoja("Pueblo_indigena", "Personas desocupadas según pertenencia a pueblo indígena u originario - MJJ 2026", t_indig, True)
hoja("Jefatura_hogar", "Personas desocupadas según jefatura de hogar - MJJ 2026", t_jefe, True)
hoja("Condicion", "Personas desocupadas según condición: cesante o busca trabajo por primera vez - MJJ 2026", t_cond, False)
hoja("Duracion_busqueda", "Personas desocupadas según duración de la búsqueda de empleo - MJJ 2026", t_busq, False,
     "Meses transcurridos desde que declara haber comenzado a buscar trabajo (e6).")
hoja("Duracion_cesantia", "Personas cesantes según tiempo transcurrido desde su último empleo - MJJ 2026", t_cesa, False,
     "Universo: solo personas cesantes (cae_general = 4). Base: e21.")
hoja("Motivo_termino", "Personas cesantes según motivo de término de su último empleo - MJJ 2026", t_motivo, False,
     "Universo: solo personas cesantes (e22, pregunta E19 del cuestionario). Quienes responden 'despido' "
     "detallan el motivo en e23 y quienes responden 'renuncia' lo detallan en e24.")
hoja("Motivo_despido", "Personas cesantes despedidas según motivo del despido - MJJ 2026", t_despido, False,
     "Universo: personas cesantes cuyo último empleo terminó por despido (e22 = 1). Variable e23. "
     "Pregunta incorporada a partir del trimestre JJA 2020, recomendada por la OIT en el contexto del "
     "impacto de la pandemia de COVID-19 sobre la población no ocupada (INE, 2020c).")
hoja("Motivo_renuncia", "Personas cesantes que renunciaron según motivo de la renuncia - MJJ 2026", t_renuncia, False,
     "Universo: personas cesantes cuyo último empleo terminó por renuncia (e22 = 2). Variable e24. "
     "Pregunta incorporada a partir del trimestre JJA 2020, recomendada por la OIT en el contexto del "
     "impacto de la pandemia de COVID-19 sobre la población no ocupada (INE, 2020c).")
hoja("Jornada_buscada", "Personas desocupadas según tipo de jornada buscada - MJJ 2026", t_jorn, False)
hoja("Metodos_busqueda", "Personas desocupadas según gestiones de búsqueda realizadas - MJJ 2026", t_metodos, False,
     "Respuesta múltiple: una persona puede declarar varias gestiones, por lo que los porcentajes no suman 100. "
     "El denominador es el total de personas desocupadas.", multiple=True)

def mil(x):  # separador de miles con punto (formato chileno)
    return f"{int(x):,}".replace(",", ".")

NOTA_PANEL = (
    f"Panel ENE: {mil(pan['casos'])} de las {mil(tot['casos_desoc'])} personas desocupadas de MJJ 2026 "
    f"({pan['casos']/tot['casos_desoc']*100:.1f}%) también fueron entrevistadas en MJJ 2025. "
    "La submuestra panel no tiene diseño ni factores de expansión propios, por lo que las cifras "
    "expandidas son indicativas y no constituyen estimaciones oficiales; el porcentaje es la lectura "
    "recomendada.")

hoja("Panel_situacion_2025",
     "¿En qué situación estaban en MJJ 2025 las personas desocupadas de MJJ 2026?",
     t_p_sit, False, NOTA_PANEL + " Fuerza de trabajo potencial: personas fuera de la fuerza de "
     "trabajo disponibles para trabajar o que buscaron sin estar disponibles (cae_general 6 a 8).",
     pct="% de las personas enlazadas")
hoja("Panel_rama_2025",
     "Rama de actividad en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
     t_p_rama, False, "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 "
     f"({pan['ocup']} casos). Clasificación CAENES (CIIU Rev. 4 CL), variable r_p_rev4cl_caenes de 2025. "
     + NOTA_PANEL, pct="% de quienes estaban ocupados/as")
hoja("Panel_contrato_2025",
     "Tipo de contrato en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
     t_p_contrato, False, "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 "
     f"({pan['ocup']} casos). Construido con b8 (¿tiene contrato escrito?) y b9 (duración del contrato "
     "o acuerdo). No aplica a empleadores, cuenta propia ni familiares no remunerados. " + NOTA_PANEL,
     pct="% de quienes estaban ocupados/as")
hoja("Panel_categoria_2025",
     "Categoría ocupacional en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
     t_p_categ, False, "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 "
     f"({pan['ocup']} casos). " + NOTA_PANEL, pct="% de quienes estaban ocupados/as")
hoja("Panel_formalidad_2025",
     "Formalidad del empleo en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
     t_p_formal, False, "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 "
     f"({pan['ocup']} casos). Variable ocup_form (condición de formalidad del empleo, marco OIT 17a CIET). "
     + NOTA_PANEL, pct="% de quienes estaban ocupados/as")
hoja("Panel_sector_inst_2025",
     "Sector institucional del empleo en MJJ 2025 de quienes están desocupados/as en MJJ 2026",
     t_p_sectinst, False, "Universo: personas desocupadas en MJJ 2026 que estaban ocupadas en MJJ 2025 "
     f"({pan['ocup']} casos). Variable sector: unidad productiva formal, informal u hogares. "
     + NOTA_PANEL, pct="% de quienes estaban ocupados/as")

out = "desocupados_mjj2026.xlsx"
wb.save(out)
print("OK", out)
