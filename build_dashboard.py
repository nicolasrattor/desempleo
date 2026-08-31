# -*- coding: utf-8 -*-
"""Construye el dashboard HTML "Desempleo en Chile, mas alla del 9,5%".

Lee dos fuentes:
  1. desocupados_mjj2026.xlsx  -> los 21 tabulados transversales y de panel
  2. input/data/ano=*/mes_central=*  -> la serie de trimestres moviles 2010-2026

y escribe un unico archivo HTML autocontenido (los datos van embebidos como JSON;
lo unico externo es Chart.js desde CDN).

Uso, desde la raiz del repositorio:  python3 ene_empleo_microdatos/build_dashboard.py
"""
import glob, json, re
import numpy as np
import openpyxl
import pandas as pd
import pyarrow.parquet as pq

import os

# Los microdatos de la ENE no viajan en el repositorio (pesan ~1,2 GB). Se esperan en ./data
# con la particion original ano=/mes_central=, o en la ruta que indique la variable ENE_DATA.
BASE = os.environ.get("ENE_DATA", "data")
XL   = "desocupados_mjj2026.xlsx"
OUT  = "dashboard_desempleo.html"
REPO = "https://github.com/nicolasrattor/desempleo"
TOTAL_DESOCUPADOS = 981044   # base fija para la hoja de respuesta multiple

# ---------------------------------------------------------------- serie 2010-2026
def serie_historica():
    """Tasa de desocupacion y N de desocupados de cada trimestre movil disponible.

    No se usa la variable `ft`: solo existe en los archivos recientes. La fuerza de
    trabajo se reconstruye como ocupados + desocupados (activ 1 y 2), que es
    exactamente su definicion.
    """
    filas = []
    for f in sorted(glob.glob(f"{BASE}/ano=*/mes_central=*/part-0.parquet")):
        m = re.search(r"ano=(\d+)/mes_central=(\d+)", f)
        ano, mes = int(m.group(1)), int(m.group(2))
        t = pq.read_table(f, columns=["activ", "fact_cal"]).to_pandas()
        d = t.loc[t.activ == 2, "fact_cal"].sum()
        o = t.loc[t.activ == 1, "fact_cal"].sum()
        filas.append(dict(ano=ano, mes_central=mes, desocupados=d, ft=o + d,
                          td=d / (o + d) * 100))
    return pd.DataFrame(filas).sort_values(["ano", "mes_central"])

# ------------------------------------------------- serie de larga duracion 2020-2026
def _peso_12m(df, pref):
    """Probabilidad de que el episodio lleve 12 meses o mas, fila a fila.

    La persona declara mes y ano de inicio (de la busqueda, e6; o del termino del
    ultimo empleo, e21). El mes falta en una fraccion no despreciable de los casos
    -y esa fraccion cambia mucho entre trimestres-, asi que en vez de descartarlos
    se resuelve por la brecha de anos, que casi siempre basta:

      * mismo ano de la encuesta            -> menos de 12 meses (peso 0)
      * dos anos o mas de diferencia        -> 12 meses o mas     (peso 1)
      * exactamente un ano de diferencia    -> depende del mes: si se declaro, se
        compara contra el mes de la encuesta; si no, se reparte suponiendo el mes
        uniforme dentro del ano (peso = mes_encuesta/12).

    Solo el ultimo caso es imputado y pesa ~1% de la muestra, salvo entre diciembre
    de 2023 y diciembre de 2024, cuando una ola de no respuesta del mes lo lleva
    hasta el 50% de las personas cesantes. Esa ventana se marca en el grafico.
    """
    a = df[f"{pref}_ano"].where((df[f"{pref}_ano"] >= 1900) &
                                (df[f"{pref}_ano"] <= df.ano_encuesta))
    m = df[f"{pref}_mes"].where(df[f"{pref}_mes"].between(1, 12))
    gap = df.ano_encuesta - a
    w = pd.Series(np.nan, index=df.index)
    w[gap == 0]  = 0.0
    w[gap >= 2]  = 1.0
    con = (gap == 1) & m.notna();  w[con] = (m <= df.mes_encuesta)[con].astype(float)
    sin = (gap == 1) & m.isna();   w[sin] = (df.mes_encuesta / 12)[sin]
    return w, sin

def serie_duracion(desde=2020):
    """% de desocupados que llevan 12 meses o mas buscando y % de cesantes que
    llevan 12 meses o mas sin empleo, trimestre a trimestre."""
    fs = [f for f in glob.glob(f"{BASE}/ano=*/mes_central=*/part-0.parquet")
          if int(re.search(r"ano=(\d+)", f).group(1)) >= desde]
    fs.sort(key=lambda f: [int(x) for x in
                           re.search(r"ano=(\d+)/mes_central=(\d+)", f).groups()])
    filas = []
    for f in fs:
        ano, mes = [int(x) for x in re.search(r"ano=(\d+)/mes_central=(\d+)", f).groups()]
        cols = ["activ", "cae_general", "fact_cal", "mes_encuesta", "ano_encuesta",
                "e6_mes", "e6_ano"]
        # e21 (termino del ultimo empleo) recien aparece con el cuestionario de julio de 2020
        hay21 = "e21_mes" in pq.read_schema(f).names
        if hay21:
            cols += ["e21_mes", "e21_ano"]
        t = pq.read_table(f, columns=cols).to_pandas()
        d = t[t.activ == 2].copy()
        fila = dict(a=ano, m=mes)
        pares = [("bus", d, "e6")] + ([("ces", d[d.cae_general == 4].copy(), "e21")] if hay21 else [])
        for k, base, pref in pares:
            w, sin = _peso_12m(base, pref)
            ok = w.notna()
            fila[k] = round((base.fact_cal[ok] * w[ok]).sum() / base.fact_cal[ok].sum() * 100, 2)
            fila[k + "imp"] = round(base.fact_cal[sin].sum() / base.fact_cal.sum() * 100, 1)
            fila[k + "n"] = int(round(w[ok].sum()))
        filas.append(fila)
    return filas

# ---------------------------------------------------------------- lectura del Excel
def leer_excel():
    wb = openpyxl.load_workbook(XL)
    out = {}
    for n in wb.sheetnames:
        if n == "Notas":
            continue
        ws = wb[n]
        if n == "Resumen":
            d = {}
            for r in ws.iter_rows(min_row=4, values_only=True):
                if r[0] and not isinstance(r[1], str):
                    d[r[0].strip()] = r[1]
            out["Resumen"] = d
            continue
        hdr = [c for c in next(ws.iter_rows(min_row=3, max_row=3, values_only=True)) if c]
        rows, nota = [], None
        for r in ws.iter_rows(min_row=4, values_only=True):
            if r[0] is None:
                continue
            if str(r[0]).startswith("Nota"):
                nota = r[0]; continue
            if str(r[0]).startswith("Total"):
                continue
            rows.append(dict(cat=r[0], casos=r[1], n=r[2],
                             ft=r[4] if len(r) > 4 and isinstance(r[4], (int, float)) else None))
        # Los porcentajes y tasas viven como formulas en el Excel; aqui se recalculan
        # desde los valores para no depender de un motor de calculo.
        tot = TOTAL_DESOCUPADOS if n == "Metodos_busqueda" else sum(x["n"] for x in rows)
        for x in rows:
            x["pct"]  = round(x["n"] / tot * 100, 1) if tot else None
            x["tasa"] = round(x["n"] / x["ft"] * 100, 1) if x["ft"] else None
        out[n] = dict(titulo=ws["A1"].value, hdr=hdr, rows=rows, total=tot, nota=nota)
    return out

MES = {1:"DEF",2:"EFM",3:"FMA",4:"MAM",5:"AMJ",6:"MJJ",
       7:"JJA",8:"JAS",9:"ASO",10:"SON",11:"OND",12:"NDE"}

D = leer_excel()
s = serie_historica()
D["serie"] = [dict(a=int(r.ano), m=int(r.mes_central),
                   lab=f"{MES[int(r.mes_central)]} {int(r.ano)}",
                   td=round(r.td, 2), d=int(round(r.desocupados)))
              for r in s.itertuples()]
D["serie_dur"] = [dict(lab=f"{MES[x['m']]} {x['a']}", **x) for x in serie_duracion()]
DATA_JS = json.dumps(D, ensure_ascii=False)

HTML = """<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Desempleo en Chile, más allá del 9,5%</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{
  --ink:#101418; --ink2:#3c4650; --muted:#7b8894; --line:#e3e7ec;
  --bg:#f6f7f9; --card:#ffffff;
  --a1:#0b3d5c; --a2:#e2603c; --a3:#2f8f83; --a4:#c9a227; --a5:#7d5ba6;
  --serif:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;
  --sans:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,Roboto,Helvetica,Arial,sans-serif;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--sans);
     -webkit-font-smoothing:antialiased;line-height:1.5}
.wrap{max-width:1180px;margin:0 auto;padding:0 24px 80px}

/* ---------- portada ---------- */
.hero{background:linear-gradient(160deg,#0b3d5c 0%,#12293b 60%,#0d1a26 100%);color:#fff;
      padding:64px 0 0;margin-bottom:36px}
.hero .wrap{padding-bottom:0}
.kicker{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:#8fb6cd;margin:0 0 14px}
h1{font-family:var(--serif);font-weight:600;font-size:clamp(34px,5.2vw,60px);line-height:1.05;
   margin:0 0 14px;letter-spacing:-.5px}
h1 em{font-style:normal;color:#f0a58c}
.sub{max-width:660px;color:#c3d3de;font-size:16.5px;margin:0 0 34px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1px;
      background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.14);border-radius:10px;overflow:hidden}
.kpi{background:#0e3247;padding:16px 18px}
.kpi b{display:block;font-family:var(--serif);font-size:30px;line-height:1.1;letter-spacing:-.5px}
.kpi span{display:block;font-size:11.5px;color:#9fbdd0;margin-top:5px;letter-spacing:.03em}
.herochart{background:var(--card);border-radius:12px;margin-top:32px;padding:22px 22px 14px;
           box-shadow:0 18px 40px -22px rgba(0,0,0,.55);transform:translateY(36px)}
.herochart h2{margin:0 0 2px;font-family:var(--serif);font-size:20px;color:var(--ink)}
.herochart p.lead{margin:0 0 14px;font-size:13px;color:var(--muted)}
.spacer{height:60px}

/* ---------- estructura ---------- */
section{margin:56px 0 0}
.shead{border-top:2px solid var(--ink);padding-top:12px;margin-bottom:22px}
.shead .num{font-size:11px;letter-spacing:.18em;color:var(--a2);font-weight:700}
.shead h2{font-family:var(--serif);font-size:29px;margin:4px 0 6px;letter-spacing:-.3px}
.shead p{margin:0;color:var(--ink2);font-size:15px;max-width:820px}
.grid{display:grid;gap:18px}
.g2{grid-template-columns:repeat(auto-fit,minmax(330px,1fr))}
.card{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:18px 20px 16px}
.card.span{grid-column:1/-1}
.card h3{margin:0 0 3px;font-size:14.5px;font-weight:650;letter-spacing:-.1px}
.card .hint{margin:0 0 14px;font-size:12px;color:var(--muted)}
.cw{position:relative}
canvas{max-width:100%}

/* ---------- tablas / barras html ---------- */
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-weight:600;font-size:11px;letter-spacing:.06em;text-transform:uppercase;
   color:var(--muted);border-bottom:1px solid var(--line);padding:0 8px 7px}
td{padding:7px 8px;border-bottom:1px solid #f0f2f5;vertical-align:middle}
td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr:last-child td{border-bottom:none}
.bar{height:7px;border-radius:4px;background:var(--a1);min-width:2px}
.barwrap{background:#eef1f4;border-radius:4px;width:100%}

.callout{background:#fff;border-left:3px solid var(--a2);padding:16px 20px;border-radius:0 10px 10px 0;
         font-size:15px;color:var(--ink2);margin:0 0 20px}
.callout b{color:var(--ink)}
.notes{margin-top:56px;background:#eef1f4;border-radius:11px;padding:22px 24px;font-size:13px;color:var(--ink2)}
.notes h3{margin:0 0 10px;font-family:var(--serif);font-size:19px;color:var(--ink)}
.notes p{margin:0 0 10px}
.tabs{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}
.tab{font-size:12px;padding:5px 12px;border-radius:20px;border:1px solid var(--line);background:#fff;
     cursor:pointer;color:var(--ink2)}
.tab.on{background:var(--ink);color:#fff;border-color:var(--ink)}
.tab.kast.on{background:var(--a2);border-color:var(--a2)}
.caption{margin:10px 0 0;font-size:12px;color:var(--muted);border-left:2px solid var(--line);padding-left:10px}
footer{margin-top:40px;font-size:12px;color:var(--muted);text-align:center}

/* ---------- enlace al repositorio ---------- */
.gh{display:inline-flex;align-items:center;gap:8px;text-decoration:none;font-size:12.5px;
    padding:7px 13px 7px 11px;border-radius:20px;transition:background .15s,border-color .15s}
.gh svg{width:17px;height:17px;flex:none}
.hero .gh{color:#c3d3de;border:1px solid rgba(255,255,255,.22);background:rgba(255,255,255,.06);
          margin-bottom:26px}
.hero .gh:hover{background:rgba(255,255,255,.14);border-color:rgba(255,255,255,.4);color:#fff}
.hero .gh svg{fill:#fff}
footer .gh{color:var(--ink2);border:1px solid var(--line);background:#fff;margin-bottom:12px}
footer .gh:hover{border-color:var(--ink);color:var(--ink)}
footer .gh svg{fill:var(--ink)}
.gh b{font-weight:600}

/* ---------- panel de dos graficos dentro de una tarjeta ---------- */
.duo{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:20px}
.duo .mini{margin:0 0 10px;font-size:11.5px;color:var(--muted)}

/* ---------- pestanas de nivel 1 ---------- */
#nav{position:sticky;top:0;z-index:50;background:rgba(246,247,249,.94);
     backdrop-filter:saturate(160%) blur(8px);border-bottom:1px solid var(--line);
     margin-bottom:8px}
#nav .wrap{padding:0 24px;display:flex;gap:2px;overflow-x:auto;scrollbar-width:none}
#nav .wrap::-webkit-scrollbar{display:none}
#nav button{appearance:none;background:none;border:0;cursor:pointer;white-space:nowrap;
     font-family:inherit;font-size:13.5px;color:var(--muted);padding:15px 14px 13px;
     border-bottom:2.5px solid transparent;display:flex;align-items:baseline;gap:7px}
#nav button:hover{color:var(--ink)}
#nav button.on{color:var(--ink);font-weight:650;border-bottom-color:var(--a2)}
#nav button i{font-style:normal;font-size:10.5px;letter-spacing:.1em;color:var(--a2);font-weight:700}
section{margin:34px 0 0}
section[hidden]{display:none}
@media(max-width:640px){.herochart{transform:none;margin-top:24px}.spacer{height:0}
  #nav button{font-size:12.5px;padding:13px 10px 11px}}
</style>

<!-- El logo de GitHub va inline como SVG para que el archivo siga siendo autocontenido. -->
<template id="ghlink">
  <a class="gh" href="__REPO__" target="_blank" rel="noopener" title="Ver el código y los datos en GitHub">
    <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"/></svg>
    <span><b>nicolasrattor/desempleo</b> · código y metodología</span>
  </a>
</template>

<div class="hero">
  <div class="wrap">
    <p class="kicker">Encuesta Nacional de Empleo · INE Chile · Trimestre móvil mayo–julio 2026</p>
    <h1>Desempleo en Chile,<br>más allá del <em>9,5%</em></h1>
    <p class="sub">981 mil personas desocupadas detrás de una sola cifra. Quiénes son, cuánto llevan buscando,
       por qué dejaron su último empleo y —gracias a la estructura de panel de la ENE— dónde estaban hace exactamente un año.</p>
    <div id="gh_hero"></div>
    <div class="kpis" id="kpis"></div>
    <div class="herochart">
      <h2>Dieciséis años de desempleo, mes a mes</h2>
      <p class="lead">Tasa de desocupación (línea) y número de personas desocupadas (área), trimestres móviles desde 2010 hasta MJJ 2026.</p>
      <div class="tabs" id="rangeTabs"></div>
      <div class="cw" style="height:330px"><canvas id="serie"></canvas></div>
      <p class="caption" id="serieCaption"></p>
    </div>
  </div>
</div>
<div class="spacer"></div>

<nav id="nav"><div class="wrap" id="navtabs"></div></nav>

<div class="wrap">

<section>
  <div class="shead"><div class="num">01</div><h2>Quiénes están desocupados</h2>
  <p>La tasa agregada esconde diferencias grandes: por edad va de 24% a menos de 6%, y entre regiones hay más de siete puntos de distancia.</p></div>
  <div class="callout">De cada 100 personas desocupadas, <b>91 son cesantes</b> —perdieron o dejaron un empleo que ya tenían— y solo
    <b>9 buscan trabajo por primera vez</b>. El desempleo chileno es, sobre todo, un problema de reinserción.</div>
  <div class="grid g2">
    <div class="card"><h3>Sexo</h3><p class="hint">Tasa de desocupación, %</p><div class="cw" style="height:170px"><canvas id="c_sexo"></canvas></div></div>
    <div class="card"><h3>Tramo de edad · tasa</h3><p class="hint">Tasa de desocupación de cada tramo, %</p><div class="cw" style="height:230px"><canvas id="c_edad"></canvas></div></div>
    <div class="card"><h3>Tramo de edad · distribución</h3><p class="hint">Cómo se reparten las 981 mil personas desocupadas, %</p><div class="cw" style="height:230px"><canvas id="c_edad_d"></canvas></div></div>
    <div class="card span"><h3>Región</h3><p class="hint">Tasa de desocupación, % — ordenada de mayor a menor</p><div class="cw" style="height:300px"><canvas id="c_region"></canvas></div></div>
    <div class="card span"><h3>Provincia</h3>
      <p class="hint">Tasa de desocupación, % — la ENE tiene representatividad regional, no provincial: leer con cautela</p>
      <div class="tabs" id="provTabs">
        <div class="tab on" data-min="50">Solo provincias con 50+ casos</div>
        <div class="tab" data-min="0">Las 52 provincias</div>
      </div>
      <div id="t_prov"></div>
      <p class="hint" style="margin:12px 0 0">Las provincias marcadas con ▵ tienen menos de 50 casos muestrales de personas
        desocupadas: su tasa tiene un error muestral alto y no debe leerse como una estimación puntual. Isla de Pascua y
        Palena no aparecen porque no forman parte de la muestra del trimestre.</p>
    </div>
    <div class="card"><h3>Nivel educacional</h3><div id="t_educ"></div></div>
    <div class="card"><h3>Condición: cesantes y quienes buscan por primera vez</h3><div id="t_cond"></div></div>
    <div class="card"><h3>Nacionalidad</h3><div id="t_nac"></div></div>
    <div class="card"><h3>Jefatura de hogar y pueblos indígenas</h3><div id="t_jef"></div><div style="height:10px"></div><div id="t_pi"></div></div>
  </div>
</section>

<section>
  <div class="shead"><div class="num">02</div><h2>Cuánto llevan buscando</h2>
  <p>La duración del episodio separa la rotación normal del mercado laboral del desempleo que se enquista.</p></div>
  <div class="grid g2">
    <div class="card span"><h3>El desempleo de larga duración, trimestre a trimestre</h3>
      <p class="hint">Dos indicadores desde 2020: personas desocupadas que llevan 12 meses o más buscando trabajo
        (% del total de desocupados) y personas cesantes que llevan 12 meses o más sin empleo (% del total de cesantes).</p>
      <div class="cw" style="height:300px"><canvas id="c_larga"></canvas></div>
      <p class="caption" id="largaCaption"></p>
    </div>
    <div class="card"><h3>Duración de la búsqueda de empleo</h3><p class="hint">% del total de personas desocupadas</p><div class="cw" style="height:230px"><canvas id="c_dbus"></canvas></div></div>
    <div class="card"><h3>Duración de la cesantía</h3><p class="hint">Tiempo sin empleo, solo personas cesantes</p><div class="cw" style="height:230px"><canvas id="c_dces"></canvas></div></div>
    <div class="card"><h3>Métodos de búsqueda utilizados</h3><p class="hint">Respuesta múltiple: los porcentajes no suman 100</p><div id="t_met"></div></div>
    <div class="card"><h3>Jornada buscada</h3><div id="t_jor"></div></div>
  </div>
</section>

<section data-tab="Por qué terminó el empleo">
  <div class="shead"><div class="num">03</div><h2>Por qué terminó el último empleo</h2>
  <p>Entre las personas cesantes, seis de cada diez salieron por el fin de un contrato, una faena o una temporada: el desempleo llega, la mayoría de las veces, por la vía del empleo temporal.</p></div>
  <div class="grid g2">
    <div class="card span"><h3>Motivo de término del último empleo</h3><p class="hint">% de las personas cesantes</p><div id="t_mot"></div></div>
    <div class="card"><h3>Si fue despido, ¿por qué?</h3><div id="t_desp"></div></div>
    <div class="card"><h3>Si fue renuncia, ¿por qué?</h3><div id="t_ren"></div></div>
  </div>
</section>

<section data-tab="De dónde vienen">
  <div class="shead"><div class="num">04</div><h2>De dónde vienen: el panel MJJ 2025 → MJJ 2026</h2>
  <p>La ENE reentrevista a cada persona doce meses después. Enlazando ambas olas se recupera algo que el corte transversal no puede
     mostrar: la situación laboral de las personas hoy desocupadas exactamente un año antes.</p></div>
  <div class="callout">Menos de la mitad —<b>46%</b>— de quienes hoy están desocupados estaba ocupado hace un año.
    Un <b>23%</b> ya estaba desocupado y un <b>31%</b> estaba fuera de la fuerza de trabajo.
    El desempleo actual no es solo empleo perdido: también es desempleo que persiste y entrada al mercado laboral.</div>
  <div class="grid g2">
    <div class="card span"><h3>¿En qué situación estaban en MJJ 2025?</h3><p class="hint">% de las 1.086 personas desocupadas de 2026 enlazadas con 2025</p><div class="cw" style="height:210px"><canvas id="c_psit"></canvas></div></div>
    <div class="card span"><h3>¿En qué sector trabajaban quienes sí estaban ocupados?</h3><p class="hint">Rama de actividad CAENES (CIIU Rev. 4 CL), 491 casos</p><div class="cw" style="height:340px"><canvas id="c_prama"></canvas></div></div>
    <div class="card"><h3>Tipo de contrato en 2025</h3><div id="t_pcon"></div></div>
    <div class="card"><h3>Categoría ocupacional en 2025</h3><div id="t_pcat"></div></div>
    <div class="card"><h3>Formalidad del empleo</h3><div id="t_pfor"></div><p class="hint" style="margin-top:10px">Referencia: la informalidad del total de personas ocupadas del país bordea el 28–29%.</p></div>
    <div class="card"><h3>Sector institucional</h3><div id="t_psec"></div></div>
  </div>
</section>

<section data-tab="Nota metodológica"><div class="notes" id="notas" style="margin-top:0"></div></section>
<footer><div id="gh_foot"></div>Elaboración propia a partir de microdatos de la Encuesta Nacional de Empleo, Instituto Nacional de Estadísticas de Chile.<br>
Cifras expandidas con el factor trimestral <code>fact_cal</code> (proyecciones de población base Censo 2017).</footer>
</div>

<script>
const D = __DATA__;
const F = n => new Intl.NumberFormat('es-CL').format(Math.round(n));
const P = n => (n==null?'—':n.toLocaleString('es-CL',{minimumFractionDigits:1,maximumFractionDigits:1}));
const CSS = k => getComputedStyle(document.documentElement).getPropertyValue(k).trim();
const A1=CSS('--a1'),A2=CSS('--a2'),A3=CSS('--a3'),A4=CSS('--a4'),A5=CSS('--a5'),INK=CSS('--ink'),MUT=CSS('--muted'),LINE=CSS('--line');
Chart.defaults.font.family = CSS('--sans');
Chart.defaults.font.size = 11.5;
Chart.defaults.color = MUT;
Chart.defaults.plugins.legend.display = false;
Chart.defaults.maintainAspectRatio = false;

/* ---- KPIs ---- */
const R = D.Resumen;
const kpi = [
  ['9,5%','Tasa de desocupación MJJ 2026'],
  [F(R['Personas desocupadas']),'Personas desocupadas'],
  [F(R['Cesantes']),'Cesantes (91,1%)'],
  [F(R['Buscan trabajo por primera vez']),'Buscan trabajo por 1ª vez'],
  [F(R['Fuerza de trabajo']),'Fuerza de trabajo'],
  ['4.412','Casos muestrales']
];
document.getElementById('kpis').innerHTML = kpi.map(k=>`<div class="kpi"><b>${k[0]}</b><span>${k[1]}</span></div>`).join('');

/* ---- helpers de tabla ---- */
function tabla(el, sheet, opts={}){
  const S = D[sheet]; if(!S) return;
  const rows = (opts.filter? S.rows.filter(opts.filter): S.rows).slice();
  if(opts.sort) rows.sort((a,b)=>b.pct-a.pct);
  const max = Math.max(...rows.map(r=>opts.val? opts.val(r): r.pct));
  const col = opts.color||A1;
  const valh = opts.head || '%';
  document.getElementById(el).innerHTML =
   `<table><thead><tr><th>${opts.cath||'Categoría'}</th><th style="width:34%"></th><th class="n">${valh}</th><th class="n">Casos</th></tr></thead><tbody>`+
   rows.map(r=>{const v = opts.val? opts.val(r): r.pct;
     return `<tr><td>${r.cat}</td>
       <td><div class="barwrap"><div class="bar" style="width:${Math.max(2,v/max*100)}%;background:${col}"></div></div></td>
       <td class="n"><b>${P(v)}</b></td><td class="n" style="color:var(--muted)">${F(r.casos)}</td></tr>`}).join('')+
   `</tbody></table>`;
}
const baseOpts = {responsive:true, plugins:{tooltip:{backgroundColor:'#101418',padding:10,cornerRadius:6,displayColors:false}}};
function barH(id, labels, vals, color, suf='%'){
  new Chart(document.getElementById(id),{type:'bar',data:{labels,datasets:[{data:vals,backgroundColor:color,borderRadius:3,barPercentage:.75}]},
   options:{...baseOpts,indexAxis:'y',
    scales:{x:{grid:{color:LINE,drawTicks:false},border:{display:false},ticks:{callback:v=>v+suf}},
            y:{grid:{display:false},border:{display:false},ticks:{autoSkip:false,font:{size:11}}}},
    plugins:{...baseOpts.plugins,tooltip:{...baseOpts.plugins.tooltip,callbacks:{label:c=>P(c.raw)+suf}}}}});
}

/* ---- serie histórica ---- */
const S = D.serie;
let chartSerie=null;
function drawSerie(r){
  // El corte se expresa como [año, mes_central] para poder aislar periodos de pocos meses.
  const s = S.filter(x => x.a*100 + x.m >= r.from[0]*100 + r.from[1]);
  const pocos = s.length <= 12;   // con pocos puntos conviene marcar cada trimestre
  if(chartSerie) chartSerie.destroy();
  const ctx = document.getElementById('serie').getContext('2d');
  const g = ctx.createLinearGradient(0,0,0,300);
  g.addColorStop(0,'rgba(11,61,92,.20)'); g.addColorStop(1,'rgba(11,61,92,.02)');
  chartSerie = new Chart(ctx,{data:{labels:s.map(x=>x.lab),datasets:[
    {type:'line',label:'Personas desocupadas',data:s.map(x=>x.d),yAxisID:'y1',borderColor:'rgba(11,61,92,.35)',
     backgroundColor:g,fill:true,pointRadius:0,borderWidth:1,tension:.25},
    {type:'line',label:'Tasa de desocupación',data:s.map(x=>x.td),yAxisID:'y',borderColor:A2,
     borderWidth:2.4,pointRadius:pocos?3.5:0,pointBackgroundColor:A2,pointHoverRadius:5,tension:pocos?.15:.25}
  ]},options:{...baseOpts,interaction:{mode:'index',intersect:false},
   scales:{
     x:{grid:{display:false},border:{color:LINE},
        ticks:{maxTicksLimit:pocos?12:12,autoSkip:!pocos,maxRotation:0}},
     y:{position:'left',grid:{color:LINE},border:{display:false},
        // En rangos cortos el eje se autoescala y saca ticks con dos o tres decimales: un decimal basta.
        ticks:{callback:v=>(pocos? P(v): v)+'%'},
        title:{display:true,text:'Tasa de desocupación',color:A2,font:{size:11,weight:'600'}}},
     y1:{position:'right',grid:{display:false},border:{display:false},ticks:{callback:v=>F(v/1000)+'k'},
        title:{display:true,text:'N° de personas desocupadas',color:A1,font:{size:11,weight:'600'}}}},
   plugins:{...baseOpts.plugins,legend:{display:true,position:'top',align:'end',labels:{boxWidth:10,boxHeight:10,usePointStyle:true,pointStyle:'circle'}},
     tooltip:{...baseOpts.plugins.tooltip,callbacks:{
       title:it=>r.mesIni!=null? `Mes ${it[0].dataIndex + r.mesIni} · ${it[0].label}` : it[0].label,
       label:c=>c.datasetIndex===1? 'Tasa: '+P(c.raw)+'%' : 'Desocupados: '+F(c.raw)}}}}});
  const pri = s[0], ult = s[s.length-1];
  const dif = ult.td - pri.td, difn = ult.d - pri.d;
  const sg = v => (v>=0?'+':'−') + P(Math.abs(v));
  const sgn = v => (v>=0?'+':'−') + F(Math.abs(v));
  document.getElementById('serieCaption').innerHTML = (r.cap ? r.cap+' ' : '') +
    `${pri.lab} – ${ult.lab}: la tasa pasó de ${P(pri.td)}% a ${P(ult.td)}% ` +
    `(${sg(dif)} puntos porcentuales) y las personas desocupadas, de ${F(pri.d)} a ${F(ult.d)} (${sgn(difn)}).`;
}
/* El Gobierno de José Antonio Kast asumió el 11 de marzo de 2026. El punto de partida (mes 0) es
   DEF 2026 —el trimestre móvil diciembre-enero-febrero, último dato previo a la asunción— y los
   cinco meses acumulados llegan hasta MJJ 2026, publicado por el INE el 28 de agosto de 2026. */
const rangos=[
  {from:[2010,1], lab:'2010–2026'},
  {from:[2016,1], lab:'Últimos 10 años'},
  {from:[2020,1], lab:'Desde la pandemia'},
  {from:[2024,1], lab:'Últimos 3 años'},
  {from:[2026,2], lab:'Gobierno Kast', cls:'kast', mesIni:1,
   cap:'El Gobierno de José Antonio Kast asumió el 11 de marzo de 2026. Se muestran los cinco '+
       'trimestres móviles del período, desde EFM 2026 hasta MJJ 2026, dato publicado por el INE '+
       'el 28 de agosto de 2026.'}
];
document.getElementById('rangeTabs').innerHTML = rangos.map((r,i)=>
  `<div class="tab ${r.cls||''}${i?'':' on'}" data-i="${i}">${r.lab}</div>`).join('');
document.querySelectorAll('#rangeTabs .tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('#rangeTabs .tab').forEach(x=>x.classList.remove('on'));
  t.classList.add('on'); drawSerie(rangos[+t.dataset.i]);});
drawSerie(rangos[0]);

/* ---- sección 1 ---- */
const sx = D.Sexo.rows;
barH('c_sexo', sx.map(r=>r.cat), sx.map(r=>r.tasa), [A1,A2]);
const ed = D.Edad.rows;
barH('c_edad', ed.map(r=>r.cat), ed.map(r=>r.tasa), A1);
barH('c_edad_d', ed.map(r=>r.cat), ed.map(r=>r.pct), A2);
const rg = D.Region.rows.slice().sort((a,b)=>b.tasa-a.tasa);
barH('c_region', rg.map(r=>r.cat.replace(/^Región de(l)? /,'')), rg.map(r=>r.tasa), A1);
/* Provincias: tabla propia en dos columnas, ordenada por tasa. Las provincias con pocos casos
   se marcan y pueden ocultarse, porque su tasa es demasiado ruidosa para compararla. */
function pintarProv(min){
  const rows = D.Provincia.rows.filter(r=>r.casos>=min).sort((a,b)=>b.tasa-a.tasa);
  const max = Math.max(...rows.map(r=>r.tasa));
  const mitad = Math.ceil(rows.length/2);
  const col = rs => `<table><thead><tr><th>Provincia</th><th style="width:30%"></th>
      <th class="n">Tasa</th><th class="n">Casos</th></tr></thead><tbody>` +
    rs.map(r=>{const flojo = r.casos<50;
      const nom = r.cat.replace(/ \\(([^)]+)\\)$/, ' <span style="color:var(--muted)">· $1</span>');
      return `<tr${flojo?' style="opacity:.6"':''}><td>${nom}${flojo?' <span title="menos de 50 casos" style="color:var(--a4)">▵</span>':''}</td>
        <td><div class="barwrap"><div class="bar" style="width:${Math.max(2,r.tasa/max*100)}%;
          background:${flojo?'#b9c2cb':A1}"></div></div></td>
        <td class="n"><b>${P(r.tasa)}%</b></td>
        <td class="n" style="color:var(--muted)">${F(r.casos)}</td></tr>`}).join('') + `</tbody></table>`;
  document.getElementById('t_prov').innerHTML =
    `<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:0 26px">
       ${col(rows.slice(0,mitad))}${col(rows.slice(mitad))}</div>`;
}
document.querySelectorAll('#provTabs .tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('#provTabs .tab').forEach(x=>x.classList.remove('on'));
  t.classList.add('on'); pintarProv(+t.dataset.min);});
pintarProv(50);

tabla('t_educ','Educacion',{head:'% desoc.'});
tabla('t_cond','Condicion',{color:A2});
tabla('t_nac','Nacionalidad',{color:A3});
tabla('t_jef','Jefatura_hogar',{color:A5});
tabla('t_pi','Pueblo_indigena',{color:A5});

/* ---- sección 2: serie de larga duración ---- */
{
  const L = D.serie_dur;
  new Chart(document.getElementById('c_larga'),{type:'line',
    data:{labels:L.map(x=>x.lab),datasets:[
      {label:'Buscan hace 12 meses o más (% de las personas desocupadas)',
       data:L.map(x=>x.bus),borderColor:A3,backgroundColor:A3,borderWidth:2.4,pointRadius:0,
       pointHoverRadius:4,tension:.25},
      {label:'Cesantes hace 12 meses o más (% de las personas cesantes)',
       data:L.map(x=>x.ces==null?null:x.ces),borderColor:A2,backgroundColor:A2,borderWidth:2.4,
       pointRadius:0,pointHoverRadius:4,tension:.25,spanGaps:false}
    ]},
    options:{...baseOpts,interaction:{mode:'index',intersect:false},
     scales:{x:{grid:{display:false},border:{color:LINE},ticks:{maxTicksLimit:14,maxRotation:0}},
             y:{beginAtZero:true,grid:{color:LINE},border:{display:false},ticks:{callback:v=>v+'%'}}},
     plugins:{...baseOpts.plugins,
       legend:{display:true,position:'top',align:'start',
               labels:{boxWidth:10,boxHeight:10,usePointStyle:true,pointStyle:'circle',padding:14}},
       tooltip:{...baseOpts.plugins.tooltip,callbacks:{
         label:c=>(c.datasetIndex? 'Cesantes 12m+: ':'Buscando 12m+: ')+P(c.raw)+'%'}}}}});
  const u = L[L.length-1], pico = L.reduce((m,x)=>x.bus>m.bus?x:m,L[0]);
  document.getElementById('largaCaption').innerHTML =
    `En ${u.lab} el ${P(u.bus)}% de las personas desocupadas llevaba doce meses o más buscando trabajo, y el ` +
    `${P(u.ces)}% de las personas cesantes llevaba doce meses o más sin empleo. El máximo de la serie es ` +
    `${P(pico.bus)}% en ${pico.lab}, la resaca de la pandemia. La línea de cesantía parte en JJA 2020 porque ` +
    `la pregunta sobre el término del último empleo se incorpora al cuestionario en julio de ese año. ` +
    `Los porcentajes son algo más altos que los de la tabla de tramos que sigue: aquí se recuperan, por la vía ` +
    `del año declarado, los casos que no recuerdan el mes exacto, que son sobre todo episodios largos.`;
}

/* ---- sección 2 ---- */
const db = D.Duracion_busqueda.rows;
barH('c_dbus', db.map(r=>r.cat), db.map(r=>r.pct), A3);
const dc = D.Duracion_cesantia.rows;
barH('c_dces', dc.map(r=>r.cat), dc.map(r=>r.pct), A3);
tabla('t_met','Metodos_busqueda',{sort:true,color:A4,head:'% desoc.'});
tabla('t_jor','Jornada_buscada',{color:A4});

/* ---- sección 3 ---- */
tabla('t_mot','Motivo_termino',{sort:true,color:A2});
tabla('t_desp','Motivo_despido',{sort:true,color:A2});
tabla('t_ren','Motivo_renuncia',{sort:true,color:A2});

/* ---- sección 4: panel ---- */
const ps = D.Panel_situacion_2025.rows;
barH('c_psit', ps.map(r=>r.cat), ps.map(r=>r.pct), [A3,A2,A2,A4,A5]);
const pr = D.Panel_rama_2025.rows.filter(r=>r.n>0).sort((a,b)=>b.pct-a.pct);
barH('c_prama', pr.map(r=>r.cat.replace(/^[A-Z]{1,2}\\. /,'')), pr.map(r=>r.pct), A1);
tabla('t_pcon','Panel_contrato_2025',{color:A1});
tabla('t_pcat','Panel_categoria_2025',{sort:true,color:A1});
tabla('t_pfor','Panel_formalidad_2025',{color:A2});
tabla('t_psec','Panel_sector_inst_2025',{color:A3});

/* ---- enlace al repositorio: mismo bloque en la portada y en el pie ---- */
{
  const tpl = document.getElementById('ghlink').innerHTML;
  document.getElementById('gh_hero').innerHTML = tpl;
  document.getElementById('gh_foot').innerHTML = tpl;
}

/* ---- pestañas: cada título de nivel 1 es una pestaña ---------------------------
   Se activa al final, cuando todos los gráficos ya se dibujaron con las secciones
   visibles: si Chart.js crea un canvas dentro de un contenedor oculto lo mide en
   cero y no siempre se recupera al mostrarlo. */
(function(){
  const secs = [...document.querySelectorAll('.wrap > section')];
  const nombre = s => s.dataset.tab || s.querySelector('.shead h2').textContent;
  const num    = s => s.querySelector('.shead .num') ? s.querySelector('.shead .num').textContent : '';
  document.getElementById('navtabs').innerHTML = secs.map((s,i)=>
    `<button data-i="${i}">${num(s)?`<i>${num(s)}</i>`:''}${nombre(s)}</button>`).join('');
  const btns = [...document.querySelectorAll('#navtabs button')];
  // El número ya va en la pestaña; repetirlo dentro de la sección es ruido.
  secs.forEach(s=>{ const n=s.querySelector('.shead .num'); if(n) n.remove(); });

  function ir(i, saltar){
    secs.forEach((s,j)=>s.hidden = j!==i);
    btns.forEach((b,j)=>b.classList.toggle('on', j===i));
    try{ localStorage.setItem('tab', i); }catch(e){}
    if(saltar){
      const nav = document.getElementById('nav');
      const y = nav.getBoundingClientRect().top + window.scrollY;
      if(window.scrollY > y) window.scrollTo({top:y, behavior:'instant'});
    }
    // Chart.js mide el canvas al crearlo: hay que pedirle que recalcule al mostrarlo.
    secs[i].querySelectorAll('canvas').forEach(cv=>{
      const ch = Chart.getChart(cv); if(ch) ch.resize();
    });
  }
  btns.forEach((b,i)=> b.onclick = ()=> ir(i, true));
  let ini = 0;
  try{ const g = +localStorage.getItem('tab'); if(g>=0 && g<secs.length) ini = g; }catch(e){}
  ir(ini, false);
})();

/* ---- notas ---- */
document.getElementById('notas').innerHTML = `<h3>Nota metodológica</h3>
<p><b>Fuente.</b> Microdatos de la Encuesta Nacional de Empleo (ENE), Instituto Nacional de Estadísticas de Chile. El corte principal es el trimestre móvil mayo–junio–julio de 2026 (97.946 personas encuestadas). La serie histórica cubre todos los trimestres móviles disponibles desde 2010 hasta MJJ 2026 (${S.length} trimestres).</p>
<p><b>Definiciones.</b> Se considera desocupada a la persona sin empleo que buscó trabajo en las últimas cuatro semanas y está disponible para trabajar (marco conceptual de la OIT, 19ª CIET). Dentro de ese grupo se distingue a las personas cesantes, con experiencia laboral previa, de quienes buscan trabajo por primera vez. La tasa de desocupación es el cociente entre personas desocupadas y fuerza de trabajo. Todas las cifras de personas están expandidas con el factor trimestral <code>fact_cal</code>.</p>
<p><b>Panel.</b> La ENE es un panel rotativo: cada persona se entrevista tres meses seguidos, sale de la muestra y vuelve doce meses después. El enlace entre MJJ 2025 y MJJ 2026 se hace por identificador de persona contra el mismo mes calendario, validando igual sexo y una diferencia de edad de cero a dos años. Quedan 25.486 pares válidos, de los cuales 1.086 corresponden a personas desocupadas en 2026 (24,6% del total). <b>La submuestra panel no tiene diseño muestral ni factores de expansión propios</b>: las cifras expandidas son solo indicativas, la distribución porcentual es la lectura recomendada y no corresponde presentarlas como estimaciones oficiales del INE.</p>
<p><b>Larga duración.</b> La serie trimestral de episodios de doce meses o más se construye desde los microdatos de cada trimestre móvil, no desde el corte de MJJ 2026. La duración se calcula contra el mes y año que declara la persona: el inicio de la búsqueda (e6) para las personas desocupadas y el término del último empleo (e21, disponible desde el cuestionario de julio de 2020) para las cesantes. Una fracción variable de las respuestas entrega el año pero no el mes, así que la clasificación se resuelve por la brecha de años —mismo año, menos de doce meses; dos años o más, doce meses o más— y solo queda por imputar el caso de exactamente un año de diferencia sin mes declarado, al que se le asigna la proporción que corresponde suponiendo el mes uniforme dentro del año. Ese grupo pesa alrededor del 1% de la muestra, salvo entre diciembre de 2023 y diciembre de 2024, cuando una ola de no respuesta del mes en la pregunta e21 lo lleva hasta la mitad de las personas cesantes: el nivel de esa línea en ese tramo debe leerse con cautela. Como el procedimiento recupera casos que la tabla de tramos deja en "no declarado", y esos casos son sobre todo episodios largos, los porcentajes de la serie son algo más altos que los del tabulado transversal.</p>
<p><b>Precisión.</b> Las columnas de casos muestrales se muestran junto a cada porcentaje porque en las desagregaciones más finas el número de observaciones es pequeño: conviene ser cauto con cualquier categoría bajo 50 casos. Los métodos de búsqueda son de respuesta múltiple, por lo que sus porcentajes no suman cien.</p>`;
</script>
"""

html = HTML.replace("__DATA__", DATA_JS).replace("__REPO__", REPO)
open(OUT, "w", encoding="utf-8").write(html)
print(f"escrito {OUT} ({len(html):,} caracteres, {len(D['serie'])} trimestres)")
