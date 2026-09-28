# -*- coding: utf-8 -*-
"""Cuadro de ocupación de Transformación Digital, leído de Odoo por XML-RPC.

La tabla sigue la hoja de seguimiento: Contratadas, Acumulado, Mes, BackLog,
Planning, Done y Desface. Al marcar un proyecto se despliega el detalle.

Secrets en `.streamlit/secrets.toml` (local) o en Settings → Secrets (Cloud).
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from odoo_io import (
    OdooError,
    clasificar_movimiento,
    cuadro_entrega,
    cuadro_reporte,
    es_bolsa_de_horas,
    load_extra_projects,
    load_horas_equipo,
    load_movimientos,
    load_plantilla,
    load_registered,
    load_snapshot,
)

st.set_page_config(
    page_title="Ocupación de proyectos · Transformación Digital",
    page_icon="📋",
    layout="wide",
)

COLUMNAS_EXCEL = [
    "Proyecto", "Etapa Proyecto", "Fecha de Inicio", "Fecha de Fin", "Etiquetas",
    "Vendedor", "Gerente", "Pry Cierre", "Rapidez", "Asignadas", "Contratadas",
    "Acumulado", "Mes", "BackLog", "Planning", "Done", "Desface",
]
HORAS = ["Asignadas", "Contratadas", "Acumulado", "Mes", "BackLog", "Planning", "Done", "Desface"]


def csv_bytes(frame: pd.DataFrame) -> bytes:
    return frame.to_csv(index=False).encode("utf-8-sig")


def fecha_txt(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return pd.Timestamp(value).strftime("%d/%m/%Y")


def etiquetar(projects: pd.DataFrame) -> pd.DataFrame:
    out = projects.copy()
    if out.empty:
        out["proyecto_ui"] = pd.Series(dtype="str")
        return out
    conteo = out.groupby("proyecto")["project_id"].nunique()
    repetidos = set(conteo[conteo > 1].index)
    out["proyecto_ui"] = [
        f"{nombre} · {pid}" if nombre in repetidos else nombre
        for nombre, pid in zip(out["proyecto"], out["project_id"], strict=True)
    ]
    return out


def columna_numero(nombre: str, ayuda: str | None = None):
    return st.column_config.NumberColumn(nombre, format="%.2f", help=ayuda)


def columna_numero(nombre: str, ayuda: str | None = None):
    return st.column_config.NumberColumn(nombre, format="%.2f", help=ayuda)


def _detalle_proyecto(proyectos, acciones, horas_reg, fecha_reporte, elegido):
    pid = int(elegido["project_id"])
    ficha_df = proyectos[proyectos["project_id"] == pid]
    ficha = ficha_df.iloc[0] if not ficha_df.empty else None
    if ficha is not None:
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"**Etapa**  \n{ficha['etapa_proyecto']}")
        c2.markdown(f"**Vendedor**  \n{ficha['vendedor']}")
        c3.markdown(f"**Gerente**  \n{ficha['gerente']}")
        c4.markdown(f"**Rapidez**  \n{float(ficha['rapidez']):.1f}")
        d1, d2, d3 = st.columns(3)
        d1.markdown(f"**Inicio**  \n{fecha_txt(ficha['fecha_inicio'])}")
        d2.markdown(f"**Fin**  \n{fecha_txt(ficha['fecha_fin'])}")
        d3.markdown(f"**Cierre real**  \n{fecha_txt(ficha['fecha_real_cierre'])}")
        if ficha["etiquetas"]:
            st.markdown(f"**Etiquetas**  \n{ficha['etiquetas']}")
        extra_horas = []
        if float(ficha["horas_bloqueadas"] or 0) > 0:
            extra_horas.append(f"Bloqueadas: {float(ficha['horas_bloqueadas']):,.2f}")
        if float(ficha["horas_cerradas_en_curso"] or 0) > 0:
            extra_horas.append(f"Cerradas en curso: {float(ficha['horas_cerradas_en_curso']):,.2f}")
        if extra_horas:
            st.caption(
                "Estas horas están en las tareas del proyecto y no tienen columna en el cuadro: "
                + " · ".join(extra_horas)
            )

    st.markdown("**Recursos**")
    detalle_acc = acciones[acciones["project_id"] == pid] if not acciones.empty else acciones
    if detalle_acc.empty:
        st.caption("Sin horas de tareas asignadas a un responsable.")
    else:
        recursos = (
            detalle_acc.pivot_table(
                index="usuario", columns="accion", values="cantidad", aggfunc="sum", fill_value=0,
            )
            .reset_index()
            .rename(columns={
                "usuario": "Recurso",
                "H. Pendientes": "BackLog",
                "H. Planeadas": "Planning",
                "H. Ejecutadas": "Done",
                "H. Bloqueadas": "Bloqueadas",
            })
        )
        st.dataframe(
            recursos,
            hide_index=True,
            use_container_width=True,
            column_config={
                c: st.column_config.NumberColumn(format="%.2f")
                for c in recursos.columns if c != "Recurso"
            },
        )
        st.caption(
            "Si la tarea tiene varios responsables, cada uno recibe las horas completas. "
            "El BackLog, Planning y Done del cuadro cuentan la tarea una sola vez."
        )

    st.markdown("**Horas registradas**")
    detalle_horas = horas_reg[horas_reg["project_id"] == pid] if not horas_reg.empty else horas_reg
    if detalle_horas.empty:
        st.caption("Sin horas registradas en el período.")
    else:
        detalle = detalle_horas.copy()
        es_mes = (detalle["anio"] == fecha_reporte.year) & (detalle["mes"] == fecha_reporte.month)
        detalle["Parte"] = es_mes.map({True: "Mes", False: "Acumulado"})
        detalle = detalle.rename(columns={
            "responsable": "Responsable",
            "periodo": "Período",
            "horas": "Horas",
        })[["Responsable", "Período", "Parte", "Horas"]]
        detalle = detalle.sort_values(["Parte", "Período", "Responsable"], kind="stable")
        st.dataframe(
            detalle,
            hide_index=True,
            use_container_width=True,
            column_config={"Horas": st.column_config.NumberColumn(format="%.2f")},
        )


def pintar_cuadro(proyectos, acciones, horas_reg, fecha_reporte, clave: str, archivo: str, vacio: str):
    """Tabla del cuadro y, al seleccionar una fila, el detalle del proyecto."""
    cuadro = cuadro_reporte(proyectos, horas_reg, fecha_reporte)
    if cuadro.empty:
        st.info(vacio)
        return

    base = proyectos.drop_duplicates("project_id").set_index("project_id")
    ids = cuadro["project_id"]

    def traer(columna):
        return ids.map(base[columna]) if columna in base.columns else None

    tabla = pd.DataFrame({
        "Proyecto": cuadro["proyecto_ui"].to_numpy(),
        "Etapa Proyecto": cuadro["etapa_proyecto"].to_numpy(),
        "Fecha de Inicio": traer("fecha_inicio").to_numpy(),
        "Fecha de Fin": traer("fecha_fin").to_numpy(),
        "Etiquetas": traer("etiquetas").to_numpy(),
        "Vendedor": traer("vendedor").to_numpy(),
        "Gerente": traer("gerente").to_numpy(),
        "Pry Cierre": traer("fecha_real_cierre").to_numpy(),
        "Rapidez": traer("rapidez").to_numpy(),
        "Asignadas": traer("horas_tareas").to_numpy(),
        "Contratadas": cuadro["contratadas"].to_numpy(),
        "Acumulado": cuadro["acumulado"].to_numpy(),
        "Mes": cuadro["mes"].to_numpy(),
        "BackLog": cuadro["backlog"].to_numpy(),
        "Planning": cuadro["planning"].to_numpy(),
        "Done": cuadro["done"].to_numpy(),
        "Desface": cuadro["defase"].to_numpy(),
    })
    evento = st.dataframe(
        tabla[COLUMNAS_EXCEL],
        hide_index=True,
        use_container_width=True,
        height=420,
        on_select="rerun",
        selection_mode="single-row",
        key=clave,
        column_config={
            "Proyecto": st.column_config.TextColumn("Proyecto", width="large"),
            "Etapa Proyecto": st.column_config.TextColumn("Etapa Proyecto", width="medium"),
            "Fecha de Inicio": st.column_config.DateColumn("Fecha de Inicio", format="DD/MM/YYYY"),
            "Fecha de Fin": st.column_config.DateColumn("Fecha de Fin", format="DD/MM/YYYY"),
            "Pry Cierre": st.column_config.DateColumn("Pry Cierre", format="DD/MM/YYYY"),
            "Rapidez": st.column_config.NumberColumn("Rapidez", format="%.1f"),
            "Asignadas": columna_numero("Asignadas", "Suma de horas de las tareas raíz"),
            "Contratadas": columna_numero("Contratadas", "Horas vendidas del proyecto"),
            "Acumulado": columna_numero("Acumulado", "Horas registradas antes del mes del reporte"),
            "Mes": columna_numero("Mes", "Horas registradas en el mes de la fecha del reporte"),
            "BackLog": columna_numero("BackLog", "Tareas raíz en etapa Inicio"),
            "Planning": columna_numero("Planning", "Tareas raíz en Planeado o En ejecución, sin Hecho ni Cancelado"),
            "Done": columna_numero("Done", "Tareas raíz en etapa Finalizado"),
            "Desface": columna_numero("Desface", "Contratadas − Acumulado − Mes"),
        },
    )

    filas = evento.selection.rows if evento is not None and evento.selection is not None else []
    if not filas:
        st.caption("Marca la casilla de un proyecto para desplegar sus recursos y sus horas registradas.")
    else:
        elegido = cuadro.iloc[filas[0]]
        with st.expander(f"Detalle · {elegido['proyecto_ui']}", expanded=True):
            _detalle_proyecto(proyectos, acciones, horas_reg, fecha_reporte, elegido)

    totales = {nombre: pd.to_numeric(tabla[nombre], errors="coerce").sum() for nombre in HORAS}
    st.caption(
        "Totales · "
        + " · ".join(f"{nombre}: {valor:,.2f}" for nombre, valor in totales.items())
    )
    st.download_button(
        "Descargar cuadro",
        csv_bytes(tabla),
        file_name=archivo,
        mime="text/csv",
        key=f"csv_{clave}",
    )


# ─────────────────────────────────────────────
# Período: desde el 1 de enero del año inicial hasta el día del reporte
# ─────────────────────────────────────────────
st.sidebar.title("Ocupación TD")
st.sidebar.caption("Cuadro de proyectos desde Odoo")

hoy = date.today()


def _factor_inicial() -> float:
    try:
        raw = st.secrets["odoo"].get("factor_entrega")
        if raw in (None, ""):
            return 1.0
        valor = float(raw)
        return valor / 100 if valor > 2 else valor
    except Exception:
        return 1.0


def _factor(valor: float) -> float:
    return valor / 100 if valor > 2 else valor


def pintar_mes(proyectos, acciones, por_empleado, fecha_reporte, factor: float, filtro_etapa, buscar: str):
    """Horas que se debían entregar en el mes y movimiento de proyectos."""
    inicio_mes = date(fecha_reporte.year, fecha_reporte.month, 1)
    st.caption(
        f"Del {inicio_mes:%d/%m/%Y} al {fecha_reporte:%d/%m/%Y}. "
        "Debería = (horario − festivos − ausencias) × factor. "
        "Entregó = todas las horas que la persona registró en Odoo ese mes, en cualquier proyecto."
    )

    horas_mes = por_empleado
    if horas_mes is None or horas_mes.empty or proyectos.empty:
        horas_mes = por_empleado.iloc[0:0] if por_empleado is not None else por_empleado
    elif not proyectos.empty and "project_id" in por_empleado.columns:
        horas_mes = por_empleado[por_empleado["project_id"].isin(set(proyectos["project_id"]))]
        horas_mes = horas_mes[(horas_mes["anio"] == fecha_reporte.year) & (horas_mes["mes"] == fecha_reporte.month)]

    ids = tuple(sorted({
        int(valor) for valor in (horas_mes["employee_id"].dropna().tolist() if horas_mes is not None and not horas_mes.empty else [])
    }))
    nombres = tuple(sorted({
        str(nombre) for nombre in (
            acciones["usuario"].dropna().tolist() if acciones is not None and not acciones.empty else []
        ) if nombre and nombre != "-"
    }))
    try:
        plantilla = load_plantilla(ids, nombres, inicio_mes.isoformat(), fecha_reporte.isoformat())
    except OdooError as exc:
        st.error(str(exc))
        plantilla = None
    for aviso in (plantilla or {}).get("warnings") or []:
        st.warning(aviso)

    horas_entregadas = horas_mes if horas_mes is not None else pd.DataFrame()
    if plantilla is not None and not plantilla["empleados"].empty:
        equipo = tuple(sorted({
            int(valor) for valor in plantilla["empleados"]["employee_id"].dropna().tolist()
        }))
        try:
            horas_entregadas = load_horas_equipo(equipo, inicio_mes.isoformat(), fecha_reporte.isoformat())
        except OdooError as exc:
            st.warning(str(exc))

    st.markdown("**Horas del mes**")
    if plantilla is None:
        entrega = pd.DataFrame()
    else:
        entrega = cuadro_entrega(
            plantilla["empleados"],
            plantilla["asistencia"],
            plantilla["ausencias"],
            horas_entregadas,
            inicio_mes,
            fecha_reporte,
            plantilla["tz"],
            factor,
            plantilla["calendario_compania"],
        )
    if entrega.empty:
        st.info("No hay empleados con horario ni horas registradas en este mes.")
    else:
        tabla_horas = entrega.rename(columns={
            "empleado": "Empleado",
            "horario": "Horario",
            "festivos": "Festivos",
            "ausencias": "Ausencias",
            "laborables": "Laborables",
            "factor": "Factor",
            "deberia": "Debería",
            "entrego": "Entregó",
            "diferencia": "Diferencia",
        })
        st.dataframe(
            tabla_horas,
            hide_index=True,
            use_container_width=True,
            column_config={
                columna: st.column_config.NumberColumn(format="%.2f")
                for columna in tabla_horas.columns if columna != "Empleado"
            },
        )
        st.caption(
            "Totales · "
            + " · ".join(
                f"{nombre}: {pd.to_numeric(tabla_horas[nombre], errors='coerce').sum():,.2f}"
                for nombre in ("Horario", "Festivos", "Ausencias", "Debería", "Entregó", "Diferencia")
            )
        )

    st.markdown("**Proyectos del mes**")
    try:
        extra, avisos = load_movimientos(inicio_mes.isoformat(), fecha_reporte.isoformat())
    except OdooError as exc:
        extra, avisos = pd.DataFrame(), [str(exc)]
    for aviso in avisos:
        st.warning(aviso)
    universo = proyectos
    if extra is not None and not extra.empty:
        extra = etiquetar(extra)
        if filtro_etapa:
            extra = extra[extra["etapa_proyecto"].isin(filtro_etapa)]
        if buscar.strip():
            texto = buscar.strip().casefold()
            extra = extra[extra["proyecto_ui"].astype(str).str.casefold().str.contains(texto, regex=False)]
        if not proyectos.empty:
            extra = extra[~extra["project_id"].isin(set(proyectos["project_id"]))]
        universo = pd.concat([proyectos, extra], ignore_index=True)
    movimiento = clasificar_movimiento(universo, inicio_mes, fecha_reporte)
    etiquetas = {
        "abiertos": "Abiertos en el mes",
        "cerrados": "Cerrados en el mes",
        "en_proceso": "Siguen en proceso",
    }
    for clave, titulo in etiquetas.items():
        bloque = movimiento[clave]
        st.markdown(f"**{titulo}** · {len(bloque)}")
        if bloque.empty:
            st.caption("Ninguno.")
            continue
        vista = bloque.rename(columns={
            "proyecto_ui": "Proyecto",
            "etapa_proyecto": "Etapa Proyecto",
            "fecha_inicio": "Inicio",
            "fecha_real_cierre": "Cierre real",
            "horas_vendidas": "Contratadas",
        })
        columnas = [c for c in ("Proyecto", "Etapa Proyecto", "Inicio", "Cierre real", "Contratadas") if c in vista.columns]
        st.dataframe(
            vista[columnas],
            hide_index=True,
            use_container_width=True,
            column_config={
                "Inicio": st.column_config.DateColumn("Inicio", format="DD/MM/YYYY"),
                "Cierre real": st.column_config.DateColumn("Cierre real", format="DD/MM/YYYY"),
                "Contratadas": st.column_config.NumberColumn(format="%.2f"),
            },
        )
anios = list(range(hoy.year, 2019, -1))
anio = st.sidebar.selectbox(
    "Año inicial",
    anios,
    index=anios.index(2024),
    help="Las horas de Acumulado se cargan desde el 1 de enero de este año. En el Excel el histórico arranca en 2024.",
)
fecha_reporte = st.sidebar.date_input(
    "Fecha del reporte",
    value=hoy,
    format="DD/MM/YYYY",
    help="Por defecto es el día en que se abre el reporte. Acumulado llega hasta el mes anterior. Mes es este mes, hasta este día.",
)
if not isinstance(fecha_reporte, date):
    fecha_reporte = hoy
inicio = date(int(anio), 1, 1)
if fecha_reporte < inicio:
    st.sidebar.warning("La fecha del reporte no puede ser anterior al 1 de enero del año inicial.")
    fecha_reporte = inicio

factor_capturado = st.sidebar.number_input(
    "Factor de entrega",
    min_value=0.0,
    max_value=150.0,
    value=_factor_inicial(),
    step=0.05,
    help="En el Excel cada persona tiene su factor (por ejemplo 0,85 o 0,70). Si Odoo no tiene ese campo, este valor se aplica a todo el equipo. Puede escribir 0,8 o 80.",
)
factor = _factor(float(factor_capturado))

if st.sidebar.button("Refrescar datos", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

st.sidebar.caption(f"Caché de 10 min · {datetime.now():%H:%M}")

try:
    proyectos_base, acciones_base, info = load_snapshot()
    registradas, por_empleado, info_horas = load_registered(inicio.isoformat(), fecha_reporte.isoformat())
except OdooError as exc:
    st.error(str(exc))
    st.stop()

nombres_horas = tuple(
    (int(pid), str(nombre))
    for pid, nombre in (
        registradas[["project_id", "proyecto"]].drop_duplicates().itertuples(index=False, name=None)
        if not registradas.empty else []
    )
)
conocidos = set(proyectos_base["project_id"]) if not proyectos_base.empty else set()
faltantes = tuple(sorted({pid for pid, _nombre in nombres_horas if pid not in conocidos}))
try:
    extras = load_extra_projects(faltantes, nombres_horas)
except OdooError as exc:
    extras = pd.DataFrame(columns=proyectos_base.columns)
    info.setdefault("warnings", []).append(str(exc))

if not extras.empty:
    proyectos_base = pd.concat([proyectos_base, extras], ignore_index=True)
proyectos_base = etiquetar(proyectos_base)
if not proyectos_base.empty:
    etiqueta = dict(zip(proyectos_base["project_id"], proyectos_base["proyecto_ui"], strict=True))
    if not acciones_base.empty:
        acciones_base = acciones_base.copy()
        acciones_base["proyecto"] = acciones_base["project_id"].map(etiqueta).fillna(acciones_base["proyecto"])
    if not registradas.empty:
        registradas = registradas.copy()
        registradas["proyecto"] = registradas["project_id"].map(etiqueta).fillna(registradas["proyecto"])

for aviso in list(info.get("warnings") or []) + list(info_horas.get("warnings") or []):
    st.sidebar.warning(aviso)

etapas = sorted(p for p in proyectos_base["etapa_proyecto"].dropna().unique()) if not proyectos_base.empty else []
filtro_etapa = st.sidebar.multiselect("Etapa del proyecto", etapas)
buscar = st.sidebar.text_input("Buscar proyecto", placeholder="Nombre o código")

with st.sidebar.expander("Criterio de las horas"):
    st.markdown(
        f"Compañía **{info.get('company_id', info_horas.get('company_id'))}** · "
        f"línea `{info.get('service_line', '')}`."
    )
    st.markdown(
        f"Registradas del **{inicio:%d/%m/%Y}** al **{fecha_reporte:%d/%m/%Y}**. "
        f"Se excluyen {info_horas.get('exclude_project_ids', [])} "
        f"y se suman también {info_horas.get('include_project_ids', [])}."
    )
    etapas_xml = pd.DataFrame(info.get("etapas") or [])
    if not etapas_xml.empty:
        st.dataframe(etapas_xml, hide_index=True, use_container_width=True)

proyectos = proyectos_base
acciones = acciones_base
horas_reg = registradas
if filtro_etapa:
    proyectos = proyectos[proyectos["etapa_proyecto"].isin(filtro_etapa)]
if buscar.strip():
    texto = buscar.strip().casefold()
    proyectos = proyectos[proyectos["proyecto_ui"].astype(str).str.casefold().str.contains(texto, regex=False)]
ids = set(proyectos["project_id"])
acciones = acciones[acciones["project_id"].isin(ids)] if not acciones.empty else acciones
horas_reg = horas_reg[horas_reg["project_id"].isin(ids)] if not horas_reg.empty else horas_reg

if proyectos.empty:
    bolsas = proyectos
    resto = proyectos
else:
    mask_bolsa = proyectos["etiquetas"].map(es_bolsa_de_horas)
    bolsas = proyectos[mask_bolsa]
    resto = proyectos[~mask_bolsa]

st.title("Ocupación de proyectos")
st.caption(
    f"Desde el 1 de enero de {int(anio)} hasta el {fecha_reporte:%d/%m/%Y}, "
    "día del reporte. Acumulado es lo registrado hasta el mes anterior. "
    "Mes es el mes de la fecha del reporte, hasta ese día. "
    "Desface = Contratadas − Acumulado − Mes. "
    "Las bolsas de horas están en su propia pestaña, por la etiqueta Bolsa de Horas."
)

tab_proyectos, tab_bolsas, tab_mes = st.tabs(["Proyectos", "Bolsas de horas", "Mes"])
with tab_proyectos:
    st.caption("Proyectos sin la etiqueta Bolsa de Horas.")
    pintar_cuadro(
        resto, acciones, horas_reg, fecha_reporte,
        clave="cuadro_proyectos",
        archivo=f"ocupacion_td_{fecha_reporte:%Y%m%d}.csv",
        vacio="No hay proyectos en esta pestaña con los filtros actuales.",
    )
with tab_bolsas:
    st.caption("Proyectos con la etiqueta Bolsa de Horas.")
    pintar_cuadro(
        bolsas, acciones, horas_reg, fecha_reporte,
        clave="cuadro_bolsas",
        archivo=f"bolsas_horas_td_{fecha_reporte:%Y%m%d}.csv",
        vacio="No hay bolsas de horas con los filtros actuales.",
    )
with tab_mes:
    pintar_mes(proyectos, acciones, por_empleado, fecha_reporte, factor, filtro_etapa, buscar)
