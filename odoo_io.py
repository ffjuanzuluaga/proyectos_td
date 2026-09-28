# -*- coding: utf-8 -*-
"""Lectura de Odoo 19 por XML-RPC para el tablero de ocupación de proyectos TD.

Equivalencia con las consultas SQL que armaban el Excel:

1. Proyectos activos con ``service_line = digital_transformation`` de la compañía.
   Horas del proyecto = suma de ``allocated_hours`` de tareas raíz cuyas etapas
   son los xmlid de ``l10n_co_firefly_project`` (inicio, planeado, ejecución,
   bloqueado, finalizado).
2. Misma población, partida por responsable:
   - H. Pendientes: etapa Inicio. Sin responsable, se usa el gerente.
   - H. Planeadas: Planeado o En ejecución, y el estado no es Hecho ni Cancelado.
     Sin responsable, se usa el gerente.
   - H. Ejecutadas: etapa Finalizado. Sin responsable queda «-».
   Si la tarea tiene varios responsables, las horas se atribuyen completas a
   cada uno (el JOIN de ``project_task_user_rel`` no parte las horas).
   El total del proyecto no se multiplica: se cuenta cada tarea una sola vez.
3. Horas registradas: ``account.analytic.line.unit_amount`` entre dos fechas.
   Entran los proyectos DT excepto el id configurado (GS Gestión Soporte, 567)
   y además el proyecto indicado aunque no tenga línea de servicio (Customer
   Care, 119). El proyecto es el de la tarea si la línea tiene tarea; si no,
   el de la línea analítica.

Credenciales y parámetros van en ``st.secrets["odoo"]``.
"""

from __future__ import annotations

import threading
import unicodedata
import xmlrpc.client
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

MODULE_ETAPAS = "l10n_co_firefly_project"
XML_ACCION = {
    "project_task_type_start": "H. Pendientes",
    "project_task_type_planning": "H. Planeadas",
    "project_task_type_execution": "H. Planeadas",
    "project_task_type_blocked": "H. Bloqueadas",
    "project_task_type_end": "H. Ejecutadas",
}
ACCION_ORDEN = {
    "H. Pendientes": 1,
    "H. Planeadas": 2,
    "H. Ejecutadas": 3,
    "H. Bloqueadas": 4,
}
ACCION_CODIGO = {
    "H. Pendientes": -3,
    "H. Planeadas": -4,
    "H. Ejecutadas": -5,
    "H. Bloqueadas": -6,
}
# En la consulta SQL, pendientes y planeadas caen al gerente del proyecto
# cuando la tarea no tiene responsables. Ejecutadas no.
ACCION_CON_GERENTE = {"H. Pendientes", "H. Planeadas", "H. Bloqueadas"}
ESTADOS_FUERA_DE_PLANEADAS = {"1_done", "1_canceled"}

PROJECT_FIELDS = [
    "name",
    "stage_id",
    "tag_ids",
    "date_start",
    "date",
    "allocated_hours",
    "user_id",
    "x_studio_vendedor_1",
    "x_studio_rapidez",
    "x_studio_fecha_cierre",
    "active",
]
TASK_FIELDS = ["project_id", "stage_id", "allocated_hours", "user_ids", "user_id", "state", "active", "parent_id"]

PROJECT_COLS = [
    "project_id",
    "proyecto",
    "etapa_proyecto",
    "etiquetas",
    "fecha_inicio",
    "fecha_fin",
    "horas_vendidas",
    "vendedor",
    "gerente",
    "rapidez",
    "fecha_real_cierre",
    "horas_tareas",
    "horas_pendientes",
    "horas_planeadas",
    "horas_ejecutadas",
    "horas_bloqueadas",
    "horas_cerradas_en_curso",
    "activo",
]
ACTION_COLS = [
    "project_id",
    "proyecto",
    "etapa",
    "accion",
    "codigo",
    "usuario",
    "cantidad",
    "orden",
]
HOURS_COLS = ["project_id", "proyecto", "accion", "responsable", "anio", "mes", "periodo", "horas"]


class OdooError(RuntimeError):
    """Fallo de configuración, autenticación o consulta XML-RPC."""


def _round2(value) -> float:
    """ROUND(x, 2) de PostgreSQL: half away from zero, no el redondeo bancario de Python."""
    try:
        number = Decimal(str(float(value)))
    except (TypeError, ValueError):
        return 0.0
    return float(number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _num(value) -> float:
    if value in (None, False):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _m2o_id(value):
    if isinstance(value, (list, tuple)) and value:
        try:
            return int(value[0])
        except (TypeError, ValueError):
            return None
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    return None


def _m2o_name(value, default: str = "-") -> str:
    if isinstance(value, (list, tuple)) and len(value) >= 2 and value[1]:
        return str(value[1])
    return default


def _id_list(value) -> list[int]:
    if not value or value is False:
        return []
    out: list[int] = []
    for item in value:
        try:
            out.append(int(item))
        except (TypeError, ValueError):
            continue
    return out


def _text(value, default: str = "-") -> str:
    if value in (None, False):
        return default
    text = str(value).strip()
    return text or default


def _secret_ids(section, key: str, default: list[int]) -> list[int]:
    try:
        if key not in section:
            return list(default)
        val = section[key]
    except Exception:
        return list(default)
    if val in (None, False, ""):
        return list(default)
    if isinstance(val, bool):
        return list(default)
    if isinstance(val, (int, float)):
        return [int(val)]
    return [int(x) for x in val]


def accion_de_tarea(stage_id, state, stage_accion: dict[int, str]) -> str | None:
    """Acción de la consulta 2. None si la etapa no aplica o si es planeada ya cerrada."""
    accion = stage_accion.get(int(stage_id)) if stage_id not in (None, False) else None
    if not accion:
        return None
    if accion == "H. Planeadas" and state in ESTADOS_FUERA_DE_PLANEADAS:
        return None
    return accion


def build_projects(projects: list[dict], tasks: list[dict], stage_accion: dict[int, str]) -> pd.DataFrame:
    """Una fila por proyecto. Las horas de tarea se suman una vez, sin partir por usuario."""
    if not projects:
        return pd.DataFrame(columns=PROJECT_COLS)

    rows = []
    for task in tasks:
        stage_id = _m2o_id(task.get("stage_id"))
        if stage_id not in stage_accion:
            continue
        project_id = _m2o_id(task.get("project_id"))
        if not project_id:
            continue
        base = stage_accion[stage_id]
        state = task.get("state") or None
        if state is False:
            state = None
        rows.append({
            "project_id": project_id,
            "horas": _num(task.get("allocated_hours")),
            "accion": accion_de_tarea(stage_id, state, stage_accion),
            "accion_base": base,
        })
    horas = pd.DataFrame(rows)

    out_rows = []
    for project in projects:
        pid = int(project["id"])
        part = horas[horas["project_id"] == pid] if not horas.empty else horas
        total = float(part["horas"].sum()) if not part.empty else 0.0
        by_accion = {}
        if not part.empty:
            counted = part[part["accion"].notna()]
            if not counted.empty:
                by_accion = counted.groupby("accion")["horas"].sum().to_dict()
        buckets = {name: float(by_accion.get(name, 0.0)) for name in ACCION_ORDEN}
        cerradas = total - sum(buckets.values())
        if abs(cerradas) < 1e-6:
            cerradas = 0.0
        out_rows.append({
            "project_id": pid,
            "proyecto": _text(project.get("name"), "Sin nombre"),
            "etapa_proyecto": _m2o_name(project.get("stage_id")),
            "etiquetas": project.get("_etiquetas") or "",
            "fecha_inicio": project.get("date_start") or None,
            "fecha_fin": project.get("date") or None,
            "horas_vendidas": _num(project.get("allocated_hours")),
            "vendedor": _m2o_name(project.get("x_studio_vendedor_1")),
            "gerente": _m2o_name(project.get("user_id")),
            "rapidez": _num(project.get("x_studio_rapidez")),
            "fecha_real_cierre": project.get("x_studio_fecha_cierre") or None,
            "horas_tareas": total,
            "horas_pendientes": buckets["H. Pendientes"],
            "horas_planeadas": buckets["H. Planeadas"],
            "horas_ejecutadas": buckets["H. Ejecutadas"],
            "horas_bloqueadas": buckets["H. Bloqueadas"],
            "horas_cerradas_en_curso": cerradas,
            "activo": project.get("active") is not False,
        })
    df = pd.DataFrame(out_rows, columns=PROJECT_COLS)
    for col in ("fecha_inicio", "fecha_fin", "fecha_real_cierre"):
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df.sort_values("proyecto", kind="stable").reset_index(drop=True)


def build_actions(
    projects: pd.DataFrame,
    tasks: list[dict],
    stage_accion: dict[int, str],
    user_names: dict[int, str],
) -> pd.DataFrame:
    """Horas por proyecto, acción y responsable. Varios responsables reciben las horas completas."""
    if projects.empty or not tasks:
        return pd.DataFrame(columns=ACTION_COLS)
    by_id = projects.set_index("project_id")
    rows = []
    for task in tasks:
        project_id = _m2o_id(task.get("project_id"))
        stage_id = _m2o_id(task.get("stage_id"))
        if project_id not in by_id.index or stage_id not in stage_accion:
            continue
        state = task.get("state") or None
        if state is False:
            state = None
        accion = accion_de_tarea(stage_id, state, stage_accion)
        if not accion:
            continue
        project = by_id.loc[project_id]
        if isinstance(project, pd.DataFrame):
            project = project.iloc[0]
        # La consulta SQL solo usa project_task_user_rel (user_ids). user_id
        # es el respaldo de instalaciones sin ese many2many.
        if "user_ids" in task:
            user_ids = list(dict.fromkeys(_id_list(task.get("user_ids"))))
        else:
            solo = _m2o_id(task.get("user_id"))
            user_ids = [solo] if solo else []
        if user_ids:
            usuarios = [user_names.get(uid, "-") for uid in user_ids]
        elif accion in ACCION_CON_GERENTE:
            usuarios = [_text(project["gerente"])]
        else:
            usuarios = ["-"]
        horas = _num(task.get("allocated_hours"))
        for usuario in usuarios:
            rows.append({
                "project_id": int(project_id),
                "proyecto": project["proyecto"],
                "etapa": project["etapa_proyecto"],
                "accion": accion,
                "codigo": ACCION_CODIGO[accion],
                "usuario": usuario,
                "cantidad": horas,
                "orden": ACCION_ORDEN[accion],
            })
    if not rows:
        return pd.DataFrame(columns=ACTION_COLS)
    df = pd.DataFrame(rows)
    group_cols = ["project_id", "proyecto", "etapa", "accion", "codigo", "usuario", "orden"]
    df = df.groupby(group_cols, as_index=False)["cantidad"].sum()
    df["cantidad"] = df["cantidad"].astype(float)
    return df[ACTION_COLS].sort_values(["proyecto", "orden", "usuario"], kind="stable").reset_index(drop=True)


def build_registered(lines: list[dict]) -> pd.DataFrame:
    """Suma horas registradas por proyecto, responsable y mes. El año se conserva."""
    if not lines:
        return pd.DataFrame(columns=HOURS_COLS)
    frame = pd.DataFrame(lines)
    frame["date"] = pd.to_datetime(frame["date"], errors="coerce")
    frame = frame[frame["date"].notna() & frame["project_id"].notna()]
    if frame.empty:
        return pd.DataFrame(columns=HOURS_COLS)
    frame["anio"] = frame["date"].dt.year.astype(int)
    frame["mes"] = frame["date"].dt.month.astype(int)
    frame["horas"] = frame["horas"].map(_num)
    grouped = (
        frame.groupby(["project_id", "proyecto", "responsable", "anio", "mes"], as_index=False)["horas"]
        .sum()
    )
    grouped["horas"] = grouped["horas"].map(_round2)
    grouped["accion"] = "H. Registradas"
    grouped["periodo"] = grouped.apply(lambda r: f"{int(r['anio'])}-{int(r['mes']):02d}", axis=1)
    grouped = grouped[HOURS_COLS]
    return grouped.sort_values(["proyecto", "responsable", "periodo"], kind="stable").reset_index(drop=True)


EMPLEADO_COLS = ["employee_id", "project_id", "responsable", "anio", "mes", "horas"]


def build_horas_empleado(lines: list[dict]) -> pd.DataFrame:
    """Horas registradas por empleado y mes. employee_id puede quedar vacío."""
    if not lines:
        return pd.DataFrame(columns=EMPLEADO_COLS)
    frame = pd.DataFrame(lines)
    frame["date"] = pd.to_datetime(frame["date"], errors="coerce")
    frame = frame[frame["date"].notna()]
    if frame.empty:
        return pd.DataFrame(columns=EMPLEADO_COLS)
    frame["anio"] = frame["date"].dt.year.astype(int)
    frame["mes"] = frame["date"].dt.month.astype(int)
    frame["horas"] = frame["horas"].map(_num)
    frame["responsable"] = frame["responsable"].fillna("-")
    frame["project_id"] = pd.to_numeric(frame.get("project_id"), errors="coerce").astype("Int64")
    frame["employee_id"] = pd.to_numeric(frame.get("employee_id"), errors="coerce").astype("Int64")
    grouped = frame.groupby(EMPLEADO_COLS[:-1], as_index=False, dropna=False)["horas"].sum()
    grouped["horas"] = grouped["horas"].map(_round2)
    return grouped[EMPLEADO_COLS]


def es_bolsa_de_horas(etiquetas: str) -> bool:
    """True si alguna etiqueta del proyecto es Bolsa de Horas."""
    for parte in str(etiquetas or "").split(","):
        plano = unicodedata.normalize("NFD", parte)
        plano = "".join(c for c in plano if unicodedata.category(c) != "Mn")
        plano = "".join(ch for ch in plano.casefold() if ch.isalnum())
        if plano == "bolsadehoras" or plano.startswith("bolsadehoras"):
            return True
    return False


def es_etapa_cerrada(etapa: str) -> bool:
    plano = unicodedata.normalize("NFD", str(etapa or ""))
    plano = "".join(c for c in plano if unicodedata.category(c) != "Mn")
    plano = "".join(ch for ch in plano.casefold() if ch.isalnum())
    return plano.startswith("cerrad") or plano.startswith("cancel")


def _en_rango(serie: pd.Series, inicio: date, fin: date) -> pd.Series:
    fechas = pd.to_datetime(serie, errors="coerce")
    return fechas.notna() & (fechas >= pd.Timestamp(inicio)) & (fechas <= pd.Timestamp(fin))


def clasificar_movimiento(projects: pd.DataFrame, inicio: date, fin: date) -> dict[str, pd.DataFrame]:
    """Abiertos en el mes, cerrados en el mes y los que ya venían y siguen abiertos.

    Un proyecto abierto en el mes no entra en «en proceso». Si se abrió y se
    cerró en el mismo mes, aparece en las dos primeras listas.
    """
    if projects is None or projects.empty:
        vacio = projects.iloc[0:0].copy() if projects is not None else pd.DataFrame()
        return {"abiertos": vacio, "cerrados": vacio, "en_proceso": vacio}
    frame = projects.copy()
    inicio_ts = pd.to_datetime(frame["fecha_inicio"], errors="coerce")
    cierre = pd.to_datetime(frame["fecha_real_cierre"], errors="coerce")
    abiertos = _en_rango(frame["fecha_inicio"], inicio, fin)
    cerrados = _en_rango(frame["fecha_real_cierre"], inicio, fin)
    etapa_cerrada = frame["etapa_proyecto"].map(es_etapa_cerrada)
    ya_cerro = (cierre.notna() & (cierre <= pd.Timestamp(fin))) | (cierre.isna() & etapa_cerrada)
    venia_de_antes = inicio_ts.notna() & (inicio_ts < pd.Timestamp(inicio))
    sin_inicio = inicio_ts.isna()
    en_proceso = (venia_de_antes | sin_inicio) & ~ya_cerro & ~abiertos
    return {
        "abiertos": frame.loc[abiertos].sort_values("proyecto", kind="stable"),
        "cerrados": frame.loc[cerrados].sort_values("proyecto", kind="stable"),
        "en_proceso": frame.loc[en_proceso].sort_values("proyecto", kind="stable"),
    }


MESES_CORTO = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")


def periodos_recientes(fin: date, cantidad: int = 6) -> list[dict]:
    """Los últimos N meses, cortados el día del reporte si el mes no ha cerrado."""
    periodos = []
    anio, mes = fin.year, fin.month
    for _ in range(cantidad):
        inicio = date(anio, mes, 1)
        if mes == 12:
            ultimo = date(anio, 12, 31)
        else:
            ultimo = date(anio, mes + 1, 1) - timedelta(days=1)
        corte = fin if (anio, mes) == (fin.year, fin.month) else ultimo
        periodos.append({
            "anio": anio,
            "mes": mes,
            "inicio": inicio,
            "fin": corte,
            "etiqueta": f"{MESES_CORTO[mes - 1]} {anio}",
        })
        mes -= 1
        if mes == 0:
            mes = 12
            anio -= 1
    return list(reversed(periodos))


def serie_ingreso_entrega(projects: pd.DataFrame, horas: pd.DataFrame, fin: date, cantidad: int = 6) -> pd.DataFrame:
    """Horas vendidas que entran al mes (proyectos que inician) frente a horas registradas."""
    filas = []
    for periodo in periodos_recientes(fin, cantidad):
        if projects is None or projects.empty:
            ingresan = 0.0
        else:
            entran = _en_rango(projects["fecha_inicio"], periodo["inicio"], periodo["fin"])
            ingresan = float(projects.loc[entran, "horas_vendidas"].sum())
        if horas is None or horas.empty:
            entregadas = 0.0
        else:
            del_mes = horas[(horas["anio"] == periodo["anio"]) & (horas["mes"] == periodo["mes"])]
            entregadas = float(del_mes["horas"].sum())
        filas.append({
            "Mes": periodo["etiqueta"],
            "anio": periodo["anio"],
            "mes": periodo["mes"],
            "inicio": periodo["inicio"],
            "fin": periodo["fin"],
            "Horas que ingresan": _round2(ingresan),
            "Horas entregadas": _round2(entregadas),
        })
    return pd.DataFrame(filas)


def serie_apertura_cierre(projects: pd.DataFrame, fin: date, cantidad: int = 6) -> pd.DataFrame:
    filas = []
    for periodo in periodos_recientes(fin, cantidad):
        if projects is None or projects.empty:
            abiertos = cerrados = 0
        else:
            abiertos = int(_en_rango(projects["fecha_inicio"], periodo["inicio"], periodo["fin"]).sum())
            cerrados = int(_en_rango(projects["fecha_real_cierre"], periodo["inicio"], periodo["fin"]).sum())
        filas.append({
            "Mes": periodo["etiqueta"],
            "Proyectos abiertos": abiertos,
            "Proyectos cerrados": cerrados,
        })
    return pd.DataFrame(filas)


def _dias(a, b) -> float | None:
    inicio = pd.to_datetime(a, errors="coerce")
    fin = pd.to_datetime(b, errors="coerce")
    if pd.isna(inicio) or pd.isna(fin):
        return None
    return float((fin.normalize() - inicio.normalize()).days)


def salud_proyectos(projects: pd.DataFrame, horas: pd.DataFrame, fecha: date) -> pd.DataFrame:
    """Planeado contra entregado, en horas y en fechas."""
    columnas = [
        "Proyecto", "Estado", "Tipo", "Inicio", "Fin planeado", "Cierre real",
        "H. vendidas", "H. ejecutadas", "H. pendientes", "% avance",
        "Pendientes reales", "Desfase horas %", "Desfase tiempo %",
    ]
    if projects is None or projects.empty:
        return pd.DataFrame(columns=columnas)
    registradas = {}
    if horas is not None and not horas.empty:
        registradas = horas.groupby("project_id")["horas"].sum().to_dict()
    filas = []
    for rec in projects.to_dict("records"):
        vendidas = _num(rec.get("horas_vendidas"))
        ejecutadas = _num(rec.get("horas_ejecutadas"))
        pendientes = _num(rec.get("horas_pendientes"))
        entregadas = _num(registradas.get(rec.get("project_id"), 0))
        avance = ejecutadas / vendidas if vendidas else None
        reales = vendidas - entregadas
        desfase_horas = (entregadas - vendidas) / vendidas if vendidas else None
        duracion = _dias(rec.get("fecha_inicio"), rec.get("fecha_fin"))
        cierre = rec.get("fecha_real_cierre")
        if pd.notna(pd.to_datetime(cierre, errors="coerce")) and duracion:
            atraso = _dias(rec.get("fecha_fin"), cierre)
        elif duracion and pd.to_datetime(rec.get("fecha_fin"), errors="coerce") < pd.Timestamp(fecha):
            atraso = _dias(rec.get("fecha_fin"), fecha)
        else:
            atraso = 0 if duracion else None
        desfase_tiempo = (atraso / duracion) if duracion else None
        filas.append({
            "Proyecto": rec.get("proyecto_ui") or rec.get("proyecto"),
            "Estado": rec.get("etapa_proyecto"),
            "Tipo": rec.get("etiquetas") or "",
            "Inicio": rec.get("fecha_inicio"),
            "Fin planeado": rec.get("fecha_fin"),
            "Cierre real": cierre,
            "H. vendidas": _round2(vendidas),
            "H. ejecutadas": _round2(ejecutadas),
            "H. pendientes": _round2(pendientes),
            "% avance": avance,
            "Pendientes reales": _round2(reales),
            "Desfase horas %": desfase_horas,
            "Desfase tiempo %": desfase_tiempo,
        })
    return pd.DataFrame(filas, columns=columnas)


def _clave_persona(nombre: str) -> str:
    plano = unicodedata.normalize("NFD", str(nombre or ""))
    plano = "".join(c for c in plano if unicodedata.category(c) != "Mn")
    return " ".join(sorted(parte for parte in plano.casefold().replace(",", " ").split() if parte))


def productividad_backlog(acciones: pd.DataFrame, capacidad: dict[str, float]) -> pd.DataFrame:
    """Horas de backlog y planeación asignadas hoy, frente a la capacidad del mes.

    No usa el parte de horas. La salida es lo que ya está en tareas finalizadas.
    """
    columnas = [
        "Recurso", "Backlog asignado", "Planeadas", "Salida ejecutada",
        "Capacidad del mes", "Desfase vs capacidad",
    ]
    if acciones is None or acciones.empty:
        return pd.DataFrame(columns=columnas)
    cuadro = (
        acciones.pivot_table(index="usuario", columns="accion", values="cantidad", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    indice = {_clave_persona(nombre): valor for nombre, valor in (capacidad or {}).items()}
    filas = []
    for rec in cuadro.to_dict("records"):
        nombre = str(rec.get("usuario") or "-")
        backlog = _num(rec.get("H. Pendientes"))
        planeadas = _num(rec.get("H. Planeadas"))
        salida = _num(rec.get("H. Ejecutadas"))
        cupo = (capacidad or {}).get(nombre)
        if cupo is None:
            cupo = indice.get(_clave_persona(nombre))
        asignadas = backlog + planeadas
        filas.append({
            "Recurso": nombre,
            "Backlog asignado": _round2(backlog),
            "Planeadas": _round2(planeadas),
            "Salida ejecutada": _round2(salida),
            "Capacidad del mes": None if cupo is None else _round2(cupo),
            "Desfase vs capacidad": None if cupo is None else _round2(asignadas - cupo),
        })
    out = pd.DataFrame(filas, columns=columnas)
    return out.sort_values("Backlog asignado", ascending=False, kind="stable").reset_index(drop=True)


def capacidad_periodo(empleados, asistencia, ausencias, inicio: date, fin: date, tz_name: str, factor: float) -> tuple[dict[str, float], float]:
    """Capacidad de entrega (horario menos festivos y ausencias, por el factor) de cada persona."""
    por_nombre: dict[str, float] = {}
    total = 0.0
    if empleados is None or empleados.empty:
        return por_nombre, 0.0
    for rec in empleados.to_dict("records"):
        cal = rec.get("calendar_id")
        cap = capacidad_empleado(
            int(cal) if cal else None,
            int(rec["resource_id"]) if rec.get("resource_id") else None,
            asistencia or [],
            ausencias or [],
            inicio,
            fin,
            tz_name,
            _factor_persona(rec.get("factor"), factor),
        )
        por_nombre[str(rec.get("nombre") or "-")] = cap["deberia"]
        total += cap["deberia"]
    return por_nombre, _round2(total)


def _horas_del_dia(dia: date, calendar_id: int | None, asistencia: list[dict]) -> float:
    """Horas del día. Si dos franjas se solapan, se cuentan una sola vez."""
    if not calendar_id:
        return 0.0
    semana = str(dia.isocalendar()[1] % 2)
    dow = str(dia.weekday())
    tramos = []
    for slot in asistencia:
        if int(slot.get("calendar_id") or 0) != int(calendar_id):
            continue
        dia_slot = str(slot.get("dayofweek")).split(".")[0]
        if dia_slot != dow:
            continue
        tipo = slot.get("week_type")
        if tipo not in (None, False, "") and str(tipo).split(".")[0] != semana:
            continue
        inicio = _num(slot.get("hour_from"))
        fin = _num(slot.get("hour_to"))
        if fin > inicio:
            tramos.append((inicio, fin))
    if not tramos:
        return 0.0
    tramos.sort()
    actual_ini, actual_fin = tramos[0]
    total = 0.0
    for inicio, fin in tramos[1:]:
        if inicio <= actual_fin:
            actual_fin = max(actual_fin, fin)
        else:
            total += actual_fin - actual_ini
            actual_ini, actual_fin = inicio, fin
    return total + (actual_fin - actual_ini)


def _a_local(value, tz: ZoneInfo):
    momento = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(momento):
        return None
    return momento.tz_convert(tz)


def _fechas_cubiertas(inicio_local, fin_local) -> list[date]:
    if inicio_local is None or fin_local is None:
        return []
    primero = inicio_local.date()
    ultimo = fin_local.date()
    if (
        fin_local.hour == 0 and fin_local.minute == 0 and fin_local.second == 0
        and ultimo > primero
    ):
        ultimo = ultimo - timedelta(days=1)
    if ultimo < primero:
        return []
    dias = []
    cursor = primero
    while cursor <= ultimo:
        dias.append(cursor)
        cursor += timedelta(days=1)
    return dias


def _descuento_ausencia(inicio_local, fin_local, dia: date, programadas: float) -> float:
    if dia not in _fechas_cubiertas(inicio_local, fin_local):
        return 0.0
    if inicio_local.date() == fin_local.date() == dia:
        duracion = (fin_local - inicio_local).total_seconds() / 3600
        if 0 < duracion < programadas:
            return duracion
    return programadas


def _indice_ausencias(ausencias, tz: ZoneInfo, inicio: date, fin: date):
    festivos_global: set[date] = set()
    festivos_cal: dict[int, set[date]] = {}
    por_recurso: dict[int, dict] = {}
    for ausencia in ausencias or []:
        inicio_local = _a_local(ausencia.get("date_from"), tz)
        fin_local = _a_local(ausencia.get("date_to"), tz)
        dias = [dia for dia in _fechas_cubiertas(inicio_local, fin_local) if inicio <= dia <= fin]
        if not dias:
            continue
        if not ausencia.get("resource_id"):
            cal = ausencia.get("calendar_id")
            if cal:
                festivos_cal.setdefault(int(cal), set()).update(dias)
            else:
                festivos_global.update(dias)
            continue
        recurso = int(ausencia["resource_id"])
        bolsillo = por_recurso.setdefault(recurso, {})
        parcial = None
        if inicio_local and fin_local and inicio_local.date() == fin_local.date():
            duracion = (fin_local - inicio_local).total_seconds() / 3600
            if 0 < duracion < 12:
                parcial = duracion
        for dia in dias:
            if parcial is None:
                bolsillo[dia] = "completo"
            else:
                anterior = bolsillo.get(dia)
                bolsillo[dia] = parcial if anterior in (None, "completo") else float(anterior) + parcial
    return festivos_global, festivos_cal, por_recurso


def capacidad_empleado(
    calendar_id: int | None,
    resource_id: int | None,
    asistencia: list[dict],
    ausencias: list[dict],
    inicio: date,
    fin: date,
    tz_name: str,
    factor: float,
) -> dict:
    """Horas del horario, menos festivos y ausencias, multiplicadas por el factor."""
    try:
        tz = ZoneInfo(tz_name or "America/Bogota")
    except Exception:
        tz = ZoneInfo("America/Bogota")
    festivos_global, festivos_cal, por_recurso = _indice_ausencias(ausencias, tz, inicio, fin)
    propios = por_recurso.get(int(resource_id or 0), {})
    horario = festivos = personales = 0.0
    cursor = inicio
    while cursor <= fin:
        programadas = _horas_del_dia(cursor, calendar_id, asistencia)
        horario += programadas
        if programadas:
            if cursor in festivos_global or cursor in festivos_cal.get(int(calendar_id or 0), ()):
                festivos += programadas
            else:
                descuento = propios.get(cursor)
                if descuento == "completo":
                    personales += programadas
                elif descuento:
                    personales += min(float(descuento), programadas)
        cursor += timedelta(days=1)
    laborales = max(0.0, horario - festivos - personales)
    return {
        "horario": _round2(horario),
        "festivos": _round2(festivos),
        "ausencias": _round2(personales),
        "laborables": _round2(laborales),
        "deberia": _round2(laborales * factor),
    }


CAMPOS_FACTOR = ("x_studio_factor", "x_studio_factor_entrega", "x_factor")


def _factor_de_odoo(rec: dict):
    for campo in CAMPOS_FACTOR:
        if campo in rec and rec.get(campo) not in (None, False):
            return rec.get(campo)
    return None


def _factor_persona(valor, respaldo: float) -> float:
    """El Excel guarda el factor por persona. En Odoo se usa si el campo existe."""
    if valor in (None, False) or (isinstance(valor, float) and pd.isna(valor)):
        return respaldo
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return respaldo
    if numero == 0:
        return respaldo
    return numero / 100 if numero > 2 else numero


def cuadro_entrega(
    empleados: pd.DataFrame,
    asistencia: list[dict],
    ausencias: list[dict],
    horas: pd.DataFrame,
    inicio: date,
    fin: date,
    tz_name: str,
    factor: float,
    calendario_compania: int | None,
) -> pd.DataFrame:
    """Por empleado: lo que debía entregar en el mes y lo que registró."""
    columnas = [
        "empleado", "horario", "festivos", "ausencias", "laborables",
        "factor", "deberia", "entrego", "diferencia",
    ]
    if (empleados is None or empleados.empty) and (horas is None or horas.empty):
        return pd.DataFrame(columns=columnas)
    entregado: dict[int, float] = {}
    sin_id: dict[str, float] = {}
    if horas is not None and not horas.empty:
        del_mes = horas[(horas["anio"] == fin.year) & (horas["mes"] == fin.month)]
        for rec in del_mes.to_dict("records"):
            horas_reg = _num(rec.get("horas"))
            emp = rec.get("employee_id")
            if pd.notna(emp):
                entregado[int(emp)] = entregado.get(int(emp), 0.0) + horas_reg
            else:
                nombre = str(rec.get("responsable") or "-")
                sin_id[nombre] = sin_id.get(nombre, 0.0) + horas_reg
    filas = []
    nombres_empleado = {}
    if empleados is not None and not empleados.empty:
        for rec in empleados.to_dict("records"):
            emp_id = int(rec["employee_id"])
            cal = rec.get("calendar_id") or calendario_compania
            factor_emp = _factor_persona(rec.get("factor"), factor)
            cap = capacidad_empleado(
                int(cal) if cal else None,
                int(rec["resource_id"]) if rec.get("resource_id") else None,
                asistencia,
                ausencias,
                inicio,
                fin,
                tz_name,
                factor_emp,
            )
            entrego = _round2(entregado.get(emp_id, 0.0))
            nombre = str(rec.get("nombre") or "-")
            nombres_empleado[nombre] = emp_id
            filas.append({
                "empleado": nombre,
                **cap,
                "factor": factor_emp,
                "entrego": entrego,
                "diferencia": _round2(entrego - cap["deberia"]),
            })
    for nombre, cantidad in sin_id.items():
        if nombre in nombres_empleado and int(nombres_empleado[nombre]) in entregado:
            continue
        if nombre in nombres_empleado:
            for fila in filas:
                if fila["empleado"] == nombre and fila["entrego"] == 0:
                    fila["entrego"] = _round2(cantidad)
                    fila["diferencia"] = _round2(fila["entrego"] - fila["deberia"])
            continue
        cap = capacidad_empleado(
            int(calendario_compania) if calendario_compania else None,
            None,
            asistencia,
            ausencias,
            inicio,
            fin,
            tz_name,
            factor,
        )
        entrego = _round2(cantidad)
        filas.append({
            "empleado": nombre,
            **cap,
            "factor": factor,
            "entrego": entrego,
            "diferencia": _round2(entrego - cap["deberia"]),
        })
    if not filas:
        return pd.DataFrame(columns=columnas)
    out = pd.DataFrame(filas, columns=columnas)
    return out.sort_values(["diferencia", "empleado"], kind="stable").reset_index(drop=True)


def cuadro_reporte(projects: pd.DataFrame, registered: pd.DataFrame, fecha_reporte) -> pd.DataFrame:
    """Cuadro del Excel: una fila por proyecto.

    Acumulado = horas registradas desde el inicio del rango hasta el mes
    anterior a ``fecha_reporte``. Mes = horas del mes de esa fecha
    (el rango ya viene cortado el día del reporte). DeFase = contratadas
    − acumulado − mes. Mes en cero se deja vacío, como en la hoja.
    """
    columnas = [
        "project_id", "proyecto_ui", "etapa_proyecto",
        "contratadas", "acumulado", "mes", "backlog", "planning", "done", "defase",
    ]
    if projects is None or projects.empty:
        return pd.DataFrame(columns=columnas)
    out = projects.copy()
    if "proyecto_ui" not in out.columns:
        out["proyecto_ui"] = out["proyecto"]
    reg = registered if registered is not None else pd.DataFrame()
    anio = int(fecha_reporte.year)
    mes_corte = int(fecha_reporte.month)
    if not reg.empty:
        antes = (reg["anio"] < anio) | ((reg["anio"] == anio) & (reg["mes"] < mes_corte))
        del_mes = (reg["anio"] == anio) & (reg["mes"] == mes_corte)
        acum_map = reg.loc[antes].groupby("project_id")["horas"].sum()
        mes_map = reg.loc[del_mes].groupby("project_id")["horas"].sum()
    else:
        acum_map = pd.Series(dtype="float")
        mes_map = pd.Series(dtype="float")
    contratadas = out["horas_vendidas"].map(_round2)
    acumulado = out["project_id"].map(acum_map).fillna(0.0).map(_round2)
    mes = out["project_id"].map(mes_map).fillna(0.0).map(_round2)
    out["contratadas"] = contratadas
    out["acumulado"] = acumulado
    out["mes"] = mes.where(mes.abs() > 1e-9)
    out["backlog"] = out["horas_pendientes"].map(_round2)
    out["planning"] = out["horas_planeadas"].map(_round2)
    out["done"] = out["horas_ejecutadas"].map(_round2)
    out["defase"] = (contratadas - acumulado - mes).map(_round2)
    return out[columnas].sort_values("proyecto_ui", kind="stable").reset_index(drop=True)


def join_occupancy(projects: pd.DataFrame, registered: pd.DataFrame) -> pd.DataFrame:
    """Suma horas registradas al cuadro de proyectos y calcula avance y consumo."""
    if projects.empty:
        out = projects.copy()
        out["horas_registradas"] = pd.Series(dtype="float")
        out["avance"] = pd.Series(dtype="float")
        out["consumo"] = pd.Series(dtype="float")
        return out
    out = projects.copy()
    if registered is None or registered.empty:
        out["horas_registradas"] = 0.0
    else:
        summed = registered.groupby("project_id")["horas"].sum()
        out["horas_registradas"] = out["project_id"].map(summed).fillna(0.0).astype(float)
    vendidas = out["horas_vendidas"]
    out["avance"] = (out["horas_ejecutadas"] / vendidas).where(vendidas > 0)
    out["consumo"] = (out["horas_registradas"] / vendidas).where(vendidas > 0)
    return out


# ─────────────────────────────────────────────
# Conexión
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner="Conectando con Odoo...")
def get_connection():
    """Autentica una vez y comparte el lock entre sesiones.

    El lock vive dentro del recurso cacheado: un lock global del módulo se
    recrea en cada rerun de Streamlit y no serializa de verdad las llamadas.
    """
    try:
        raw = st.secrets["odoo"]
    except Exception as exc:
        raise OdooError(
            "Faltan los secrets de Odoo. En local crea `.streamlit/secrets.toml`. "
            "En Streamlit Cloud van en Manage app → Settings → Secrets. "
            "Hace falta la tabla [odoo] con url, db, username y api_key."
        ) from exc
    missing = [key for key in ("url", "db", "username", "api_key") if not raw.get(key)]
    if missing:
        raise OdooError("En [odoo] faltan: " + ", ".join(missing) + ".")
    cfg = {
        "url": str(raw["url"]).strip().rstrip("/"),
    }
    if not cfg["url"].startswith(("http://", "https://")):
        cfg["url"] = "https://" + cfg["url"]
    cfg.update({
        "db": str(raw["db"]),
        "username": str(raw["username"]),
        "api_key": str(raw["api_key"]),
        "lang": str(raw.get("lang") or "es_CO"),
        "company_id": int(raw["company_id"]) if raw.get("company_id") else None,
        "service_line": str(raw.get("service_line") or "digital_transformation"),
        "exclude_project_ids": _secret_ids(raw, "exclude_project_ids", [567]),
        "include_project_ids": _secret_ids(raw, "include_project_ids", [119]),
    })
    common = xmlrpc.client.ServerProxy(f"{cfg['url']}/xmlrpc/2/common", allow_none=True)
    lock = threading.Lock()
    try:
        with lock:
            uid = common.authenticate(cfg["db"], cfg["username"], cfg["api_key"], {})
    except Exception as exc:
        raise OdooError(f"No se pudo contactar Odoo: {exc}") from exc
    if not uid:
        raise OdooError("Autenticación fallida. Revisa url, db, username y api_key en los secrets.")
    models = xmlrpc.client.ServerProxy(f"{cfg['url']}/xmlrpc/2/object", allow_none=True)
    return cfg, int(uid), models, lock


def odoo_call(model: str, method: str, args: list, kwargs: dict | None = None):
    cfg, uid, models, lock = get_connection()
    try:
        with lock:
            return models.execute_kw(
                cfg["db"], uid, cfg["api_key"], model, method, args, kwargs or {}
            )
    except xmlrpc.client.Fault as exc:
        message = (exc.faultString or "").strip().splitlines()
        detail = message[-1] if message else str(exc)
        raise OdooError(f"Odoo rechazó {model}.{method}: {detail}") from exc
    except OdooError:
        raise
    except Exception as exc:
        raise OdooError(f"Error consultando {model}.{method}: {exc}") from exc


def _context() -> dict:
    cfg, _, _, _ = get_connection()
    return {"lang": cfg["lang"]}


def search_read(model: str, domain: list, fields: list[str], order: str = "id") -> list[dict]:
    """Lee todas las páginas. Sin `limit`, un modelo grande puede truncarse según el servidor."""
    rows: list[dict] = []
    offset = 0
    page = 400
    ctx = _context()
    while True:
        batch = odoo_call(
            model,
            "search_read",
            [domain],
            {"fields": fields, "limit": page, "offset": offset, "order": order, "context": ctx},
        )
        rows.extend(batch)
        if len(batch) < page:
            break
        offset += page
    return rows


def search_read_in(model: str, field: str, ids: list[int], domain: list, fields: list[str]) -> list[dict]:
    if not ids:
        return []
    rows: list[dict] = []
    unique = list(dict.fromkeys(int(i) for i in ids))
    for start in range(0, len(unique), 300):
        chunk = unique[start:start + 300]
        rows.extend(search_read(model, [(field, "in", chunk), *domain], fields))
    return rows


@st.cache_data(ttl=3600, show_spinner=False)
def available_fields(model: str) -> list[str]:
    info = odoo_call(model, "fields_get", [[]], {"attributes": ["type"]})
    return sorted(info.keys())


def _pick(model: str, wanted: list[str]) -> list[str]:
    have = set(available_fields(model))
    return [field for field in wanted if field in have]


def _company_id() -> int:
    cfg, uid, _, _ = get_connection()
    if cfg["company_id"]:
        return int(cfg["company_id"])
    users = search_read("res.users", [("id", "=", uid)], ["company_id"])
    if not users:
        raise OdooError("No se pudo leer la compañía del usuario. Indica company_id en los secrets.")
    company = _m2o_id(users[0].get("company_id"))
    if not company:
        raise OdooError("El usuario no tiene compañía. Indica company_id en los secrets.")
    return company


def _stage_accion() -> tuple[dict[int, str], list[dict], list[str]]:
    """res_id de project.task.type → acción, más avisos si falta algún xmlid."""
    records = search_read(
        "ir.model.data",
        [
            ("module", "=", MODULE_ETAPAS),
            ("model", "=", "project.task.type"),
            ("name", "in", list(XML_ACCION)),
        ],
        ["name", "res_id"],
    )
    found = {rec["name"]: int(rec["res_id"]) for rec in records if rec.get("res_id")}
    warnings = []
    missing = [name for name in XML_ACCION if name not in found]
    if missing:
        warnings.append(
            "No están estos xmlid de etapa en ir.model.data: " + ", ".join(missing) + "."
        )
    if not found:
        raise OdooError(
            "No se encontraron las etapas de l10n_co_firefly_project. "
            "Sin esos xmlid no se puede separar pendientes, planeadas y ejecutadas."
        )
    stage_accion = {res_id: XML_ACCION[name] for name, res_id in found.items()}
    names = search_read(
        "project.task.type",
        [("id", "in", list(stage_accion))],
        ["name"],
    )
    detail = []
    name_by_id = {int(rec["id"]): _text(rec.get("name"), "") for rec in names}
    for xml_name, res_id in sorted(found.items()):
        detail.append({
            "xmlid": f"{MODULE_ETAPAS}.{xml_name}",
            "etapa_id": res_id,
            "etapa": name_by_id.get(res_id, ""),
            "accion": XML_ACCION[xml_name],
        })
    return stage_accion, detail, warnings


def _tag_labels(tag_ids: list[int]) -> dict[int, str]:
    if not tag_ids:
        return {}
    records = search_read("project.tags", [("id", "in", list(dict.fromkeys(tag_ids)))], ["name"])
    labels = {}
    for rec in records:
        labels[int(rec["id"])] = _text(rec.get("name"), "")
    return labels


def _with_etiquetas(projects: list[dict]) -> list[dict]:
    all_tags: list[int] = []
    for project in projects:
        all_tags.extend(_id_list(project.get("tag_ids")))
    labels = _tag_labels(all_tags)
    for project in projects:
        names = sorted({labels[i] for i in _id_list(project.get("tag_ids")) if labels.get(i)})
        project["_etiquetas"] = ", ".join(names)
    return projects


def _fetch_projects(domain: list) -> tuple[list[dict], list[str]]:
    fields = _pick("project.project", ["id", *PROJECT_FIELDS])
    warnings = []
    have = set(fields)
    for studio in ("x_studio_vendedor_1", "x_studio_rapidez", "x_studio_fecha_cierre", "allocated_hours"):
        if studio not in have:
            warnings.append(f"project.project no tiene `{studio}`; esa columna queda vacía o en cero.")
    if "service_line" not in available_fields("project.project"):
        raise OdooError(
            "project.project no tiene service_line. Hace falta el módulo l10n_co_firefly_project."
        )
    records = search_read("project.project", domain, fields, order="name, id")
    return _with_etiquetas(records), warnings


def _fetch_tasks(project_ids: list[int], stage_ids: list[int]) -> list[dict]:
    if not project_ids or not stage_ids:
        return []
    fields = _pick("project.task", ["id", *TASK_FIELDS])
    domain = [
        ("parent_id", "=", False),
        ("active", "=", True),
        ("stage_id", "in", stage_ids),
    ]
    return search_read_in("project.task", "project_id", project_ids, domain, fields)


def _user_names(tasks: list[dict]) -> dict[int, str]:
    ids: list[int] = []
    for task in tasks:
        ids.extend(_id_list(task.get("user_ids")))
        solo = _m2o_id(task.get("user_id"))
        if solo:
            ids.append(solo)
    if not ids:
        return {}
    records = search_read("res.users", [("id", "in", list(dict.fromkeys(ids)))], ["name"])
    return {int(rec["id"]): _text(rec.get("name")) for rec in records}


@st.cache_data(ttl=600, show_spinner="Cargando proyectos y tareas de Transformación Digital...")
def load_snapshot() -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    cfg, _, _, _ = get_connection()
    company_id = _company_id()
    stage_accion, stage_detail, warnings = _stage_accion()
    domain = [
        ("company_id", "=", company_id),
        ("active", "=", True),
        ("service_line", "=", cfg["service_line"]),
    ]
    projects, field_warnings = _fetch_projects(domain)
    warnings = warnings + field_warnings
    tasks = _fetch_tasks([int(p["id"]) for p in projects], list(stage_accion))
    users = _user_names(tasks)
    project_df = build_projects(projects, tasks, stage_accion)
    action_df = build_actions(project_df, tasks, stage_accion, users)
    info = {
        "company_id": company_id,
        "service_line": cfg["service_line"],
        "etapas": stage_detail,
        "warnings": warnings,
        "n_tareas": len(tasks),
        "exclude_project_ids": cfg["exclude_project_ids"],
        "include_project_ids": cfg["include_project_ids"],
    }
    return project_df, action_df, info


def _timesheet_projects(company_id: int) -> dict[int, str]:
    """Ids que entran a horas registradas → nombre. Incluye archivados, como la consulta SQL."""
    cfg, _, _, _ = get_connection()
    domain = [
        ("company_id", "=", company_id),
        ("service_line", "=", cfg["service_line"]),
        "|", ("active", "=", True), ("active", "=", False),
    ]
    records = search_read("project.project", domain, ["name"])
    allowed = {}
    excluded = set(cfg["exclude_project_ids"])
    for rec in records:
        pid = int(rec["id"])
        if pid in excluded:
            continue
        allowed[pid] = _text(rec.get("name"), "Sin nombre")
    extras = [pid for pid in cfg["include_project_ids"] if pid not in allowed and pid not in excluded]
    if extras:
        extra_recs = search_read(
            "project.project",
            [
                ("id", "in", extras),
                ("company_id", "=", company_id),
                "|", ("active", "=", True), ("active", "=", False),
            ],
            ["name"],
        )
        for rec in extra_recs:
            allowed[int(rec["id"])] = _text(rec.get("name"), "Sin nombre")
    return allowed


def _analytic_lines(project_ids: list[int], date_from: str, date_to: str) -> tuple[list[dict], bool]:
    fields = _pick(
        "account.analytic.line",
        ["id", "date", "unit_amount", "project_id", "task_id", "employee_id", "user_id"],
    )
    date_domain = [("date", ">=", date_from), ("date", "<=", date_to)]
    by_project = search_read_in(
        "account.analytic.line", "project_id", project_ids, date_domain, fields
    )
    by_task = []
    task_domain_ok = True
    try:
        by_task = search_read_in(
            "account.analytic.line", "task_id.project_id", project_ids, date_domain, fields
        )
    except OdooError:
        task_domain_ok = False
    merged: dict[int, dict] = {}
    for rec in [*by_project, *by_task]:
        merged[int(rec["id"])] = rec
    return list(merged.values()), task_domain_ok


def _resolve_line_projects(lines: list[dict], allowed: dict[int, str]) -> tuple[list[dict], list[str]]:
    warnings = []
    task_ids = []
    for line in lines:
        task_id = _m2o_id(line.get("task_id"))
        if task_id:
            task_ids.append(task_id)
    task_project: dict[int, int] = {}
    if task_ids:
        tasks = search_read_in("project.task", "id", task_ids, [], ["project_id"])
        for task in tasks:
            project_id = _m2o_id(task.get("project_id"))
            if project_id:
                task_project[int(task["id"])] = project_id
    rows = []
    for line in lines:
        task_id = _m2o_id(line.get("task_id"))
        project_id = task_project.get(task_id) if task_id else None
        if not project_id:
            project_id = _m2o_id(line.get("project_id"))
        if project_id not in allowed:
            continue
        responsable = _m2o_name(line.get("employee_id"), "")
        if not responsable:
            responsable = _m2o_name(line.get("user_id"))
        rows.append({
            "project_id": project_id,
            "proyecto": allowed.get(project_id, "Sin nombre"),
            "responsable": responsable or "-",
            "employee_id": _m2o_id(line.get("employee_id")),
            "date": line.get("date") or None,
            "horas": _num(line.get("unit_amount")),
        })
    return rows, warnings


@st.cache_data(ttl=600, show_spinner="Cargando horas registradas...")
def load_registered(date_from: str, date_to: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    cfg, _, _, _ = get_connection()
    company_id = _company_id()
    allowed = _timesheet_projects(company_id)
    warnings: list[str] = []
    if not allowed:
        vacio_emp = pd.DataFrame(columns=EMPLEADO_COLS)
        return pd.DataFrame(columns=HOURS_COLS), vacio_emp, {
            "warnings": ["No hay proyectos para las horas registradas."],
            "task_domain_ok": True,
            "allowed_ids": [],
            "company_id": company_id,
            "exclude_project_ids": cfg["exclude_project_ids"],
            "include_project_ids": cfg["include_project_ids"],
        }
    lines, task_domain_ok = _analytic_lines(list(allowed), date_from, date_to)
    if not task_domain_ok:
        warnings.append(
            "No se pudo filtrar por el proyecto de la tarea. "
            "Las horas usan el proyecto de la línea analítica."
        )
    resolved, more = _resolve_line_projects(lines, allowed)
    warnings.extend(more)
    hours = build_registered(resolved)
    por_empleado = build_horas_empleado(resolved)
    return hours, por_empleado, {
        "warnings": warnings,
        "task_domain_ok": task_domain_ok,
        "allowed_ids": sorted(allowed),
        "company_id": company_id,
        "exclude_project_ids": cfg["exclude_project_ids"],
        "include_project_ids": cfg["include_project_ids"],
    }


def _blank_project(project_id: int, nombre: str) -> dict:
    row = {col: None for col in PROJECT_COLS}
    row.update({
        "project_id": project_id,
        "proyecto": nombre,
        "etapa_proyecto": "-",
        "etiquetas": "",
        "horas_vendidas": 0.0,
        "vendedor": "-",
        "gerente": "-",
        "rapidez": 0.0,
        "horas_tareas": 0.0,
        "horas_pendientes": 0.0,
        "horas_planeadas": 0.0,
        "horas_ejecutadas": 0.0,
        "horas_bloqueadas": 0.0,
        "horas_cerradas_en_curso": 0.0,
        "activo": True,
    })
    return row


@st.cache_data(ttl=600, show_spinner=False)
def load_extra_projects(missing_ids: tuple[int, ...], names: tuple[tuple[int, str], ...]) -> pd.DataFrame:
    """Proyectos con horas registradas que no están en el snapshot activo de TD."""
    if not missing_ids:
        return pd.DataFrame(columns=PROJECT_COLS)
    name_by_id = dict(names)
    records, _warnings = _fetch_projects([
        ("id", "in", list(missing_ids)),
        "|", ("active", "=", True), ("active", "=", False),
    ])
    frame = build_projects(records, [], {})
    found = set(frame["project_id"]) if not frame.empty else set()
    stubs = [
        _blank_project(pid, name_by_id.get(pid, "Sin nombre"))
        for pid in missing_ids
        if pid not in found
    ]
    if stubs:
        frame = pd.concat([frame, pd.DataFrame(stubs)], ignore_index=True)
    return frame


def _slots(records: list[dict]) -> list[dict]:
    slots = []
    for rec in records:
        slots.append({
            "calendar_id": _m2o_id(rec.get("calendar_id")),
            "dayofweek": rec.get("dayofweek"),
            "hour_from": _num(rec.get("hour_from")),
            "hour_to": _num(rec.get("hour_to")),
            "week_type": rec.get("week_type") if rec.get("week_type") not in (False, None, "") else None,
        })
    return slots


def _ausencias(records: list[dict]) -> list[dict]:
    filas = []
    for rec in records:
        if rec.get("time_type") == "other":
            continue
        filas.append({
            "calendar_id": _m2o_id(rec.get("calendar_id")),
            "resource_id": _m2o_id(rec.get("resource_id")),
            "date_from": rec.get("date_from") or None,
            "date_to": rec.get("date_to") or None,
        })
    return filas


@st.cache_data(ttl=600, show_spinner="Cargando horarios y festivos...")
def load_plantilla(
    employee_ids: tuple[int, ...],
    nombres: tuple[str, ...],
    date_from: str,
    date_to: str,
) -> dict:
    """Calendario laboral, festivos y ausencias de los empleados del mes."""
    company_id = _company_id()
    warnings: list[str] = []
    vacio = {
        "empleados": pd.DataFrame(columns=["employee_id", "nombre", "calendar_id", "resource_id", "factor"]),
        "asistencia": [],
        "ausencias": [],
        "tz": "America/Bogota",
        "calendario_compania": None,
        "warnings": warnings,
    }
    try:
        compania = search_read("res.company", [("id", "=", company_id)], ["resource_calendar_id"])
    except OdooError as exc:
        warnings.append(str(exc))
        return vacio
    calendario_compania = _m2o_id(compania[0].get("resource_calendar_id")) if compania else None
    tz = "America/Bogota"
    if calendario_compania:
        try:
            cal = search_read("resource.calendar", [("id", "=", calendario_compania)], _pick("resource.calendar", ["tz", "hours_per_day"]))
            if cal and cal[0].get("tz"):
                tz = str(cal[0]["tz"])
        except OdooError:
            pass

    emp_fields = _pick(
        "hr.employee",
        ["name", "resource_calendar_id", "resource_id", "company_id", *CAMPOS_FACTOR],
    )
    records: list[dict] = []
    nombres_ok = tuple(n for n in nombres if n and n != "-")
    try:
        if employee_ids and "name" in emp_fields:
            records.extend(search_read("hr.employee", [("id", "in", list(employee_ids))], emp_fields))
        if nombres_ok and "name" in emp_fields:
            records.extend(search_read(
                "hr.employee",
                [("company_id", "=", company_id), ("name", "in", list(nombres_ok))],
                emp_fields,
            ))
    except OdooError as exc:
        warnings.append(f"No se pudo leer el horario de los empleados: {exc}")
        return vacio
    if "resource_calendar_id" not in emp_fields:
        warnings.append("hr.employee no tiene horario laboral. Las horas esperadas quedan en cero.")
    if not any(campo in emp_fields for campo in CAMPOS_FACTOR):
        warnings.append(
            "Odoo no tiene factor por empleado. Se usa el factor del panel para todo el equipo."
        )

    unicos: dict[int, dict] = {}
    for rec in records:
        unicos[int(rec["id"])] = rec
    empleados = []
    for emp_id, rec in unicos.items():
        cal = _m2o_id(rec.get("resource_calendar_id")) or calendario_compania
        empleados.append({
            "employee_id": emp_id,
            "nombre": _text(rec.get("name")),
            "calendar_id": cal,
            "resource_id": _m2o_id(rec.get("resource_id")),
            "factor": _factor_de_odoo(rec),
        })
    empleados_df = pd.DataFrame(
        empleados, columns=["employee_id", "nombre", "calendar_id", "resource_id", "factor"]
    )

    calendarios = sorted({
        int(c) for c in empleados_df["calendar_id"].dropna().tolist() if c
    } | ({int(calendario_compania)} if calendario_compania else set()))
    asistencia: list[dict] = []
    if calendarios:
        try:
            slots = search_read(
                "resource.calendar.attendance",
                [("calendar_id", "in", calendarios)],
                _pick("resource.calendar.attendance", ["calendar_id", "dayofweek", "hour_from", "hour_to", "week_type"]),
            )
            asistencia = _slots(slots)
        except OdooError as exc:
            warnings.append(f"No se pudo leer el horario: {exc}")
    cubiertos = {int(s["calendar_id"]) for s in asistencia if s.get("calendar_id")}
    for cal in calendarios:
        if cal in cubiertos:
            continue
        for dow in range(5):
            asistencia.append({
                "calendar_id": cal,
                "dayofweek": str(dow),
                "hour_from": 8.0,
                "hour_to": 16.0,
                "week_type": None,
            })
        warnings.append(
            f"El horario {cal} no tiene franjas. Se usó lunes a viernes, 8 horas, como respaldo."
        )

    recursos = sorted({int(r) for r in empleados_df["resource_id"].dropna().tolist() if r})
    leave_domain = [
        ("date_from", "<=", f"{date_to} 23:59:59"),
        ("date_to", ">=", date_from),
    ]
    if recursos:
        leave_domain = [
            *leave_domain,
            "|", ("resource_id", "=", False), ("resource_id", "in", recursos),
        ]
    else:
        leave_domain.append(("resource_id", "=", False))
    ausencias: list[dict] = []
    try:
        leaves = search_read(
            "resource.calendar.leaves",
            leave_domain,
            _pick("resource.calendar.leaves", ["calendar_id", "resource_id", "date_from", "date_to", "time_type"]),
        )
        ausencias = _ausencias(leaves)
    except OdooError as exc:
        warnings.append(f"No se pudieron leer festivos ni ausencias: {exc}")

    return {
        "empleados": empleados_df,
        "asistencia": asistencia,
        "ausencias": ausencias,
        "tz": tz,
        "calendario_compania": calendario_compania,
        "warnings": warnings,
    }


@st.cache_data(ttl=600, show_spinner="Cargando aperturas y cierres del mes...")
def load_movimientos(inicio: str, fin: str) -> tuple[pd.DataFrame, list[str]]:
    """Proyectos abiertos o cerrados entre dos fechas, incluidos los archivados."""
    cfg, _, _, _ = get_connection()
    company_id = _company_id()
    warnings: list[str] = []
    have = set(available_fields("project.project"))
    cierre = "x_studio_fecha_cierre" if "x_studio_fecha_cierre" in have else "date"
    if cierre == "date":
        warnings.append(
            "No está x_studio_fecha_cierre. El cierre del mes usa la fecha fin del proyecto."
        )
    rango_cierre = ["&", (cierre, ">=", inicio), (cierre, "<=", fin)]
    domain = [
        ("company_id", "=", company_id),
        ("service_line", "=", cfg["service_line"]),
        "|", ("active", "=", True), ("active", "=", False),
        "|", "|",
        "&", ("date_start", ">=", inicio), ("date_start", "<=", fin),
        *rango_cierre,
        "&", ("date", ">=", inicio), ("date", "<=", fin),
    ]
    try:
        records, avisos = _fetch_projects(domain)
    except OdooError as exc:
        return pd.DataFrame(columns=PROJECT_COLS), warnings + [str(exc)]
    warnings.extend(avisos)
    frame = build_projects(records, [], {})
    if cierre == "date" and not frame.empty:
        frame = frame.copy()
        frame["fecha_real_cierre"] = frame["fecha_fin"]
    return frame, warnings


@st.cache_data(ttl=600, show_spinner="Cargando horas entregadas del equipo...")
def load_horas_equipo(employee_ids: tuple[int, ...], date_from: str, date_to: str) -> pd.DataFrame:
    """Horas de parte de horas del empleado en cualquier proyecto, no solo TD.

    La columna Entregó compara contra el horario del mes. Si solo se suman los
    proyectos de Transformación Digital, alguien que registró el resto en
    soporte u otra línea aparece con casi cero.
    """
    if not employee_ids:
        return pd.DataFrame(columns=EMPLEADO_COLS)
    fields = _pick("account.analytic.line", ["date", "unit_amount", "employee_id", "project_id", "task_id"])
    domain = [
        ("date", ">=", date_from),
        ("date", "<=", date_to),
        ("unit_amount", "!=", 0),
        "|", ("project_id", "!=", False), ("task_id", "!=", False),
    ]
    try:
        lineas = search_read_in("account.analytic.line", "employee_id", list(employee_ids), domain, fields)
    except OdooError:
        lineas = search_read_in(
            "account.analytic.line", "employee_id", list(employee_ids),
            [("date", ">=", date_from), ("date", "<=", date_to), ("unit_amount", "!=", 0)],
            [campo for campo in fields if campo != "task_id"],
        )
    rows = []
    for linea in lineas:
        emp = _m2o_id(linea.get("employee_id"))
        if not emp:
            continue
        rows.append({
            "employee_id": emp,
            "project_id": _m2o_id(linea.get("project_id")) or 0,
            "responsable": _m2o_name(linea.get("employee_id"), "-"),
            "date": linea.get("date") or None,
            "horas": _num(linea.get("unit_amount")),
        })
    return build_horas_empleado(rows)


@st.cache_data(ttl=600, show_spinner=False)
def load_horas_facturadas(date_from: str, date_to: str) -> tuple[pd.DataFrame, str | None]:
    """Horas de parte que ya tienen factura, por mes de la fecha del parte."""
    vacio = pd.DataFrame(columns=["anio", "mes", "horas"])
    cfg, _, _, _ = get_connection()
    campos = _pick("account.analytic.line", ["date", "unit_amount", "timesheet_invoice_id", "project_id"])
    if "timesheet_invoice_id" not in campos:
        return vacio, "Odoo no marca qué horas del parte ya se facturaron."
    dominio = [
        ("date", ">=", date_from),
        ("date", "<=", date_to),
        ("unit_amount", ">", 0),
        ("timesheet_invoice_id", "!=", False),
        ("project_id.service_line", "=", cfg["service_line"]),
    ]
    try:
        lineas = search_read("account.analytic.line", dominio, ["date", "unit_amount"])
    except OdooError as exc:
        return vacio, str(exc)
    if not lineas:
        return vacio, None
    frame = pd.DataFrame(lineas)
    frame["date"] = pd.to_datetime(frame["date"], errors="coerce")
    frame = frame[frame["date"].notna()]
    if frame.empty:
        return vacio, None
    frame["anio"] = frame["date"].dt.year.astype(int)
    frame["mes"] = frame["date"].dt.month.astype(int)
    frame["horas"] = frame["unit_amount"].map(_num)
    agrupado = frame.groupby(["anio", "mes"], as_index=False)["horas"].sum()
    agrupado["horas"] = agrupado["horas"].map(_round2)
    return agrupado, None

