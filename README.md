# Ocupación de proyectos · Transformación Digital

Tablero Streamlit que reemplaza los Excel de ocupación. Lee Odoo 19 por XML-RPC
con los secrets de la app (no hay archivos de credenciales en el código).

[Cómo desplegarlo en Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud).

## Secrets

Local: copia `.streamlit/secrets.toml.example` a `.streamlit/secrets.toml`.

En Cloud: Manage app → Settings → Secrets, misma tabla `[odoo]`.

| Clave | Uso |
| --- | --- |
| `url`, `db`, `username`, `api_key` | Conexión XML-RPC. La api key sale de Odoo → usuario → Seguridad de la cuenta. |
| `company_id` | Compañía de las tres consultas. Si falta, se usa la del usuario. |
| `service_line` | Por defecto `digital_transformation`. |
| `lang` | Nombres traducidos. Por defecto `es_CO`. |
| `exclude_project_ids` | No entran a horas registradas. Por defecto `[567]` (GS Gestión Soporte). |
| `include_project_ids` | Sí entran a horas registradas aunque no sean TD. Por defecto `[119]` (Customer Care). |

## Qué muestra

Un cuadro, una fila por proyecto:

| Columna | Origen |
| --- | --- |
| Contratadas | Horas vendidas del proyecto |
| Acumulado | Horas registradas desde el 1 de enero del año inicial hasta el mes anterior a la fecha del reporte |
| Mes | Horas registradas en el mes de la fecha del reporte, hasta ese día |
| BackLog | Tareas raíz en Inicio |
| Planning | Tareas raíz en Planeado o En ejecución, sin Hecho ni Cancelado |
| Done | Tareas raíz en Finalizado |
| DeFase | Contratadas − Acumulado − Mes |

La fecha del reporte arranca en el día en que se abre el tablero. Al seleccionar una fila se despliega vendedor, gerente, fechas, recursos y el detalle de horas registradas.

Los proyectos con la etiqueta **Bolsa de Horas** van en la pestaña «Bolsas de horas». El resto queda en «Proyectos».

La pestaña **Mes** compara, por empleado, lo que debía entregar y lo que registró. Lo debido es el horario laboral menos festivos y ausencias, multiplicado por `factor_entrega`. También lista los proyectos abiertos en el mes, los cerrados en el mes y los que siguen en proceso.

## Cómo correrlo

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
