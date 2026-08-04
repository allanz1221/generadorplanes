from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
import re

from flask import Flask, render_template, request, send_file
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

app = Flask(__name__)
DAY_NAMES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# Síntesis de los contenidos y actividades del PDF proporcionado.
ACTIVITIES = [
    ("Ruta de aprendizaje front-end", "Apuntes de clase sobre la ruta de aprendizaje de un desarrollador front-end y back-end."),
    ("Editor de texto e IDE", "Video-tutorial de instalación y configuraciones principales de VS Code u otro IDE."),
    ("Atajos y Emmet", "Exposición en equipo sobre atajos de teclado y complementos Emmet."),
    ("HTML y CSS", "Ejercicio individual: maquetar la estructura básica de un sitio web."),
    ("Diseño responsivo", "Crear una página web responsiva para smartphone, tablet y escritorio."),
    ("JSON", "Participar individualmente en el foro sobre JSON y argumentar tres réplicas."),
    ("JavaScript", "Resolver ejercicios de fundamentos, objetos, funciones, clases, asincronía, DOM y APIs."),
    ("Bootstrap: fundamentos", "Elaborar apuntes de clase sobre los fundamentos del framework Bootstrap."),
    ("Bootstrap: sitio web", "Desarrollar un sitio web usando Customize, Layout y Content de Bootstrap."),
    ("Bootstrap: componentes", "Desarrollar un sitio usando contenedores, componentes y formularios de Bootstrap."),
    ("Bootstrap: helpers y utilidades", "Elaborar un mapa mental de Helpers, Utilities y Bootstrap Icons."),
    ("Material UI: fundamentos", "Elaborar un cuadro sinóptico sobre los fundamentos de Material UI."),
    ("Material UI: aplicación", "Crear un sitio web utilizando el framework Material UI."),
    ("Tailwind CSS", "Investigar qué es Tailwind CSS y sus diferencias con Bootstrap."),
    ("Tailwind CSS: práctica", "Desarrollar un sitio web utilizando Tailwind CSS."),
    ("Vue JS: introducción", "Instalar Vue JS y ejecutar la primera aplicación."),
    ("Vue JS: estructura", "Investigar la estructura de un proyecto en Vue JS."),
    ("Vue JS: reactividad", "Implementar reactividad y crear componentes con Vue JS."),
    ("Vue JS: formularios y router", "Crear formulario, agregar validaciones y configurar Vue Router."),
    ("Vue JS: Pinia y Watch", "Practicar el uso de Pinia y Watch para almacenar estados."),
    ("Vue JS: peticiones HTTP", "Realizar peticiones HTTP asíncronas en Vue JS."),
    ("Vue JS: comunicación", "Implementar comunicación entre componentes mediante Props."),
    ("Vue JS: propiedades y directivas", "Usar propiedades computadas, filtros y directivas condicionales e interactivas."),
    ("Proyecto integrador Vue", "Definir y presentar un proyecto integrador tipo CRUD usando Vue."),
]


def scheduled_dates(start: date, end: date, day_indexes: set[int]) -> list[date]:
    current, dates = start, []
    while current <= end:
        if current.weekday() in day_indexes:
            dates.append(current)
        current += timedelta(days=1)
    return dates


def clean_activities(raw: str) -> list[tuple[str, str]]:
    rows = []
    for line in raw.splitlines():
        line = line.strip()
        if line:
            topic, separator, activity = line.partition("|")
            rows.append((topic.strip(), activity.strip() if separator and activity.strip() else "Actividad de aprendizaje por definir."))
    return rows or ACTIVITIES


def shade(cell, color: str) -> None:
    props = cell._tc.get_or_add_tcPr()
    fill = OxmlElement("w:shd")
    fill.set(qn("w:fill"), color)
    props.append(fill)


def set_cell_text(cell, value: str, bold=False, color=None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(value)
    run.bold, run.font.name, run.font.size = bold, "Arial", Pt(9)
    if color:
        run.font.color.rgb = RGBColor(*color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def make_docx(course: str, teacher: str, semester: str, hours: str, sessions: list[dict]) -> BytesIO:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(1.5)
    section.left_margin = section.right_margin = Cm(1.6)
    doc.styles["Normal"].font.name, doc.styles["Normal"].font.size = "Arial", Pt(9)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("PLAN DE CLASE")
    run.bold, run.font.name, run.font.size = True, "Arial", Pt(15)
    run.font.color.rgb = RGBColor(20, 76, 130)
    details = doc.add_paragraph()
    details.alignment = WD_ALIGN_PARAGRAPH.CENTER
    details.add_run(f"Curso: {course or 'Sin especificar'}\n").bold = True
    details.add_run(f"Docente: {teacher or 'Sin especificar'}  |  Periodo: {semester or 'Sin especificar'}  |  Horas por semana: {hours}")
    for session in sessions:
        table = doc.add_table(rows=2, cols=4)
        table.style = "Table Grid"
        labels = ["Sesión", "Fecha", "Tema", "Actividades de aprendizaje"]
        for cell, label in zip(table.rows[0].cells, labels):
            shade(cell, "164C82")
            set_cell_text(cell, label, True, (255, 255, 255))
        set_cell_text(table.cell(1, 0), str(session["number"]))
        set_cell_text(table.cell(1, 1), session["date"])
        set_cell_text(table.cell(1, 2), session["topic"])
        set_cell_text(table.cell(1, 3), session["activity"])
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(2)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Generado con Generador de Planes de Clase").font.size = Pt(8)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


@app.get("/")
def index():
    catalog = "\n".join(f"{topic} | {activity}" for topic, activity in ACTIVITIES)
    return render_template("index.html", catalog=catalog, days=enumerate(DAY_NAMES))


@app.post("/generar")
def generate():
    try:
        start = datetime.strptime(request.form["start_date"], "%Y-%m-%d").date()
        end = datetime.strptime(request.form["end_date"], "%Y-%m-%d").date()
    except (KeyError, ValueError):
        return "Selecciona fechas válidas.", 400
    if end < start:
        return "La fecha de fin debe ser igual o posterior a la fecha de inicio.", 400
    days = {int(day) for day in request.form.getlist("days")}
    if not days:
        return "Selecciona por lo menos un día de clase.", 400
    class_dates = scheduled_dates(start, end, days)
    if not class_dates:
        return "No hay clases en los días seleccionados dentro de ese periodo.", 400
    activities = clean_activities(request.form.get("activities", ""))
    sessions = [{"number": index, "date": class_date.strftime("%d/%m/%Y"), "topic": activities[(index - 1) % len(activities)][0], "activity": activities[(index - 1) % len(activities)][1]} for index, class_date in enumerate(class_dates, 1)]
    docx = make_docx(request.form.get("course", ""), request.form.get("teacher", ""), request.form.get("semester", ""), request.form.get("weekly_hours", ""), sessions)
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", request.form.get("course", "plan_clase")).strip("_") or "plan_clase"
    return send_file(docx, as_attachment=True, download_name=f"{name}_plan_de_clase.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
