from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from datetime import datetime
import json
from pathlib import Path
from io import BytesIO

from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

app = Flask(__name__, template_folder="templates", static_folder="static")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

LOGO_ESCOLA_PATH = STATIC_DIR / "logo_escola.png"
LOGO_ESTADO_PATH = STATIC_DIR / "logo_estado.png"

with open(DATA_DIR / "alunos.json", "r", encoding="utf-8") as f:
    estudantes = json.load(f)

with open(DATA_DIR / "config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

registros = []


def buscar_estudante(estudante_id):
    return next((e for e in estudantes if e["id"] == estudante_id), None)


def buscar_registro(registro_id):
    return next((r for r in registros if r["id"] == registro_id), None)


def formatar_lista(lista):
    if not lista:
        return ""
    if len(lista) == 1:
        return lista[0].lower()
    if len(lista) == 2:
        return f"{lista[0].lower()} e {lista[1].lower()}"
    return ", ".join(item.lower() for item in lista[:-1]) + f" e {lista[-1].lower()}"


def gerar_relatorio(estudante, registro):
    tipo_relatorio = registro["tipo_relatorio"]
    disciplina = registro["disciplina"]
    dificuldades = formatar_lista(registro["dificuldades"])
    comportamentos = formatar_lista(registro["questoes_comportamentais"])
    intervencoes = formatar_lista(registro["intervencoes"])
    encaminhamentos = formatar_lista(registro["encaminhamentos"])
    resposta = registro["resposta_aluno"].lower() if registro["resposta_aluno"] else "não informada"
    observacao = registro["observacao_livre"].strip()

    partes = [
        f"{tipo_relatorio} referente ao(à) estudante {estudante['nome']}, da turma {estudante['turma']}, "
        f"registrado em {registro['data']} às {registro['hora']}."
    ]

    if tipo_relatorio == "Relatório de Ocorrência Disciplinar":
        partes.append(f"O registro foi realizado no contexto da disciplina de {disciplina}.")
        if registro["questoes_comportamentais"]:
            partes.append(
                f"Foram observadas as seguintes ocorrências comportamentais: {comportamentos}."
            )
        if registro["intervencoes"]:
            partes.append(
                f"Diante da situação, foram adotadas as seguintes intervenções: {intervencoes}."
            )
        partes.append(f"Após a intervenção, verificou-se que o(a) estudante {resposta}.")
        if registro["encaminhamentos"]:
            partes.append(f"Como encaminhamento, recomenda-se {encaminhamentos}.")
        if observacao:
            partes.append(f"Observação complementar: {observacao}")
    else:
        partes.append(f"O registro foi realizado no contexto da disciplina de {disciplina}.")
        if registro["dificuldades"]:
            partes.append(
                f"No âmbito da aprendizagem, foram observadas dificuldades relacionadas a {dificuldades}."
            )
        if registro["questoes_comportamentais"]:
            partes.append(
                f"No aspecto comportamental, foram identificadas situações como {comportamentos}."
            )
        if registro["intervencoes"]:
            partes.append(
                f"Como intervenções pedagógicas, foram realizadas ações como {intervencoes}."
            )
        partes.append(f"Após as intervenções, verificou-se que o(a) estudante {resposta}.")
        if registro["encaminhamentos"]:
            partes.append(f"Como encaminhamento, recomenda-se {encaminhamentos}.")
        if observacao:
            partes.append(f"Observação complementar: {observacao}")

    return " ".join(partes)


def nome_arquivo_base(estudante, registro):
    base = f"{estudante['nome']}_{registro['tipo_relatorio']}_{registro['data']}"
    return base.replace("/", "-").replace(" ", "_")


def texto_completo_exportacao(estudante, registro):
    linhas = [
        "ESTADO DO PARANÁ",
        "SECRETARIA DE ESTADO DA EDUCAÇÃO",
        "PARANÁ INTEGRAL",
        "ESCOLA ESTADUAL PADRE MANUEL DA NÓBREGA",
        "",
        "Nº ________/2026",
        registro["tipo_relatorio"].upper(),
        "",
        f"Estudante: {estudante['nome']}",
        f"Turma: {estudante['turma']}",
        f"Responsável: {estudante['responsavel']}",
        f"Data: {registro['data']}",
        f"Hora: {registro['hora']}",
        f"Profissional: {registro['profissional']}",
        f"Disciplina: {registro['disciplina']}",
        "",
        registro["relatorio_texto"],
        "",
        "",
        "________________________________________",
        f"Assinatura do(a) estudante: {registro.get('assinatura_estudante', '')}",
        "",
        "________________________________________",
        f"Assinatura do(a) professor(a)/pedagogo(a): {registro.get('assinatura_profissional', '')}",
        "",
        "________________________________________",
        f"Assinatura do responsável: {registro.get('assinatura_responsavel', '')}",
    ]
    return "\n".join(linhas)


def inserir_logos_doc(document):
    table = document.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p_left = left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_ESTADO_PATH.exists():
        run = p_left.add_run()
        run.add_picture(str(LOGO_ESTADO_PATH), width=Inches(1.6))

    p_right = right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_ESCOLA_PATH.exists():
        run = p_right.add_run()
        run.add_picture(str(LOGO_ESCOLA_PATH), width=Inches(1.1))


def gerar_docx(estudante, registro):
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    inserir_logos_doc(doc)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        "ESTADO DO PARANÁ\n"
        "SECRETARIA DE ESTADO DA EDUCAÇÃO\n"
        "PARANÁ INTEGRAL\n"
        "ESCOLA ESTADUAL PADRE MANUEL DA NÓBREGA"
    )
    r.bold = True
    r.font.size = Pt(12)

    num = doc.add_paragraph()
    num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    num_run = num.add_run("Nº ________/2026")
    num_run.bold = True
    num_run.font.size = Pt(11)

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo_run = titulo.add_run(registro["tipo_relatorio"].upper())
    titulo_run.bold = True
    titulo_run.font.size = Pt(12)

    doc.add_paragraph(f"Estudante: {estudante['nome']}")
    doc.add_paragraph(f"Turma: {estudante['turma']}")
    doc.add_paragraph(f"Responsável: {estudante['responsavel']}")
    doc.add_paragraph(f"Data: {registro['data']}    Hora: {registro['hora']}")
    doc.add_paragraph(f"Profissional: {registro['profissional']}")
    doc.add_paragraph(f"Disciplina: {registro['disciplina']}")
    doc.add_paragraph("")

    corpo = doc.add_paragraph(registro["relatorio_texto"])
    corpo.paragraph_format.line_spacing = 1.5

    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("________________________________________")
    doc.add_paragraph(f"Assinatura do(a) estudante: {registro.get('assinatura_estudante', '')}")
    doc.add_paragraph("")
    doc.add_paragraph("________________________________________")
    doc.add_paragraph(f"Assinatura do(a) professor(a)/pedagogo(a): {registro.get('assinatura_profissional', '')}")
    doc.add_paragraph("")
    doc.add_paragraph("________________________________________")
    doc.add_paragraph(f"Assinatura do responsável: {registro.get('assinatura_responsavel', '')}")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def draw_wrapped_text(pdf, text, x, y, max_width, font_name="Helvetica", font_size=11, leading=16):
    words = text.split()
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        if stringWidth(test, font_name, font_size) <= max_width:
            line = test
        else:
            pdf.drawString(x, y, line)
            y -= leading
            line = word
    if line:
        pdf.drawString(x, y, line)
        y -= leading
    return y


def draw_header_pdf(pdf, width, height):
    y = height - 55

    if LOGO_ESTADO_PATH.exists():
        try:
            pdf.drawImage(
                ImageReader(str(LOGO_ESTADO_PATH)),
                40,
                y - 55,
                width=105,
                height=55,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    if LOGO_ESCOLA_PATH.exists():
        try:
            pdf.drawImage(
                ImageReader(str(LOGO_ESCOLA_PATH)),
                width - 100,
                y - 62,
                width=58,
                height=58,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawCentredString(width / 2, y, "ESTADO DO PARANÁ")
    pdf.drawCentredString(width / 2, y - 14, "SECRETARIA DE ESTADO DA EDUCAÇÃO")
    pdf.drawCentredString(width / 2, y - 28, "PARANÁ INTEGRAL")
    pdf.drawCentredString(width / 2, y - 42, "ESCOLA ESTADUAL PADRE MANUEL DA NÓBREGA")
    pdf.line(40, y - 58, width - 40, y - 58)

    return y - 78


def gerar_pdf(estudante, registro):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    y = draw_header_pdf(pdf, width, height)

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawRightString(width - 45, y, "Nº ________/2026")
    y -= 24

    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawCentredString(width / 2, y, registro["tipo_relatorio"].upper())
    y -= 28

    pdf.setFont("Helvetica", 11)
    for campo in [
        f"Estudante: {estudante['nome']}",
        f"Turma: {estudante['turma']}",
        f"Responsável: {estudante['responsavel']}",
        f"Data: {registro['data']}    Hora: {registro['hora']}",
        f"Profissional: {registro['profissional']}",
        f"Disciplina: {registro['disciplina']}",
    ]:
        pdf.drawString(45, y, campo)
        y -= 18

    y -= 8
    y = draw_wrapped_text(pdf, registro["relatorio_texto"], 45, y, width - 90, leading=17)

    y -= 20
    for titulo, assinatura in [
        ("Assinatura do(a) estudante:", registro.get("assinatura_estudante", "")),
        ("Assinatura do(a) professor(a)/pedagogo(a):", registro.get("assinatura_profissional", "")),
        ("Assinatura do responsável:", registro.get("assinatura_responsavel", "")),
    ]:
        if y < 120:
            pdf.showPage()
            y = draw_header_pdf(pdf, width, height) - 20
            pdf.setFont("Helvetica", 11)

        pdf.line(45, y, 260, y)
        y -= 16
        pdf.drawString(45, y, f"{titulo} {assinatura}")
        y -= 34

    pdf.save()
    buffer.seek(0)
    return buffer


@app.route("/")
def index():
    turma = request.args.get("turma", "").strip()
    busca = request.args.get("q", "").strip().lower()

    estudantes_filtrados = estudantes

    if turma:
        estudantes_filtrados = [e for e in estudantes_filtrados if e["turma"] == turma]

    if busca:
        estudantes_filtrados = [
            e for e in estudantes_filtrados
            if busca in e["nome"].lower() or busca in e["responsavel"].lower()
        ]

    turmas = sorted({e["turma"] for e in estudantes})

    return render_template(
        "index.html",
        estudantes=estudantes_filtrados,
        turmas=turmas,
        turma_atual=turma,
        busca=busca,
        total_estudantes=len(estudantes),
        total_registros=len(registros),
    )


@app.route("/api/estudante/<int:estudante_id>")
def api_estudante(estudante_id):
    estudante = buscar_estudante(estudante_id)
    if not estudante:
        return jsonify({"erro": "Estudante não encontrado"}), 404
    return jsonify(estudante)


@app.route("/novo_registro", methods=["GET", "POST"])
def novo_registro():
    if request.method == "POST":
        estudante_id = int(request.form["estudante_id"])
        estudante = buscar_estudante(estudante_id)

        if not estudante:
            return "Estudante não encontrado.", 404

        registro = {
            "id": len(registros) + 1,
            "estudante_id": estudante_id,
            "data": request.form.get("data_registro", "").strip() or datetime.now().strftime("%d/%m/%Y"),
            "hora": request.form.get("hora_registro", "").strip() or datetime.now().strftime("%H:%M"),
            "tipo_relatorio": request.form["tipo_relatorio"],
            "disciplina": request.form["disciplina"],
            "profissional": request.form["profissional"].strip(),
            "dificuldades": request.form.getlist("dificuldades"),
            "questoes_comportamentais": request.form.getlist("questoes_comportamentais"),
            "intervencoes": request.form.getlist("intervencoes"),
            "resposta_aluno": request.form.get("resposta_aluno", ""),
            "encaminhamentos": request.form.getlist("encaminhamentos"),
            "observacao_livre": request.form.get("observacao_livre", "").strip(),
            "assinatura_estudante": request.form.get("assinatura_estudante", "").strip(),
            "assinatura_profissional": request.form.get("assinatura_profissional", "").strip(),
            "assinatura_responsavel": request.form.get("assinatura_responsavel", "").strip(),
        }

        registro["relatorio_texto"] = gerar_relatorio(estudante, registro)
        registros.append(registro)

        return redirect(url_for("relatorio_estudante", estudante_id=estudante_id))

    turmas = sorted({e["turma"] for e in estudantes})

    return render_template(
        "novo_registro.html",
        estudantes=estudantes,
        turmas=turmas,
        config=config,
        data_atual=datetime.now().strftime("%Y-%m-%d"),
        hora_atual=datetime.now().strftime("%H:%M"),
    )


@app.route("/relatorio/<int:estudante_id>")
def relatorio_estudante(estudante_id):
    estudante = buscar_estudante(estudante_id)

    if not estudante:
        return "Estudante não encontrado.", 404

    registros_estudante = [r for r in registros if r["estudante_id"] == estudante_id]

    return render_template(
        "relatorio.html",
        estudante=estudante,
        registros=registros_estudante,
    )


@app.route("/download/<int:registro_id>/<string:formato>")
def download_relatorio(registro_id, formato):
    registro = buscar_registro(registro_id)

    if not registro:
        return "Registro não encontrado.", 404

    estudante = buscar_estudante(registro["estudante_id"])

    if not estudante:
        return "Estudante não encontrado.", 404

    nome_base = nome_arquivo_base(estudante, registro)

    if formato == "txt":
        conteudo = texto_completo_exportacao(estudante, registro)
        buffer = BytesIO(conteudo.encode("utf-8"))
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{nome_base}.txt",
            mimetype="text/plain; charset=utf-8",
        )

    if formato == "docx":
        return send_file(
            gerar_docx(estudante, registro),
            as_attachment=True,
            download_name=f"{nome_base}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    if formato == "pdf":
        return send_file(
            gerar_pdf(estudante, registro),
            as_attachment=True,
            download_name=f"{nome_base}.pdf",
            mimetype="application/pdf",
        )

    return "Formato inválido.", 400


if __name__ == "__main__":
    app.run(debug=True)
