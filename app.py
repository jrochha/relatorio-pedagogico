from flask import Flask, render_template, request, redirect, url_for, jsonify
from datetime import datetime
import json
from pathlib import Path

app = Flask(__name__, template_folder="templates", static_folder="static")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

with open(DATA_DIR / "alunos.json", "r", encoding="utf-8") as f:
    estudantes = json.load(f)

with open(DATA_DIR / "config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

registros = []


def buscar_estudante(estudante_id):
    return next((e for e in estudantes if e["id"] == estudante_id), None)


def formatar_lista(lista):
    if not lista:
        return ""
    if len(lista) == 1:
        return lista[0].lower()
    if len(lista) == 2:
        return f"{lista[0].lower()} e {lista[1].lower()}"
    return ", ".join(item.lower() for item in lista[:-1]) + f" e {lista[-1].lower()}"


def gerar_relatorio(estudante, registro):
    disciplina = registro["disciplina"]
    dificuldades = formatar_lista(registro["dificuldades"])
    comportamentos = formatar_lista(registro["questoes_comportamentais"])
    intervencoes = formatar_lista(registro["intervencoes"])
    encaminhamentos = formatar_lista(registro["encaminhamentos"])
    resposta = registro["resposta_aluno"].lower() if registro["resposta_aluno"] else "não informada"
    observacao = registro["observacao_livre"].strip()

    partes = [
        f"O presente relatório pedagógico refere-se ao(à) estudante {estudante['nome']}, da turma {estudante['turma']}, "
        f"com base no registro realizado em {registro['data']} na disciplina de {disciplina}."
    ]

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
        partes.append(
            f"Como encaminhamento, recomenda-se {encaminhamentos}."
        )

    if observacao:
        partes.append(f"Observação complementar: {observacao}")

    return " ".join(partes)


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
            "data": datetime.now().strftime("%d/%m/%Y"),
            "disciplina": request.form["disciplina"],
            "profissional": request.form["profissional"].strip(),
            "dificuldades": request.form.getlist("dificuldades"),
            "questoes_comportamentais": request.form.getlist("questoes_comportamentais"),
            "intervencoes": request.form.getlist("intervencoes"),
            "resposta_aluno": request.form.get("resposta_aluno", ""),
            "encaminhamentos": request.form.getlist("encaminhamentos"),
            "observacao_livre": request.form.get("observacao_livre", "").strip(),
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


if __name__ == "__main__":
    app.run(debug=True)
