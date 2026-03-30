from flask import Flask, render_template, request, redirect, url_for
from datetime import datetime

app = Flask(__name__)

# Banco temporário em memória para a primeira versão
estudantes = []
registros = []


@app.route("/")
def index():
    termo = request.args.get("q", "").strip().lower()
    if termo:
        filtrados = [
            e for e in estudantes
            if termo in e["nome"].lower() or termo in e["turma"].lower()
        ]
    else:
        filtrados = estudantes

    return render_template("index.html", estudantes=filtrados, termo=termo)


@app.route("/novo_estudante", methods=["GET", "POST"])
def novo_estudante():
    if request.method == "POST":
        estudante = {
            "id": len(estudantes) + 1,
            "nome": request.form["nome"].strip(),
            "turma": request.form["turma"].strip(),
            "responsavel": request.form.get("responsavel", "").strip(),
        }
        estudantes.append(estudante)
        return redirect(url_for("index"))
    return render_template("novo_estudante.html")


@app.route("/novo_registro/<int:estudante_id>", methods=["GET", "POST"])
def novo_registro(estudante_id):
    estudante = next((e for e in estudantes if e["id"] == estudante_id), None)

    if not estudante:
        return "Estudante não encontrado.", 404

    if request.method == "POST":
        registro = {
            "estudante_id": estudante_id,
            "data": datetime.now().strftime("%d/%m/%Y"),
            "professor": request.form["professor"].strip(),
            "area": request.form["area"].strip(),
            "dificuldades": request.form["dificuldades"].strip(),
            "comportamento": request.form["comportamento"].strip(),
            "intervencoes": request.form["intervencoes"].strip(),
            "encaminhamentos": request.form.get("encaminhamentos", "").strip(),
            "observacoes": request.form.get("observacoes", "").strip(),
        }
        registros.append(registro)
        return redirect(url_for("relatorio", estudante_id=estudante_id))

    return render_template("novo_registro.html", estudante=estudante)


@app.route("/relatorio/<int:estudante_id>")
def relatorio(estudante_id):
    estudante = next((e for e in estudantes if e["id"] == estudante_id), None)

    if not estudante:
        return "Estudante não encontrado.", 404

    registros_estudante = [r for r in registros if r["estudante_id"] == estudante_id]

    texto_final = []
    texto_final.append(
        f"O presente relatório pedagógico refere-se ao(à) estudante {estudante['nome']}, "
        f"da turma {estudante['turma']}, com base nos registros realizados pela equipe pedagógica e docente."
    )

    if registros_estudante:
        for r in registros_estudante:
            texto_final.append(
                f"Em {r['data']}, o(a) professor(a)/pedagogo(a) {r['professor']} registrou, na área de {r['area']}, "
                f"os seguintes apontamentos: dificuldades observadas: {r['dificuldades']}; "
                f"aspectos comportamentais: {r['comportamento']}; intervenções realizadas: {r['intervencoes']}; "
                f"encaminhamentos sugeridos: {r['encaminhamentos'] or 'não informados'}; "
                f"observações adicionais: {r['observacoes'] or 'não informadas'}."
            )

        texto_final.append(
            "Diante dos registros apresentados, recomenda-se a continuidade do acompanhamento pedagógico, "
            "com diálogo entre escola, professores, equipe pedagógica e família, buscando favorecer o "
            "desenvolvimento integral do estudante."
        )
    else:
        texto_final.append(
            "Até o momento, não há registros pedagógicos suficientes para compor um parecer mais detalhado."
        )

    relatorio_texto = "\n\n".join(texto_final)

    return render_template(
        "relatorio.html",
        estudante=estudante,
        registros=registros_estudante,
        relatorio_texto=relatorio_texto
    )


if __name__ == "__main__":
    app.run(debug=True)
