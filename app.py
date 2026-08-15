from flask import Flask, request, render_template_string, Response
from openpyxl import load_workbook, Workbook
import os
import io

app = Flask(__name__)

base_path = os.path.dirname(__file__)
arquivos_cds = {
    "HUB": os.path.join(base_path, "cd1.xlsx"),
    "DUTRA": os.path.join(base_path, "CD2.xlsx"),
    "JACANA": os.path.join(base_path, "CD3.xlsx"),
    "LESTE": os.path.join(base_path, "CD4.xlsx"),
    "JABAQUARA": os.path.join(base_path, "CD5.xlsx"),
    "CENTRO": os.path.join(base_path, "CD6.xlsx")
}

def carregar_dados(origem=None, codigo=None, nome=None, qtd_min=None, qtd_max=None, limite=3000):
    dados = []
    if origem and origem in arquivos_cds:
        caminhos = {origem: arquivos_cds[origem]}
    else:
        caminhos = arquivos_cds

    for origem, caminho in caminhos.items():
        if os.path.exists(caminho):
            wb = load_workbook(caminho, read_only=True, data_only=True)
            sheet = wb.active
            for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True)):
                if i >= limite:
                    break

                codigo_val = row[0]   # Coluna A
                nome_val = row[1]     # Coluna B
                qtd_h = row[7]        # Coluna H

                # só aceita linhas com informação nas 3 colunas
                if not (codigo_val and nome_val and qtd_h is not None):
                    continue

                # aplica filtros já na leitura
                if codigo and str(codigo_val) != codigo:
                    continue
                if nome and nome.lower() not in str(nome_val).lower():
                    continue
                if qtd_min and (qtd_h < int(qtd_min)):
                    continue
                if qtd_max and (qtd_h > int(qtd_max)):
                    continue

                dados.append({
                    "codigo": str(codigo_val),
                    "nome": nome_val,
                    "disponivel": qtd_h,
                    "origem": origem
                })
            wb.close()
    return dados

@app.route("/")
def index():
    codigo = request.args.get("codigo")
    nome = request.args.get("nome")
    qtd_min = request.args.get("qtd_min")
    qtd_max = request.args.get("qtd_max")
    origem = request.args.get("origem")

    dados = carregar_dados(origem, codigo, nome, qtd_min, qtd_max)

    html = """
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Consulta Estoques</title>
        <style>
            body { font-family: Arial, sans-serif; background: linear-gradient(135deg, #1f4037, #99f2c8); margin: 0; padding: 20px; color: #333; }
            h1 { text-align: center; color: #fff; margin-bottom: 20px; }
            form { background: #fff; padding: 15px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; }
            form input, form select { padding: 8px; border: 1px solid #ccc; border-radius: 6px; }
            form input[type="submit"], form button { background: #1f4037; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; transition: 0.3s; }
            form input[type="submit"]:hover, form button:hover { background: #0d2c22; }
            table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
            th { background: #1f4037; color: #fff; padding: 12px; text-align: center; }
            td { padding: 10px; text-align: center; border-bottom: 1px solid #ddd; }
            tr:hover { background: #f2f2f2; }
        </style>
    </head>
    <body>
        <h1>CONTROLE DE ESTOQUES</h1>
        <form method="get">
            Código: <input type="text" name="codigo" value="{{ request.args.get('codigo','') }}">
            Nome: <input type="text" name="nome" value="{{ request.args.get('nome','') }}">
            Quantidade mínima: <input type="number" name="qtd_min" value="{{ request.args.get('qtd_min','') }}">
            Quantidade máxima: <input type="number" name="qtd_max" value="{{ request.args.get('qtd_max','') }}">
            Origem:
            <select name="origem">
                <option value=""></option>
                {% for origem in ['HUB','DUTRA','JACANA','LESTE','JABAQUARA','CENTRO'] %}
                <option value="{{ origem }}" {% if request.args.get('origem')==origem %}selected{% endif %}>{{ origem }}</option>
                {% endfor %}
            </select>
            <input type="submit" value="Filtrar">
            <button type="submit" formaction="/exportar">Exportar Excel</button>
        </form>
        <table>
            <tr><th>Código</th><th>Nome</th><th>Disponível (H)</th><th>Origem</th></tr>
            {% for d in dados %}
            <tr>
                <td>{{ d.codigo }}</td>
                <td>{{ d.nome }}</td>
                <td>{{ d.disponivel }}</td>
                <td>{{ d.origem }}</td>
            </tr>
            {% endfor %}
        </table>
    </body>
    </html>
    """
    return render_template_string(html, dados=dados)

@app.route("/exportar")
def exportar():
    codigo = request.args.get("codigo")
    nome = request.args.get("nome")
    qtd_min = request.args.get("qtd_min")
    qtd_max = request.args.get("qtd_max")
    origem = request.args.get("origem")

    dados = carregar_dados(origem, codigo, nome, qtd_min, qtd_max)

    wb = Workbook()
    ws = wb.active
    ws.title = "Estoques Filtrados"
    ws.append(["Código", "Nome", "Disponível (H)", "Origem"])
    for d in dados:
        ws.append([d["codigo"], d["nome"], d["disponivel"], d["origem"]])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return Response(output,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment;filename=estoques_filtrados.xlsx"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
