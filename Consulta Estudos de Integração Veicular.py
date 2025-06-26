from flask import Flask, render_template, request, send_from_directory
import pandas as pd
import os

app = Flask(__name__)

excel_path = "Estudos de Integração Veicular.xlsx"

# Carregando os dados
df = pd.read_excel(excel_path, sheet_name="GUINDASTES", header=0)
df = df.dropna(subset=["CÓDIGO"])
df["Código"] = df["CÓDIGO"].astype(str)

# Remover colunas desnecessárias
colunas_para_remover = [
    "ALTURA", "PATOLA AUXILIAR", "PATOLA ESPECIAL", "CLIENTE", "OBS", "PDV",
    "Unnamed: 21", "Unnamed: 22", "Unnamed: 23", "TRAÇÃO.1", "VERSÕES","LANÇAS H.","LANAÇAS M.",
    "POSIÇÃO DE MONTAGEM.1", "PATOLA AUXILIAR.1", "MODELO EQUIPAMENTO.1",
    "CABINE.1", "MALHAL.1", "PLATAFORMA DE OPERAÇÃO.1", "Código"
]
df = df.drop(columns=[col for col in colunas_para_remover if col in df.columns])

# Separar colunas para tabela e filtros
columns_tabela = list(df.columns)
columns_filtros = [col for col in columns_tabela if col.upper() != "CÓDIGO"]

# Criar dicionário base de opções
options_base = {}
for col in columns_filtros:
    unique_values = df[col].dropna().astype(str).unique()
    options_base[col] = sorted(unique_values)

@app.route("/", methods=["GET", "POST"])
def index():
    filtered_df = df.copy()
    filtros_aplicados = {}

    # Captura filtros aplicados
    if request.method == "POST":
        for col in columns_filtros:
            valor = request.form.get(col, "").strip()
            if valor:
                filtros_aplicados[col] = valor
                filtered_df = filtered_df[filtered_df[col].astype(str) == valor]

    # Filtro dependente: MARCA → MODELO
    marca_selecionada = filtros_aplicados.get("MARCA", "")
    options_atualizadas = {}

    for col in columns_filtros:
        if col == "MODELO CAMINHÃO" and marca_selecionada:
            modelos_filtrados = df[df["MARCA"] == marca_selecionada][col].dropna().astype(str).unique()
            options_atualizadas[col] = sorted(modelos_filtrados)
        else:
            valores_unicos = df[col].dropna().astype(str).unique()
            options_atualizadas[col] = sorted(valores_unicos)

    # Paginação
    page = int(request.args.get("page", 1))
    per_page = 30
    start = (page - 1) * per_page
    end = start + per_page
    paged_df = filtered_df.iloc[start:end]
    has_next = len(filtered_df) > end

    # Montar resultados, tratando NaN como vazio
    resultados = []
    for _, row in paged_df.iterrows():
        result_row = {}
        for key, value in row.items():
            if pd.isna(value):
                result_row[key] = ""
            else:
                result_row[key] = value

        # Link PDF
        codigo = str(row.get("CÓDIGO", ""))
        pdf_filename = f"{codigo}.pdf"
        pdf_path = os.path.join("static", "pdfs", pdf_filename)
        result_row["PDF Link"] = f"/static/pdfs/{pdf_filename}"

        resultados.append(result_row)

    return render_template(
        "index.html",
        columns=columns_filtros,
        filtros=filtros_aplicados,
        resultados=resultados,
        options=options_atualizadas,
        page=page,
        has_next=has_next
    )

@app.route("/static/pdfs/<path:filename>")
def serve_pdf(filename):
    return send_from_directory("static/pdfs", filename)

if __name__ == "__main__":
    app.run(debug=True)
