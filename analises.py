from pathlib import Path
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OLD_FILE = DATA_DIR / "Relatório de notas fiscais e boleto antigo.csv"
NEW_FILE = DATA_DIR / "Relatório de notas fiscais e boleto novo.csv"
OUTPUT_FILE = BASE_DIR / "analise_comparativa_bases.xlsx"


def normalize_text(value):
    return "" if pd.isna(value) else str(value).strip()


def normalize_number(value):
    text = normalize_text(value).replace("R$", "").replace(" ", "")
    if not text:
        return 0.0
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def normalize_integer(value):
    text = normalize_text(value).replace(" ", "")
    if not text:
        return 0
    try:
        return int(text.replace(".", ""))
    except ValueError:
        return None


def normalize_date(value):
    date = pd.to_datetime(normalize_text(value), errors="coerce")
    return "" if pd.isna(date) else date.strftime("%Y-%m-%d")


def add_occurrence_key(frame):
    frame = frame.copy()
    frame["_cnpj"] = frame["CNPJ"].map(normalize_text)
    frame["_nota"] = frame["Nº Nota Fiscal"].map(normalize_text)
    frame["_produto"] = frame["Produto"].map(normalize_text)
    frame["_ocorrencia"] = frame.groupby(
        ["_cnpj", "_nota", "_produto"], sort=False
    ).cumcount() + 1
    frame["chave_conciliacao"] = (
        frame["_cnpj"] + "|" + frame["_nota"] + "|" + frame["_produto"]
    )
    frame["chave_linha"] = frame["chave_conciliacao"] + "|" + frame["_ocorrencia"].astype(str)
    return frame


def compare_bases(old, new):
    old = add_occurrence_key(old)
    new = add_occurrence_key(new)
    old = old.rename(columns={column: f"antigo__{column}" for column in old.columns if not column.startswith("_") and column not in {"chave_conciliacao", "chave_linha"}})
    new = new.rename(columns={column: f"novo__{column}" for column in new.columns if not column.startswith("_") and column not in {"chave_conciliacao", "chave_linha"}})
    comparison = old.merge(
        new,
        on=["_cnpj", "_nota", "_produto", "_ocorrencia", "chave_conciliacao", "chave_linha"],
        how="outer",
        indicator=True,
    )

    field_definitions = [
        ("Data de emissão", "Data Emissão", "Data emissão", normalize_date),
        ("Vencimento", "Data Vencimento", "Vencimento líquido", normalize_date),
        ("Valor original", "Valor original", "Valor Original (por produto)", normalize_number),
        ("Imposto", "Imposto retido", "Montante de imposto", normalize_number),
        ("Juros", "Juros", "Juros", normalize_number),
        ("Multa", "Multa", "Multa", normalize_number),
        ("Quantidade de transações", "Qtde Transação", "Qtde Transação", normalize_integer),
        ("Status do boleto", "Status Boleto", "Status Boleto", normalize_text),
        ("Link NFS-e", "Link  NFS-e", "Link  NFS-e", normalize_text),
        ("Link boleto", "Link  Boleto", "Link  Boleto", normalize_text),
    ]
    weights = {
        "Data de emissão": 0.10,
        "Vencimento": 0.10,
        "Valor original": 0.20,
        "Imposto": 0.05,
        "Juros": 0.05,
        "Multa": 0.05,
        "Quantidade de transações": 0.10,
        "Status do boleto": 0.15,
        "Link NFS-e": 0.10,
        "Link boleto": 0.10,
    }

    comparison["registro_encontrado_nas_duas"] = comparison["_merge"] == "both"
    comparison["score_semelhanca_%"] = 0.0
    comparison["campos_diferentes"] = ""
    for label, old_column, new_column, normalizer in field_definitions:
        old_values = comparison.get(f"antigo__{old_column}", pd.Series(index=comparison.index)).map(normalizer)
        new_values = comparison.get(f"novo__{new_column}", pd.Series(index=comparison.index)).map(normalizer)
        equal = comparison["registro_encontrado_nas_duas"] & old_values.eq(new_values)
        comparison[f"igual__{label}"] = equal
        comparison["score_semelhanca_%"] += equal.astype(float) * weights[label] * 100
        comparison.loc[~equal, "campos_diferentes"] = comparison.loc[~equal, "campos_diferentes"].apply(
            lambda current: f"{current}; {label}".lstrip("; ")
        )

    comparison["score_semelhanca_%"] = comparison["score_semelhanca_%"].round(2)
    comparison["classificacao"] = pd.cut(
        comparison["score_semelhanca_%"],
        bins=[-1, 79.99, 94.99, 100.01],
        labels=["Baixa (0-79%)", "Atenção (80-94%)", "Alta (95-100%)"],
    ).astype(str)
    return comparison, field_definitions, weights


def build_report():
    old = pd.read_csv(OLD_FILE, dtype=str).fillna("")
    new = pd.read_csv(NEW_FILE, dtype=str).fillna("")
    comparison, field_definitions, weights = compare_bases(old, new)

    client_names = old[["CNPJ", "Razão Social"]].drop_duplicates()
    client_names = client_names.rename(columns={"Razão Social": "cliente"})
    comparison = comparison.merge(client_names, left_on="_cnpj", right_on="CNPJ", how="left")
    comparison["valor_original_antigo"] = comparison["antigo__Valor original"].map(normalize_number)
    comparison["valor_original_novo"] = comparison["novo__Valor Original (por produto)"].map(normalize_number)
    comparison["quantidade_antiga"] = comparison["antigo__Qtde Transação"].map(normalize_integer)
    comparison["quantidade_nova"] = comparison["novo__Qtde Transação"].map(normalize_integer)

    total_rows = len(comparison)
    matched_rows = int(comparison["registro_encontrado_nas_duas"].sum())
    old_value_total = comparison["valor_original_antigo"].sum()
    new_value_total = comparison["valor_original_novo"].sum()
    old_transaction_total = comparison["quantidade_antiga"].sum()
    new_transaction_total = comparison["quantidade_nova"].sum()
    summary = pd.DataFrame(
        {
            "Indicador": [
                "Linhas na base antiga",
                "Linhas na base nova",
                "Linhas conciliadas pela chave",
                "Cobertura de conciliação (%)",
                "Semelhança média ponderada (%)",
                "Registros com 100% de semelhança",
                "Registros com alguma divergência",
                "Clientes analisados",
                "Valor original total - antiga",
                "Valor original total - nova",
                "Variação do valor original (%)",
                "Transações totais - antiga",
                "Transações totais - nova",
                "Variação de transações (%)",
            ],
            "Valor": [
                len(old),
                len(new),
                matched_rows,
                round(matched_rows / max(total_rows, 1) * 100, 2),
                round(comparison["score_semelhanca_%"].mean(), 2),
                int((comparison["score_semelhanca_%"] == 100).sum()),
                int((comparison["score_semelhanca_%"] < 100).sum()),
                old["CNPJ"].nunique(),
                round(old_value_total, 2),
                round(new_value_total, 2),
                round((new_value_total / old_value_total - 1) * 100, 2) if old_value_total else 0,
                int(old_transaction_total),
                int(new_transaction_total),
                round((new_transaction_total / old_transaction_total - 1) * 100, 2) if old_transaction_total else 0,
            ],
        }
    )

    client_report = comparison.groupby(["_cnpj", "cliente"], dropna=False).agg(
        linhas_antiga=("antigo__CNPJ", "count"),
        linhas_nova=("novo__CNPJ", "count"),
        registros_conciliados=("registro_encontrado_nas_duas", "sum"),
        semelhanca_media_pct=("score_semelhanca_%", "mean"),
        registros_100_pct=("score_semelhanca_%", lambda values: (values == 100).sum()),
        registros_com_divergencia=("score_semelhanca_%", lambda values: (values < 100).sum()),
        valor_original_antigo=("valor_original_antigo", "sum"),
        valor_original_novo=("valor_original_novo", "sum"),
        transacoes_antigas=("quantidade_antiga", "sum"),
        transacoes_novas=("quantidade_nova", "sum"),
    ).reset_index()
    client_report["cobertura_pct"] = (client_report["registros_conciliados"] / client_report["linhas_antiga"].clip(lower=1) * 100).round(2)
    client_report["semelhanca_media_pct"] = client_report["semelhanca_media_pct"].round(2)
    client_report["valor_original_antigo"] = client_report["valor_original_antigo"].round(2)
    client_report["valor_original_novo"] = client_report["valor_original_novo"].round(2)
    client_report["variacao_valor_pct"] = ((client_report["valor_original_novo"] / client_report["valor_original_antigo"] - 1) * 100).round(2)
    client_report["variacao_transacoes_pct"] = ((client_report["transacoes_novas"] / client_report["transacoes_antigas"] - 1) * 100).round(2)
    client_report = client_report.rename(columns={"_cnpj": "CNPJ"})

    field_report = []
    for label, _, _, _ in field_definitions:
        equal_count = int(comparison[f"igual__{label}"].sum())
        field_report.append(
            {
                "Campo analisado": label,
                "Peso no score (%)": weights[label] * 100,
                "Linhas iguais": equal_count,
                "Linhas diferentes": total_rows - equal_count,
                "Semelhança do campo (%)": round(equal_count / max(total_rows, 1) * 100, 2),
            }
        )
    field_report = pd.DataFrame(field_report)

    detail_columns = [
        "cliente", "_cnpj", "_nota", "_produto", "_ocorrencia", "chave_linha",
        "registro_encontrado_nas_duas", "score_semelhanca_%", "classificacao", "campos_diferentes",
    ]
    for label, old_column, new_column, _ in field_definitions:
        detail_columns.extend([f"antigo__{old_column}", f"novo__{new_column}", f"igual__{label}"])
    details = comparison[[column for column in detail_columns if column in comparison.columns]].copy()
    details = details.rename(columns={"_cnpj": "CNPJ", "_nota": "Nº Nota Fiscal", "_produto": "Produto", "_ocorrencia": "Ocorrência"})
    divergences = details[details["score_semelhanca_%"] < 100].copy()

    methodology = pd.DataFrame(
        {
            "Tópico": ["Chave de conciliação", "Tratamento da base nova", "Normalização", "Score", "Classificação"],
            "Descrição": [
                "CNPJ + Nº Nota Fiscal + Produto; em caso de repetição, a ocorrência sequencial também é usada.",
                "Atualizações da base nova são comparadas como estado do mesmo registro, sem penalizar linhas extras; neste recorte as duas bases têm 198 linhas.",
                "Moeda aceita R$ e separadores brasileiros; quantidade ignora pontos de milhar; datas são comparadas no formato YYYY-MM-DD; textos removem espaços nas extremidades.",
                "Média ponderada dos 10 campos comparáveis. Os pesos estão na aba Campos.",
                "Alta: 95-100%; Atenção: 80-94%; Baixa: 0-79%.",
            ],
        }
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Resumo Executivo", index=False, startrow=1)
        client_report.to_excel(writer, sheet_name="Por Cliente", index=False)
        field_report.to_excel(writer, sheet_name="Campos", index=False)
        details.to_excel(writer, sheet_name="Comparação por Registro", index=False)
        divergences.to_excel(writer, sheet_name="Divergências", index=False)
        methodology.to_excel(writer, sheet_name="Metodologia", index=False)
        format_workbook(writer)

    return OUTPUT_FILE


def format_workbook(writer):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in writer.book.worksheets:
        sheet.freeze_panes = "A2" if sheet.title != "Resumo Executivo" else "A3"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1 if sheet.title != "Resumo Executivo" else 2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for column_cells in sheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 42)
            sheet.column_dimensions[column_letter].width = max(max_length, 12)
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=sheet.title in {"Metodologia", "Divergências"})
        sheet.row_dimensions[1 if sheet.title != "Resumo Executivo" else 2].height = 30


if __name__ == "__main__":
    print(f"Relatório criado em: {build_report()}")