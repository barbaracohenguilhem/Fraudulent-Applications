#!/usr/bin/env python3
"""Gera a Database SQLite enriquecida (DataSet_1_ao_5.db) e a Database Excel formatada (Database_Excel_1_ao_5.xlsx)."""

import sqlite3
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


def main() -> None:
    df = pd.read_csv("DataSet_1_ao_5_limpo.csv", dtype=str).fillna("")

    # Enriquecer colunas auxiliares úteis para Database / Analytics
    df.insert(0, "id_vaga", range(1, len(df) + 1))
    df["pais"] = df["location"].apply(lambda x: x.split(",")[0].strip() if x.strip() else "Não Informado")
    df["pais"] = df["pais"].replace("", "Não Informado")
    df["fraude_bin"] = (df["fraudulent"] == "t").astype(int)
    df["status_vaga"] = df["fraudulent"].map({"t": "Fraudulenta", "f": "Real"}).fillna("Real")
    df["tem_logo"] = df["has_company_logo"].map({"t": "Sim", "f": "Não"}).fillna("Não")
    df["remoto"] = df["telecommuting"].map({"t": "Sim", "f": "Não"}).fillna("Não")
    df["tamanho_descricao"] = df["description"].str.len()

    # 1. Atualizar banco SQLite (DataSet_1_ao_5.db) com índices rápidos
    db_path = "DataSet_1_ao_5.db"
    conn = sqlite3.connect(db_path)
    df.to_sql("vagas", conn, if_exists="replace", index=False)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vagas_status ON vagas(status_vaga)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vagas_pais ON vagas(pais)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_vagas_emp ON vagas(employment_type)")
    conn.commit()
    conn.close()
    print(f"[OK] SQLite atualizado: {db_path} ({len(df)} registros na tabela 'vagas').")

    # 2. Criar Workbook Excel estruturado como Database + Resumo Executivo
    wb = Workbook()

    # --- ABA 1: Resumo_KPIs ---
    ws_kpi = wb.active
    ws_kpi.title = "Resumo_Dashboard"
    ws_kpi.views.sheetView[0].showGridLines = True

    header_fill = PatternFill("solid", fgColor="1E293B")
    sub_fill = PatternFill("solid", fgColor="334155")
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")
    accent_fill = PatternFill("solid", fgColor="FEF2F2")

    title_font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    reg_font = Font(name="Calibri", size=11, color="1E293B")

    thin = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws_kpi.merge_cells("A1:E1")
    ws_kpi["A1"] = "RESUMO DA DATABASE — VAGAS REAIS VS FRAUDULENTAS (PARTES 1 A 5)"
    ws_kpi["A1"].font = title_font
    ws_kpi["A1"].fill = header_fill
    ws_kpi["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_kpi.row_dimensions[1].height = 30

    total = len(df)
    total_fraude = int(df["fraude_bin"].sum())
    total_real = total - total_fraude
    taxa_fraude = (total_fraude / total) * 100

    sem_logo = df[df["tem_logo"] == "Não"]
    taxa_sem_logo = (sem_logo["fraude_bin"].mean() * 100) if len(sem_logo) else 0.0
    com_logo = df[df["tem_logo"] == "Sim"]
    taxa_com_logo = (com_logo["fraude_bin"].mean() * 100) if len(com_logo) else 0.0

    kpis = [
        ("Indicador Geral", "Valor"),
        ("Total de Vagas na Base (Partes 1 a 5)", total),
        ("Vagas Reais", total_real),
        ("Vagas Fraudulentas (Falsas)", total_fraude),
        ("Taxa Geral de Fraude (%)", round(taxa_fraude, 2)),
        ("Taxa de Fraude — Vagas SEM Logo (%)", round(taxa_sem_logo, 2)),
        ("Taxa de Fraude — Vagas COM Logo (%)", round(taxa_com_logo, 2)),
    ]

    for r_idx, (k, v) in enumerate(kpis, start=3):
        c1 = ws_kpi.cell(row=r_idx, column=1, value=k)
        c2 = ws_kpi.cell(row=r_idx, column=2, value=v)
        c1.border = c2.border = border
        if r_idx == 3:
            c1.fill = c2.fill = sub_fill
            c1.font = c2.font = header_font
        else:
            c1.font = bold_font
            c2.font = reg_font
            if "Fraude" in k or "Fraudulentas" in k:
                c1.fill = c2.fill = accent_fill

    # Tabela Top 10 Setores (Indústrias)
    ind_df = (
        df.assign(ind=df["industry"].replace("", "Não Informado"))
        .groupby("ind", as_index=False)
        .agg(Total=("id_vaga", "count"), Fraudulentas=("fraude_bin", "sum"))
        .sort_values("Total", ascending=False)
        .head(10)
    )
    ind_df["Taxa_Fraude_%"] = (ind_df["Fraudulentas"] / ind_df["Total"] * 100).round(2)

    start_r = 12
    headers_ind = ["Top 10 Indústrias / Setores", "Total de Vagas", "Fraudulentas", "Taxa de Fraude (%)"]
    for col_i, h_text in enumerate(headers_ind, start=1):
        cell = ws_kpi.cell(row=start_r, column=col_i, value=h_text)
        cell.fill = sub_fill
        cell.font = header_font
        cell.border = border

    for offset, row_data in enumerate(ind_df.itertuples(index=False), start=1):
        r = start_r + offset
        vals = [row_data.ind, int(row_data.Total), int(row_data.Fraudulentas), float(row_data._3)]
        for col_i, val in enumerate(vals, start=1):
            cell = ws_kpi.cell(row=r, column=col_i, value=val)
            cell.font = reg_font
            cell.border = border
            if offset % 2 == 0:
                cell.fill = zebra_fill

    # Tabela por Tipo de Contrato
    emp_df = (
        df.assign(emp=df["employment_type"].replace("", "Não Informado"))
        .groupby("emp", as_index=False)
        .agg(Total=("id_vaga", "count"), Fraudulentas=("fraude_bin", "sum"))
        .sort_values("Total", ascending=False)
    )
    emp_df["Taxa_Fraude_%"] = (emp_df["Fraudulentas"] / emp_df["Total"] * 100).round(2)

    start_r2 = 25
    headers_emp = ["Tipo de Contrato (Employment Type)", "Total de Vagas", "Fraudulentas", "Taxa de Fraude (%)"]
    for col_i, h_text in enumerate(headers_emp, start=1):
        cell = ws_kpi.cell(row=start_r2, column=col_i, value=h_text)
        cell.fill = sub_fill
        cell.font = header_font
        cell.border = border

    for offset, row_data in enumerate(emp_df.itertuples(index=False), start=1):
        r = start_r2 + offset
        vals = [row_data.emp, int(row_data.Total), int(row_data.Fraudulentas), float(row_data._3)]
        for col_i, val in enumerate(vals, start=1):
            cell = ws_kpi.cell(row=r, column=col_i, value=val)
            cell.font = reg_font
            cell.border = border
            if offset % 2 == 0:
                cell.fill = zebra_fill

    ws_kpi.column_dimensions["A"].width = 42
    ws_kpi.column_dimensions["B"].width = 20
    ws_kpi.column_dimensions["C"].width = 18
    ws_kpi.column_dimensions["D"].width = 20
    ws_kpi.column_dimensions["E"].width = 16

    # --- ABA 2: Base_Vagas_Limpa (Tabela Oficial com Filtro) ---
    ws_db = wb.create_sheet(title="Base_Vagas_Limpa")
    ws_db.freeze_panes = "A2"

    export_cols = [
        "id_vaga", "title", "status_vaga", "pais", "location", "department",
        "salary_range", "employment_type", "required_experience", "required_education",
        "industry", "function", "tem_logo", "remoto", "has_questions",
        "tamanho_descricao", "company_profile", "description", "requirements", "benefits",
        "fraudulent", "in_balanced_dataset"
    ]
    ws_db.append(export_cols)

    for row_tuple in df[export_cols].itertuples(index=False, name=None):
        ws_db.append(list(row_tuple))

    last_col_letter = get_column_letter(len(export_cols))
    last_row = len(df) + 1
    tab = Table(displayName="TabelaVagas", ref=f"A1:{last_col_letter}{last_row}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws_db.add_table(tab)

    widths = {
        "id_vaga": 10, "title": 34, "status_vaga": 15, "pais": 10, "location": 24,
        "department": 18, "salary_range": 16, "employment_type": 18,
        "required_experience": 20, "required_education": 22, "industry": 26,
        "function": 20, "tem_logo": 12, "remoto": 10, "has_questions": 14,
        "tamanho_descricao": 18, "company_profile": 45, "description": 55,
        "requirements": 45, "benefits": 40, "fraudulent": 12, "in_balanced_dataset": 18,
    }
    for idx, c_name in enumerate(export_cols, start=1):
        ws_db.column_dimensions[get_column_letter(idx)].width = widths.get(c_name, 18)

    excel_out = "Database_Excel_1_ao_5.xlsx"
    wb.save(excel_out)
    # Também atualizar DataSet_1_ao_5_limpo.xlsx com essa versão estruturada completa
    wb.save("DataSet_1_ao_5_limpo.xlsx")
    print(f"[OK] Excel Database gerada: {excel_out} e DataSet_1_ao_5_limpo.xlsx")


if __name__ == "__main__":
    main()
