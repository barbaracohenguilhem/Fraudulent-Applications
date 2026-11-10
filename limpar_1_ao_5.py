#!/usr/bin/env python3
"""Limpeza rápida e completa dos arquivos DataSet_parte_01.csv a DataSet_parte_05.csv.

Gera:
- DataSet_parte_01.csv .. DataSet_parte_05.csv limpos (com backup em backup_originais/)
- DataSet_1_ao_5_limpo.csv (unificado pronto para Python / Dash / Pandas)
- DataSet_1_ao_5_limpo.xlsx (planilha nativa pronta para abrir direto no Excel)
- DataSet_1_ao_5.db (banco SQLite pronto caso use SQL no Dash)
"""

import csv
import os
import re
import shutil
import sqlite3
import zipfile
from xml.sax.saxutils import escape

csv.field_size_limit(10**7)

PARTES = [f"DataSet_parte_0{i}.csv" for i in range(1, 6)]
BACKUP_DIR = "backup_originais"

# Windows-1252 C1 bytes que apareceram mal decodificados como caracteres de controle U+0080..U+009F
CP1252_FIXES = str.maketrans({
    "\x91": "'",
    "\x92": "'",
    "\x93": '"',
    "\x94": '"',
    "\x96": "-",
    "\x97": "-",
    "\x9f": "• ",
    "\xa0": " ",
    "\ufeff": "",
})

# Regexes pré-compiladas para máxima velocidade
RE_URL_HASH = re.compile(r"(?:https?://[^\s#]*|www\.[^\s#]*)?#URL_[0-9a-fA-F]+#*")
RE_EMAIL_HASH = re.compile(r"(?:[^\s#@]+@)?#EMAIL_[0-9a-fA-F]+#*")
RE_PHONE_HASH = re.compile(r"#PHONE_[0-9a-fA-F]+#*")
RE_HTML_TAG = re.compile(r"</?(?:b|i|u|p|br|div|span|ul|ol|li|strong|em|h[1-6]|a|font|table|tr|td)\b[^>]*>", re.I)
RE_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
RE_WHITESPACE = re.compile(r"\s+")


def clean_text(val: str) -> str:
    if not val:
        return ""
    # 1. Cortar vazamento de memória/CSS corrompido do dataset bruto (marcado por \ufffd após quebra de parágrafo)
    if "\ufffd" in val:
        idx = val.find("\ufffd")
        sep = val.rfind("\n", 0, idx)
        if sep != -1:
            val = val[:sep]
        else:
            val = val.replace("\ufffd", " ")

    # 2. Corrigir caracteres Windows-1252 mal codificados e remover controles invisíveis
    val = val.translate(CP1252_FIXES)
    val = RE_CONTROL.sub(" ", val)

    # 3. Substituir hashes gigantes (#URL_...#, #EMAIL_...#, #PHONE_...#) por marcadores limpos
    val = RE_URL_HASH.sub("[URL]", val)
    val = RE_EMAIL_HASH.sub("[EMAIL]", val)
    val = RE_PHONE_HASH.sub("[TELEFONE]", val)

    # 4. Remover eventuais tags HTML residuais
    val = RE_HTML_TAG.sub(" ", val)

    # 5. Normalizar quebras de linha e espaços múltiplos (1 linha limpa por registro para Excel e Dash)
    val = RE_WHITESPACE.sub(" ", val).strip()
    return val


def col_letter(idx: int) -> str:
    """Converte índice 0-based em letra de coluna do Excel (A, B, ..., R)."""
    res = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res = chr(65 + rem) + res
    return res


def write_xlsx(filepath: str, headers: list[str], rows: list[list[str]]) -> None:
    """Gera um arquivo .xlsx nativo usando apenas biblioteca padrão do Python."""
    shared_strings: list[str] = []
    str_index: dict[str, int] = {}

    def get_si(s: str) -> int:
        idx = str_index.get(s)
        if idx is None:
            idx = len(shared_strings)
            str_index[s] = idx
            shared_strings.append(s)
        return idx

    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>',
    ]

    all_rows = [headers] + rows
    col_letters = [col_letter(c) for c in range(len(headers))]

    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, val in enumerate(row):
            if val == "" or val is None:
                continue
            # Excel tem limite de 32.767 caracteres por célula
            if len(val) > 32760:
                val = val[:32760]
            si = get_si(val)
            cells.append(f'<c r="{col_letters[c_idx]}{r_idx}" t="s"><v>{si}</v></c>')
        sheet_lines.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    sheet_lines.append("</sheetData></worksheet>")

    sst_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">',
    ]
    for s in shared_strings:
        sst_lines.append(f"<si><t xml:space=\"preserve\">{escape(s)}</t></si>")
    sst_lines.append("</sst>")

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    wb_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="DataSet_1_ao_5" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(filepath, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", wb_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml", "".join(sheet_lines))
        zf.writestr("xl/sharedStrings.xml", "".join(sst_lines))


def main() -> None:
    os.makedirs(BACKUP_DIR, exist_ok=True)
    unified_rows: list[list[str]] = []
    clean_headers: list[str] = []

    for fn in PARTES:
        backup_path = os.path.join(BACKUP_DIR, fn)
        if not os.path.exists(backup_path):
            shutil.copy2(fn, backup_path)

        # Ler do backup original para garantir idempotência
        with open(backup_path, "r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            raw_headers = next(reader)
            headers = [h.replace("\ufeff", "").strip() for h in raw_headers]
            if not clean_headers:
                clean_headers = headers

            part_rows: list[list[str]] = []
            for row in reader:
                cleaned = [clean_text(cell) for cell in row]
                part_rows.append(cleaned)
                unified_rows.append(cleaned)

        # Sobrescrever a parte limpa (UTF-8 sem BOM sujo no título)
        with open(fn, "w", encoding="utf-8", newline="") as out_f:
            writer = csv.writer(out_f)
            writer.writerow(headers)
            writer.writerows(part_rows)

        print(f"[OK] {fn}: {len(part_rows)} linhas limpas.")

    # Salvar CSV unificado (partes 1 a 5) para Python / Dash
    unified_csv = "DataSet_1_ao_5_limpo.csv"
    with open(unified_csv, "w", encoding="utf-8", newline="") as out_f:
        writer = csv.writer(out_f)
        writer.writerow(clean_headers)
        writer.writerows(unified_rows)
    print(f"[OK] {unified_csv}: {len(unified_rows)} linhas totais.")

    # Salvar Excel (.xlsx) nativo pronto para abrir direto no Excel
    xlsx_path = "DataSet_1_ao_5_limpo.xlsx"
    write_xlsx(xlsx_path, clean_headers, unified_rows)
    print(f"[OK] {xlsx_path} gerado com sucesso.")

    # Salvar SQLite (.db) pronto para Dash/SQL se necessário
    db_path = "DataSet_1_ao_5.db"
    if os.path.exists(db_path):
        os.remove(db_path)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cols_ddl = ", ".join(f'"{h}" TEXT' for h in clean_headers)
    cur.execute(f"CREATE TABLE vagas ({cols_ddl})")
    placeholders = ", ".join("?" for _ in clean_headers)
    cur.executemany(f"INSERT INTO vagas VALUES ({placeholders})", unified_rows)
    conn.commit()
    conn.close()
    print(f"[OK] {db_path} (tabela 'vagas') gerado com sucesso.")


if __name__ == "__main__":
    main()

