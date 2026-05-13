"""
Engine para a máscara FOR-BPC-5114 — Relatório de Inspeção por Partícula Magnética (PM).
"""
import io
import os
import docx
from reports.utils import (
    RED_VALUE,
    add_signature_to_cell,
    replace_placeholder,
    process_replacements,
    set_cell_text,
)

def build_5114(campos):
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template_path = os.path.join(app_dir, "template_5114_pm.docx")
    
    if not os.path.exists(template_path):
        template_path = os.path.join(app_dir, "template_reprovados.docx")

    item_atual = campos["itens_certificado"][0]
    certif     = f"{campos['numero_certificado']}-{item_atual}-PM"

    replacements = {
        "Relatório":             certif,
        "Número do Certificado": certif,
        "Item(s)":               item_atual,
        "Cliente":               campos["cliente"],
        "Embarcação":            campos["embarcacao"],
        "Equipamento":           campos["item"],
        "Série":                 campos["ns"],
        "Data":                  campos["data_str"],
        "Data de Inspeção 2":    campos["data_str"],
        "Critério de Aceitação": campos["criterio"],
        "Capacidade":            f"{campos['capac']:g}",
        "Unidade":               campos["unidade_capac"],
    }

    doc_out = docx.Document(template_path)
    process_replacements(doc_out, replacements)

    for table in doc_out.tables:
        for row in table.rows:
            for cell in row.cells:
                if "ASSINATURA" in cell.text.upper() or "SIGNATURE" in cell.text.upper():
                    if "QUALIDADE" in cell.text.upper():
                        add_signature_to_cell(cell, campos.get("assinatura_qualidade"))
                    else:
                        add_signature_to_cell(cell, campos.get("assinatura_tecnico"))

    return doc_out
