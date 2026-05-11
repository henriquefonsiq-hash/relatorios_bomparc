import streamlit as st
import docx
import io
import os
from datetime import date
import re
import tempfile
import zipfile
from docx.table import _Cell
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor
from PIL import Image, ImageChops, ImageOps
from client_registry import load_client_registry
from equipment_registry import load_equipment_registry

st.title("Gerador Automático de Relatório - BOMPARC")
st.markdown("Preencha o formulário abaixo para gerar o relatório de Inspeção Visual (Word e PDF).")

import platform
import subprocess

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TECHNICIAN_SIGNATURE_DIR = os.path.join(APP_DIR, "assinatura técnicos")
QUALITY_SIGNATURE_DIR = os.path.join(APP_DIR, "assinatura qualidade")
SIGNATURE_EXTENSIONS = (".jpg", ".jpeg", ".png")
SIGNATURE_AREA_FILL = 0.82
SIGNATURE_CANVAS_HEIGHT_PX = 700
QUALITY_SIGNATURE_TARGET = "media/image1.jpeg"
TECHNICIAN_SIGNATURE_TARGET = "media/image3.jpeg"
IDEAL_TABLE_GRID_WIDTHS = [7620, 1688465, 469265, 407035, 747395, 361315, 739140, 786130, 1710055]
IDEAL_EQUIPMENT_DETAIL_CELL_WIDTHS = [3398, 3551, 3931]
IDEAL_SIGNATURE_CELL_WIDTHS = [2671, 1380, 2910, 1238, 2693]
RED_VALUE = RGBColor(255, 0, 0)
VALUE_FONT_SIZE = Pt(9)
if platform.system() == "Windows":
    import pythoncom
    import win32com.client


def list_signature_files(signature_dir):
    if not os.path.isdir(signature_dir):
        return []
    return sorted(
        os.path.join(signature_dir, name)
        for name in os.listdir(signature_dir)
        if name.lower().endswith(SIGNATURE_EXTENSIONS)
    )


def signature_label(path):
    if path is None:
        return "Selecione"
    return os.path.splitext(os.path.basename(path))[0]


def paragraph_has_image_target(paragraph, target_ref):
    for blip in paragraph._element.iter(qn("a:blip")):
        rel_id = blip.get(qn("r:embed"))
        if rel_id and paragraph.part.rels[rel_id].target_ref == target_ref:
            return True
    return False


def first_descendant(element, tag):
    return next(element.iter(qn(tag)), None)


def crop_signature_whitespace(image):
    image = ImageOps.exif_transpose(image).convert("RGBA")
    white = Image.new("RGBA", image.size, (255, 255, 255, 255))
    flattened = Image.alpha_composite(white, image).convert("RGB")
    diff = ImageChops.difference(flattened, Image.new("RGB", flattened.size, "white")).convert("L")
    mask = diff.point(lambda pixel: 255 if pixel > 18 else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return image

    left, top, right, bottom = bbox
    padding = 6
    left = max(left - padding, 0)
    top = max(top - padding, 0)
    right = min(right + padding, image.width)
    bottom = min(bottom + padding, image.height)
    return image.crop((left, top, right, bottom))


def make_signature_fit_canvas(signature_path, target_cx, target_cy):
    if not target_cx or not target_cy:
        return signature_path

    with Image.open(signature_path) as source:
        signature = crop_signature_whitespace(source)

    aspect_ratio = target_cx / target_cy
    canvas_height = SIGNATURE_CANVAS_HEIGHT_PX
    canvas_width = max(1, int(round(canvas_height * aspect_ratio)))
    max_width = max(1, int(canvas_width * SIGNATURE_AREA_FILL))
    max_height = max(1, int(canvas_height * SIGNATURE_AREA_FILL))

    scale = min(max_width / signature.width, max_height / signature.height)
    new_size = (
        max(1, int(round(signature.width * scale))),
        max(1, int(round(signature.height * scale))),
    )
    signature = signature.resize(new_size, Image.Resampling.LANCZOS)

    canvas = Image.new("RGBA", (canvas_width, canvas_height), (255, 255, 255, 0))
    x = int((canvas_width - signature.width) / 2)
    y = int((canvas_height - signature.height) / 2)
    canvas.alpha_composite(signature, (x, y))

    fd, fitted_path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    canvas.save(fitted_path)
    return fitted_path


def keep_cell_text_on_one_line(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    no_wrap = tc_pr.find(qn("w:noWrap"))
    if no_wrap is None:
        tc_pr.append(OxmlElement("w:noWrap"))
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_cell_width(cell, width):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.tcW
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width))
    tc_w.set(qn("w:type"), "dxa")


def apply_ideal_table_dimensions(doc):
    if not doc.tables:
        return

    table = doc.tables[0]
    grid_columns = table._tbl.tblGrid.gridCol_lst
    for grid_column, width in zip(grid_columns, IDEAL_TABLE_GRID_WIDTHS):
        grid_column.w = width

    if len(table.rows) <= 17:
        return

    for row_index in (6, 7):
        real_cells = list(table.rows[row_index]._tr.tc_lst)
        for tc, width in zip(real_cells, IDEAL_EQUIPMENT_DETAIL_CELL_WIDTHS):
            set_cell_width(_Cell(tc, table), width)

    for row_index in (16, 17):
        real_cells = list(table.rows[row_index]._tr.tc_lst)
        for tc, width in zip(real_cells, IDEAL_SIGNATURE_CELL_WIDTHS):
            set_cell_width(_Cell(tc, table), width)


def remove_empty_trailing_paragraphs(doc):
    body = doc._element.body

    for child in reversed(list(body)):
        if child.tag == qn("w:sectPr"):
            continue

        if child.tag != qn("w:p"):
            break

        has_text = any((text.text or "").strip() for text in child.iter(qn("w:t")))
        has_drawing = any(child.iter(qn("w:drawing")))
        has_page_break = any(
            break_element.get(qn("w:type")) == "page"
            for break_element in child.iter(qn("w:br"))
        )

        if has_text or has_drawing or has_page_break:
            break

        body.remove(child)


def replace_signature(doc, signature_path, target_ref):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if not paragraph_has_image_target(paragraph, target_ref):
                        continue

                    for drawing in paragraph._element.iter(qn("w:drawing")):
                        for blip in drawing.iter(qn("a:blip")):
                            rel_id = blip.get(qn("r:embed"))
                            if not rel_id or paragraph.part.rels[rel_id].target_ref != target_ref:
                                continue

                            extent = first_descendant(drawing, "wp:extent")
                            target_cx = int(extent.get("cx")) if extent is not None else None
                            target_cy = int(extent.get("cy")) if extent is not None else None
                            fitted_path = make_signature_fit_canvas(signature_path, target_cx, target_cy)
                            try:
                                new_rel_id, _ = paragraph.part.get_or_add_image(fitted_path)
                            finally:
                                if fitted_path != signature_path and os.path.exists(fitted_path):
                                    os.remove(fitted_path)

                            blip.set(qn("r:embed"), new_rel_id)
                            anchor = first_descendant(drawing, "wp:anchor")
                            if anchor is not None:
                                anchor.set("behindDoc", "1")
                            return True
    return False


def validate_template(doc):
    if not doc.tables:
        return "O template.docx não possui tabela principal."

    table = doc.tables[0]
    if len(table.rows) <= 17:
        return "A tabela principal do template.docx precisa ter pelo menos 18 linhas."

    for row_index in (16, 17):
        if len(table.rows[row_index].cells) <= 7:
            return "As linhas de assinatura do template.docx precisam ter pelo menos 8 colunas."

    image_targets = {
        rel.target_ref
        for rel in doc.part.rels.values()
        if rel.target_ref.startswith("media/image")
    }
    missing_targets = [
        target
        for target in (QUALITY_SIGNATURE_TARGET, TECHNICIAN_SIGNATURE_TARGET)
        if target not in image_targets
    ]
    if missing_targets:
        return "O template.docx não possui as imagens-base de assinatura esperadas: " + ", ".join(missing_targets)

    return None


def format_conversion_error(error):
    if platform.system() == "Windows":
        return (
            "Não foi possível converter o relatório para PDF. "
            "Verifique se o Microsoft Word está instalado e fechado, e tente novamente. "
            f"Detalhe técnico: {error}"
        )

    return (
        "Não foi possível converter o relatório para PDF. "
        "Verifique se o LibreOffice está instalado no ambiente. "
        f"Detalhe técnico: {error}"
    )


def parse_certificate_items(raw_items):
    items = []
    seen = set()
    for raw_item in re.split(r"[,;\n]+", raw_items):
        item_value = re.sub(r"(?i)^item\s+", "", raw_item.strip()).strip("- ")
        if not item_value or item_value in seen:
            continue
        items.append(item_value)
        seen.add(item_value)
    return items


def parse_serial_numbers(raw_serial_numbers):
    return [
        serial_number.strip()
        for serial_number in re.split(r"[,;\n]+", raw_serial_numbers)
        if serial_number.strip()
    ]


def safe_filename_part(value):
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "-", str(value)).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(". ") or "relatorio"


def word_value(value):
    return str(value).upper()

# Função para converter DOCX para PDF (Nuvem ou Local)
def convert_to_pdf(docx_path, pdf_path):
    if platform.system() == "Windows":
        # O Streamlit roda em threads, o Windows exige CoInitialize() para comunicação COM
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = PDF
        finally:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()
            pythoncom.CoUninitialize()
    else:
        # Nuvem (Linux) - Utiliza o LibreOffice via linha de comando
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf",
            os.path.abspath(docx_path),
            "--outdir", os.path.dirname(os.path.abspath(pdf_path))
        ], check=True)

    if not os.path.exists(pdf_path):
        raise RuntimeError("o arquivo PDF não foi criado.")


equipment_registry = load_equipment_registry()
equipment_options = sorted(equipment_registry)
client_registry = load_client_registry()
client_options = sorted(client_registry)


# --- FORMULÁRIO ---
with st.container():
    st.subheader("1. Identificação e Local")
    col1, col2, col3 = st.columns(3)
    cliente = col1.selectbox(
        "Cliente",
        ["Selecione", *client_options, "Outro"],
        key="cliente_item",
    )
    if cliente == "Selecione":
        cliente = ""
    if cliente == "Outro":
        cliente = col1.text_input("Especifique o Cliente", "", key="cliente_outro", placeholder="DOF SUBSEA BRASIL SERVIÇOS LTDA")

    vessel_options = sorted(client_registry.get(cliente, {}))
    embarcacao = col2.selectbox(
        "Embarcação (Local do Teste)",
        ["Selecione", *vessel_options, "Outro"],
        key="embarcacao_item",
    )
    if embarcacao == "Selecione":
        embarcacao = ""
    if embarcacao == "Outro":
        embarcacao = col2.text_input("Especifique a Embarcação", "", key="embarcacao_outro", placeholder="SKANDI CHIEFTAIN")

    if st.session_state.get("ultima_embarcacao_cliente") != (cliente, embarcacao):
        st.session_state["endereco_cliente"] = client_registry.get(cliente, {}).get(embarcacao, "")
        st.session_state["ultima_embarcacao_cliente"] = (cliente, embarcacao)
    endereco = col3.text_input("Endereço", key="endereco_cliente", placeholder="PORTO DO AÇU")

    col4, col5 = st.columns(2)
    numero_certificado = col4.text_input("Número do certificado", "", placeholder="68430204")
    item_certificado = col5.text_input("Item(s)", "", placeholder="32")
    
    st.subheader("2. Dados da Inspeção e Equipamento")
    col7, col8, col9 = st.columns(3)
    data_inspecao = col7.date_input("Data da Inspeção", value=None, format="DD/MM/YYYY")
    ns = col8.text_area(
        "NS (Número de Série)",
        "",
        placeholder="Uma NS para todos ou uma por item: AST220S/03, AST220S/04",
        height=80,
    )
    item = col9.selectbox(
        "Equipamento",
        ["Selecione", *equipment_options, "Outro"],
        key="equipamento_item",
    )
    if item == "Selecione":
        item = ""
    if item == "Outro":
        item = st.text_input("Especifique o Equipamento", "", key="equipamento_outro", placeholder="Manilha")

    if st.session_state.get("ultimo_equipamento_item") != item:
        equipment_config = equipment_registry.get(item, {})
        st.session_state["criterio_aceitacao"] = equipment_config.get("criterio_aceitacao", "")
        st.session_state["materia_prima"] = equipment_config.get("materia_prima", "")
        st.session_state["ultimo_equipamento_item"] = item
    criterio = st.text_input("Critério de Aceitação", key="criterio_aceitacao", placeholder="ABNT NBR 15637-1")
        
    col10, col11, col12, col13 = st.columns(4)
    capac = col10.number_input("Capacidade (Capac)", min_value=0.1, value=None, step=0.1, placeholder="3.00")
    unidade_capacidade = col11.text_input("Unidade da capacidade", "", placeholder="TON")
    dimensao = col12.text_input("Dimensão", "", placeholder="2 MTS")
    quantidade = col13.number_input("Quantidade", min_value=1, value=None, step=1, placeholder="1")
    
    col14, col15 = st.columns(2)
    materia_prima = col14.text_input("Matéria Prima", key="materia_prima", placeholder="POLIESTER")
    teste = col15.text_input("Teste", "", placeholder="N/A")
    
    st.subheader("3. Parâmetros e Parecer")
    col16, col17, col18 = st.columns(3)
    end = col16.text_input("END (Ensaio Não Destrutivo)", "", placeholder="N/A")
    aprov = col17.selectbox("APROV (Laudo Final)", ["Selecione", "APROVADO", "REPROVADO"])
    if aprov == "Selecione":
        aprov = ""
    obs = col18.text_area("OBS (Observações)", "", placeholder="NÃO HOUVE / NONE")

    quality_signature_files = list_signature_files(QUALITY_SIGNATURE_DIR)
    assinatura_qualidade = st.selectbox(
        "Assinatura do Controle de Qualidade",
        [None, *quality_signature_files],
        format_func=signature_label,
        index=0,
    ) if quality_signature_files else None
    if not quality_signature_files:
        st.warning(f"Nenhuma assinatura encontrada na pasta {QUALITY_SIGNATURE_DIR}.")

    signature_files = list_signature_files(TECHNICIAN_SIGNATURE_DIR)
    assinatura_tecnico = st.selectbox(
        "Assinatura do Técnico",
        [None, *signature_files],
        format_func=signature_label,
        index=0,
    ) if signature_files else None
    if not signature_files:
        st.warning(f"Nenhuma assinatura encontrada na pasta {TECHNICIAN_SIGNATURE_DIR}.")

    submit_button = st.button("Gerar Relatórios")

if submit_button:
    with st.spinner("Processando documento..."):
        numero_certificado = numero_certificado.strip().strip("-")
        itens_certificado = parse_certificate_items(item_certificado)
        numeros_serie = parse_serial_numbers(ns)
        if not numero_certificado or not itens_certificado:
            st.error("Preencha o Número do certificado e pelo menos um Item para gerar o relatório.")
            st.stop()
        if not numeros_serie:
            st.error("Preencha pelo menos um NS (Número de Série) para gerar o relatório.")
            st.stop()
        if len(numeros_serie) not in (1, len(itens_certificado)):
            st.error(
                "A quantidade de NS deve ser 1 para repetir em todos os itens "
                "ou igual à quantidade de Item(s)."
            )
            st.stop()
        unidade_capacidade = unidade_capacidade.strip().upper()
        if not unidade_capacidade:
            st.error("Preencha a unidade da capacidade para gerar o relatório.")
            st.stop()
        criterio = criterio.strip()
        if not criterio:
            st.error("Preencha o Critério de Aceitação para gerar o relatório.")
            st.stop()
        if data_inspecao is None:
            st.error("Preencha a Data da Inspeção para gerar o relatório.")
            st.stop()
        if not item:
            st.error("Preencha o Equipamento para gerar o relatório.")
            st.stop()
        if capac is None:
            st.error("Preencha a Capacidade para gerar o relatório.")
            st.stop()
        if quantidade is None:
            st.error("Preencha a Quantidade para gerar o relatório.")
            st.stop()

        # Regras Inteligentes
        # 1. Data + 1 ano
        data_str = data_inspecao.strftime("%d/%m/%Y")
        data_arquivo_str = data_inspecao.strftime("%d-%m-%Y")
        try:
            data_validade_str = data_inspecao.replace(year=data_inspecao.year + 1).strftime("%d/%m/%Y")
        except ValueError: # Tratamento para ano bissexto (29/02)
            data_validade_str = data_inspecao.replace(year=data_inspecao.year + 1, day=28).strftime("%d/%m/%Y")
            
        # 2. Equipamento = [Item] + [Capac] + [Unidade] + " X " + [Dimensao]
        equipamento = f"{item.upper()} {capac:g} {unidade_capacidade} X {dimensao.upper()}"
        
        # 3. Carga de trabalho: TON vira KG; outras unidades permanecem como preenchidas
        if unidade_capacidade == "TON":
            carga_trabalho = f"{int(capac * 1000):,} KG".replace(",", ".")
        else:
            carga_trabalho = f"{capac:g} {unidade_capacidade}"
        
        # 4. OS: usa os 4 primeiros dígitos do número do certificado.
        ordem_servico = f"BPC-OSTC-{numero_certificado[:4]}"

        red_replacement_labels = {
            "SÉRIE (Serial Number):",
            "DATA DA INSPEÇÃO (Date):",
            "DATA DE VALIDADE (Validity Date):",
        }

        following_paragraph_replacements = {
            "DIMENSÕES (Dimensiones):": dimensao,
            "MATÉRIA PRIMA (Feedstock):": materia_prima,
            "QUANTIDADE": f"{quantidade:02d} UNIDADE",
        }

        # Carregar o DOCX
        template_path = os.path.join(APP_DIR, "template.docx")
        
        if not os.path.exists(template_path):
            st.error(f"Erro: Modelo 'template.docx' não encontrado na pasta {APP_DIR}.")
            st.stop()

        def set_paragraph_text_preserving_first_run(paragraph, value):
            first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            first_run.text = word_value(value)
            set_value_run_style(first_run)
            for run in paragraph.runs[1:]:
                run.text = ""

        def set_value_run_style(run, color=None):
            run.font.name = "Lato"
            run.font.size = VALUE_FONT_SIZE
            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.rFonts
            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                r_pr.append(r_fonts)
            for font_attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                r_fonts.set(qn(font_attr), "Lato")
            if color is not None:
                run.font.color.rgb = color

        def replace_value_after_label(paragraph, label, value, color=None):
            full_text = paragraph.text
            if label not in full_text:
                return False

            value_start = full_text.index(label) + len(label)
            value_start += len(full_text[value_start:]) - len(full_text[value_start:].lstrip())
            value_end = len(full_text)

            if value_start >= value_end:
                if paragraph.runs:
                    paragraph.runs[-1].text = paragraph.runs[-1].text.rstrip() + " "
                    value_run = paragraph.add_run(word_value(value))
                else:
                    value_run = paragraph.add_run(word_value(value))
                set_value_run_style(value_run, color)
                return True

            cursor = 0
            inserted = False
            for run in paragraph.runs:
                original_text = run.text
                run_start = cursor
                run_end = cursor + len(original_text)
                cursor = run_end

                if run_end <= value_start:
                    continue
                if run_start >= value_end:
                    run.text = ""
                    continue

                if not inserted:
                    keep_before = original_text[:max(value_start - run_start, 0)]
                    keep_after = original_text[max(value_end - run_start, 0):]
                    run.text = keep_before + word_value(value) + keep_after
                    set_value_run_style(run, color)
                    inserted = True
                else:
                    keep_after = original_text[max(value_end - run_start, 0):]
                    run.text = keep_after

            return True

        def replace_in_paragraphs(paragraphs, paragraph_replacements):
            for i, p in enumerate(paragraphs):
                p_text = p.text

                for label, value in following_paragraph_replacements.items():
                    if label in p_text and i + 1 < len(paragraphs):
                        next_paragraph = paragraphs[i + 1]
                        if ":" in next_paragraph.text:
                            replace_value_after_label(next_paragraph, next_paragraph.text.split(":", 1)[0] + ":", value)
                        else:
                            set_paragraph_text_preserving_first_run(next_paragraph, value)
                        break

                for label, value in paragraph_replacements.items():
                    color = RED_VALUE if label in red_replacement_labels else None
                    if replace_value_after_label(p, label, value, color):
                        break

                if re.fullmatch(r"\s*\d{2}/\d{2}/\d{4}\s*", p_text):
                    set_paragraph_text_preserving_first_run(p, data_str)

        with tempfile.TemporaryDirectory() as temp_dir:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, item_certificado_atual in enumerate(itens_certificado):
                    certif = f"{numero_certificado}-{item_certificado_atual}"
                    numero_serie_atual = numeros_serie[0] if len(numeros_serie) == 1 else numeros_serie[index]
                    nome_base_arquivo = " ".join(
                        safe_filename_part(part)
                        for part in (certif, item, data_arquivo_str)
                    )

                    paragraph_replacements = {
                        "RELATÓRIO Nº (Number Report):": certif,
                        "CLIENTE (Client):": cliente,
                        "LOCAL DO TESTE (Local of the Test):": embarcacao,
                        "ENDEREÇO (Address):": endereco,
                        "EQUIPAMENTO (Equipment):": equipamento,
                        "SÉRIE (Serial Number):": numero_serie_atual,
                        "DATA DA INSPEÇÃO (Date):": data_str,
                        "DATA DE VALIDADE (Validity Date):": data_validade_str,
                        "CRITÉRIO DE ACEITAÇÃO (Criteria of accept):": criterio,
                        "CARGA DE TRABALHO \n(Workload):": carga_trabalho,
                        "IDENTIFICAÇÃO DE PLAQUETA (ID Plate):": certif,
                        "ORDEM DE SERVIÇO (Service Order):": ordem_servico,
                        "RELATÓRIO DE ENSAIO NÃO DESTRUTIVO (No-Destrutive Reheasal report):": end,
                        "LAUDO FINAL (Final Report):": aprov,
                        "Observações e Recomendações (Observation and Recommendations):": obs,
                    }

                    try:
                        doc = docx.Document(template_path)
                    except Exception as error:
                        st.error(f"Erro ao abrir o template.docx: {error}")
                        st.stop()

                    template_error = validate_template(doc)
                    if template_error:
                        st.error(template_error)
                        st.stop()

                    if assinatura_qualidade and not replace_signature(
                        doc,
                        assinatura_qualidade,
                        QUALITY_SIGNATURE_TARGET,
                    ):
                        st.error("Não foi possível encontrar a assinatura fixa do controle de qualidade no template.docx.")
                        st.stop()

                    if assinatura_tecnico and not replace_signature(
                        doc,
                        assinatura_tecnico,
                        TECHNICIAN_SIGNATURE_TARGET,
                    ):
                        st.error("Não foi possível encontrar a assinatura fixa do técnico no template.docx.")
                        st.stop()

                    replace_in_paragraphs(doc.paragraphs, paragraph_replacements)

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                replace_in_paragraphs(cell.paragraphs, paragraph_replacements)

                    signature_table = doc.tables[0]
                    for cell_index in (2, 7):
                        keep_cell_text_on_one_line(signature_table.rows[16].cells[cell_index])
                        keep_cell_text_on_one_line(signature_table.rows[17].cells[cell_index])
                    apply_ideal_table_dimensions(doc)
                    remove_empty_trailing_paragraphs(doc)

                    output_docx = os.path.join(temp_dir, f"{nome_base_arquivo}.docx")
                    output_pdf = os.path.join(temp_dir, f"{nome_base_arquivo}.pdf")

                    doc.save(output_docx)
                    try:
                        convert_to_pdf(output_docx, output_pdf)
                    except Exception as error:
                        st.error(format_conversion_error(error))
                        st.stop()

                    zip_file.write(output_docx, os.path.basename(output_docx))
                    zip_file.write(output_pdf, os.path.basename(output_pdf))
            zip_buffer.seek(0)

        st.success(f"{len(itens_certificado)} relatório(s) gerado(s) com sucesso!")

        st.download_button(
            label="Baixar Word e PDF",
            data=zip_buffer.getvalue(),
            file_name=f"Relatorios_{numero_certificado}.zip",
            mime="application/zip",
        )
