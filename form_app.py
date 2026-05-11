import importlib
import sys

import streamlit as st

from client_registry import load_client_registry, save_client_registry
from equipment_registry import load_equipment_registry, save_equipment_registry


REPORTS = {
    "FOR-BPC-5137 - RELATORIO BOMPARC - INSPECAO VISUAL EM EQUIPAMENTO": {
        "module": "reports.for_bpc_5137_visual_equipamento",
        "description": "Formulario atual do bonparcalfa5 para gerar Word e PDF a partir da mascara FOR-BPC-5137.",
    },
    "FOR-BPC-5112 - RELATORIO BOMPARC - REPROVADOS": {
        "module": "reports.for_bpc_5112_reprovados",
        "description": "Formulario para gerar o relatorio de equipamentos reprovados a partir da mascara FOR-BPC-5112.",
    },
}

REPORT_PLACEHOLDER = "Selecione uma mascara"

st.set_page_config(page_title="Gerador de Relatorios BOMPARC", layout="wide")

st.title("Gerador de Relatorios BOMPARC")
st.markdown("Selecione a mascara que deseja preencher.")

with st.expander("Cadastro de equipamentos"):
    equipment_registry = load_equipment_registry()
    if equipment_registry:
        st.dataframe(
            [
                {
                    "Equipamento": equipment,
                    "Materia Prima": config["materia_prima"],
                    "Criterio de Aceitacao": config["criterio_aceitacao"],
                }
                for equipment, config in equipment_registry.items()
            ],
            hide_index=True,
            use_container_width=True,
        )

    with st.form("equipment_registry_form"):
        equipment = st.text_input("Equipamento", "")
        raw_material = st.text_input("Materia Prima", "")
        acceptance_criterion = st.text_input("Criterio de Aceitacao", "")
        save_equipment_button = st.form_submit_button("Salvar equipamento")

    if save_equipment_button:
        equipment = equipment.strip()
        raw_material = raw_material.strip()
        acceptance_criterion = acceptance_criterion.strip()
        if not equipment or not raw_material or not acceptance_criterion:
            st.error("Preencha o equipamento, a materia-prima e o criterio de aceitacao.")
        else:
            equipment_registry[equipment] = {
                "materia_prima": raw_material,
                "criterio_aceitacao": acceptance_criterion,
            }
            save_equipment_registry(equipment_registry)
            st.success("Equipamento salvo.")
            st.rerun()

    if equipment_registry:
        delete_equipment = st.selectbox(
            "Remover equipamento cadastrado",
            ["Selecione", *equipment_registry.keys()],
        )
        if delete_equipment != "Selecione" and st.button("Remover equipamento"):
            equipment_registry.pop(delete_equipment, None)
            save_equipment_registry(equipment_registry)
            st.success("Equipamento removido.")
            st.rerun()

with st.expander("Cadastrar Cliente"):
    client_registry = load_client_registry()
    rows = []
    for client, vessels in client_registry.items():
        for vessel, address in vessels.items():
            rows.append({"Cliente": client, "Embarcacao": vessel, "Endereco": address})
    if rows:
        st.dataframe(rows, hide_index=True, use_container_width=True)

    with st.form("client_registry_form"):
        client = st.text_input("Cliente", "")
        vessel = st.text_input("Embarcacao", "")
        address = st.text_input("Endereco", "")
        save_client_button = st.form_submit_button("Salvar cliente/embarcacao")

    if save_client_button:
        client = client.strip()
        vessel = vessel.strip()
        address = address.strip()
        if not client or not vessel or not address:
            st.error("Preencha o cliente, a embarcacao e o endereco.")
        else:
            client_registry.setdefault(client, {})[vessel] = address
            save_client_registry(client_registry)
            st.success("Cliente e embarcacao salvos.")
            st.rerun()

    if client_registry:
        remove_col1, remove_col2 = st.columns(2)
        delete_client = remove_col1.selectbox(
            "Cliente para remover",
            ["Selecione", *client_registry.keys()],
        )
        delete_vessels = sorted(client_registry.get(delete_client, {})) if delete_client != "Selecione" else []
        delete_vessel = remove_col2.selectbox(
            "Embarcacao para remover",
            ["Todas as embarcacoes", *delete_vessels] if delete_vessels else ["Todas as embarcacoes"],
        )
        if delete_client != "Selecione" and st.button("Remover cliente/embarcacao"):
            if delete_vessel == "Todas as embarcacoes":
                client_registry.pop(delete_client, None)
            else:
                client_registry.get(delete_client, {}).pop(delete_vessel, None)
                if not client_registry.get(delete_client):
                    client_registry.pop(delete_client, None)
            save_client_registry(client_registry)
            st.success("Cadastro removido.")
            st.rerun()

report_name = st.selectbox(
    "Tipo de relatorio",
    [REPORT_PLACEHOLDER, *REPORTS.keys()],
    index=0,
)

if report_name == REPORT_PLACEHOLDER:
    st.info("Escolha um tipo de relatorio para abrir o formulario correspondente.")
    st.stop()

report_config = REPORTS[report_name]
st.caption(report_config["description"])
st.divider()

module_name = report_config["module"]
if module_name in sys.modules:
    importlib.reload(sys.modules[module_name])
else:
    importlib.import_module(module_name)
