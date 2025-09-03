
import streamlit as st
from io import BytesIO
from fill_docx import fill_document_bytes

st.set_page_config(page_title="Preencher DOCX", page_icon="📄")
st.title("Preencher modelo NDA a partir de formulário")

st.markdown("""
Preencha os campos abaixo, envie o **modelo .docx** (com textos destacados em amarelo ou tokens como `[CLIENTE]`, `{{CNPJ}}`) e clique em **Gerar documento**.
- Recomendo marcar cada local de preenchimento com um **placeholder em amarelo** (ex.: `[CLIENTE]`, `[CNPJ]`, `[ENDERECO]`, `[FORO_CIDADE]`).
- Campos comuns (edite conforme seu contrato): **CLIENTE, CNPJ, ENDERECO, CONTRATO_TIPO, CONTRATO_OBJETIVO, FORO_CIDADE, DATA_DIA, DATA_MES, DATA_ANO**.
""")

with st.form("form-docx"):
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("CLIENTE", placeholder="ACME TECNOLOGIA S.A.")
        cnpj = st.text_input("CNPJ", placeholder="12.345.678/0001-90")
        endereco = st.text_input("ENDERECO", placeholder="Av. Brasil, 1000, Centro, Vitória/ES, CEP 29000-000")
        contrato_tipo = st.text_input("CONTRATO_TIPO", placeholder="Prestação de Serviços de Software")
        foro = st.text_input("FORO_CIDADE", placeholder="Vitória")
    with col2:
        contrato_objetivo = st.text_area("CONTRATO_OBJETIVO", placeholder="implementar e dar suporte a soluções de integração de dados", height=120)
        data_dia = st.text_input("DATA_DIA", placeholder="11")
        data_mes = st.text_input("DATA_MES", placeholder="agosto")
        data_ano = st.text_input("DATA_ANO", placeholder="2025")
        extra_json = st.text_area("Campos extras (JSON opcional)", placeholder='{"RESPONSAVEL": "Fulano", "EMAIL": "contato@empresa.com"}', height=120)

    template = st.file_uploader("Modelo .docx", type=["docx"])

    submitted = st.form_submit_button("Gerar documento")

if submitted:
    if not template:
        st.error("Envie um arquivo .docx de modelo.")
    else:
        # Monta o dicionário de dados
        data = {
            "CLIENTE": cliente,
            "CNPJ": cnpj,
            "ENDERECO": endereco,
            "CONTRATO_TIPO": contrato_tipo,
            "CONTRATO_OBJETIVO": contrato_objetivo,
            "FORO_CIDADE": foro,
            "DATA_DIA": data_dia,
            "DATA_MES": data_mes,
            "DATA_ANO": data_ano,
        }
        # Mesclar extras JSON, se válido
        import json
        if extra_json.strip():
            try:
                extra = json.loads(extra_json)
                if isinstance(extra, dict):
                    data.update(extra)
                else:
                    st.warning("Os 'Campos extras' devem ser um JSON de objeto (ex.: {\"CHAVE\": \"valor\"}). Ignorando.")
            except Exception as e:
                st.warning(f"JSON inválido em 'Campos extras'. Ignorando. Erro: {e}")

        bytes_in = template.read()
        try:
            output_bytes = fill_document_bytes(bytes_in, data)
            st.success("Documento gerado com sucesso!")
            st.download_button(
                label="Baixar .docx preenchido",
                data=output_bytes,
                file_name="documento_preenchido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Falha ao gerar o documento: {e}")
            st.stop()

st.markdown("---")
st.caption("Dica: No modelo, pinte de **amarelo** cada placeholder (ex.: `[CLIENTE]`) para o app reconhecer facilmente. Também funciona com tokens entre colchetes/chaves sem destaque.")
