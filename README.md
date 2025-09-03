# Formulário -> DOCX (preenchimento automático)

Aplicativo simples (Streamlit) para preencher um **modelo .docx** a partir de um **formulário** no navegador.

## Requisitos
- Python 3.9+
- `pip install -r requirements.txt`

## Como rodar
```bash
cd form-preenche-docx
pip install -r requirements.txt
streamlit run app.py
```

Isso abrirá o app no navegador (geralmente em `http://localhost:8501`).

## Como preparar o modelo
- No Word, **destaque em amarelo** os locais que devem ser preenchidos e use nomes claros, por exemplo: `[CLIENTE]`, `[CNPJ]`, `[ENDERECO]`, `[FORO_CIDADE]`, `[DATA_DIA]`, `[DATA_MES]`, `[DATA_ANO]`.
- O app também substitui tokens sem destaque do tipo `[CHAVE]`, `{{CHAVE}}` ou `__CHAVE__`.
- No formulário, preencha cada campo. Você pode adicionar **campos extras** como JSON.

## Saída
- O app gera um **.docx preenchido** para download.

## Observações
- A normalização das chaves ignora acentos e capitaliza: use nomes em CAIXA ALTA sem acentos para evitar ambiguidade (ex.: `ENDERECO` ao invés de `ENDEREÇO`).
- O preenchimento atua em parágrafos e tabelas; preserva formatação e remove o realce amarelo dos campos preenchidos.
