# Gerador de Relatorios BOMPARC

Aplicativo Streamlit para preencher mascaras BOMPARC e gerar relatorios em Word e PDF.

## Deploy no Streamlit Community Cloud

1. Suba esta pasta para um repositorio GitHub.
2. No Streamlit Cloud, crie um novo app apontando para:
   - Branch: `main`
   - Main file path: `form_app.py`
3. Confirme que estes arquivos estao no repositorio:
   - `requirements.txt`
   - `packages.txt`
   - `template.docx`
   - `template_reprovados.docx`
   - `assinatura qualidade/`
   - `assinatura tecnicos/` ou `assinatura técnicos/`
   - `reports/`

## Dependencias

As dependencias Python ficam em `requirements.txt`.

O arquivo `packages.txt` instala o LibreOffice no ambiente Linux do deploy. Ele e necessario para converter os arquivos `.docx` em `.pdf`.

## Observacoes importantes

- `GeradorBOMPARC.exe`, `Gerador BOMPARC.bat`, logs e `__pycache__` sao artefatos locais e ficam fora do deploy pelo `.gitignore`.
- Os cadastros salvos nos arquivos JSON funcionam no ambiente online, mas em plataformas sem disco persistente eles podem ser perdidos em reinicios/redeploys. Para persistencia definitiva, use banco externo ou armazenamento conectado.
- Em Linux, a conversao para PDF depende do comando `libreoffice` instalado via `packages.txt`.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run form_app.py
```
