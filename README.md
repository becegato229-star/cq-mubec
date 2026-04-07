# MUBEC — Sistema de Certificado de Qualidade

Sistema web para geração automática de Certificados de Qualidade em PDF.

## Como usar no dia a dia

### 1. Atualizar o ERP diariamente
1. Exporte o relatório do ERP como `.xls` ou `.xlsx`
2. Acesse o repositório no GitHub
3. Vá na pasta `data/` e clique em `erp_atual.xlsx`
4. Clique no ícone de lápis (editar) → **"Upload file"** → selecione o novo arquivo
5. Clique em **"Commit changes"**
6. Em 1-2 minutos o Railway atualiza automaticamente
7. Acesse o sistema e clique em **"Carregar ERP do Dia"**

### 2. Gerar um certificado
1. Clique em **"Carregar ERP do Dia"**
2. Selecione a nota na lista à esquerda
3. Ajuste galvanização se necessário
4. Clique em **"Gerar Certificado PDF"**

## Como instalar localmente

```bash
pip install -r requirements.txt
python app.py
# Acesse http://localhost:5050
```

## Deploy no Railway

1. Faça fork/upload deste repositório no GitHub
2. Acesse [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Selecione o repositório
4. O Railway detecta automaticamente e faz o deploy

## Estrutura do projeto

```
mubec_cq/
├── app.py                  # Servidor web Flask
├── gerar_certificado.py    # Gerador de PDF
├── requirements.txt        # Dependências Python
├── Procfile                # Comando de start (Railway)
├── nixpacks.toml           # Configuração de build (LibreOffice)
├── data/
│   ├── planilha_base.xlsx  # Base de clientes e produtos (não alterar)
│   └── erp_atual.xlsx      # ← ATUALIZAR DIARIAMENTE com exportação do ERP
└── static/
    ├── logo.png
    └── assinatura.png
```
