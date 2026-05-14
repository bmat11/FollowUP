
# 📦 Dashboard Inteligente de Follow-up de Compras

Sistema desenvolvido em Python + Streamlit para automatizar o follow-up de pedidos de compras, permitindo acompanhar atrasos, gerar indicadores e facilitar o contato com fornecedores.

O objetivo do projeto é reduzir o tempo gasto em follow-up manual de pedidos, melhorar o controle de atrasos e fornecer indicadores em tempo real para o setor de compras.

## 🚀 Funcionalidades

- Upload de planilhas Excel (.xlsx e .xls)
- Filtros por:
  - comprador
  - fornecedor
  - empresa
  - ordem de compra
- Busca parcial por OC
- Cálculo automático de atraso
- Dashboard de indicadores
- Ranking de fornecedores com atraso
- Geração automática de mensagens de follow-up
- Integração com Outlook Web
- Exportação da lista filtrada para Excel

## 🛠 Tecnologias utilizadas

- Python
- Streamlit
- Pandas
- OpenPyXL
- XlsxWriter
- Outlook Web Deep Link

## ⚙ Fluxo do sistema

Planilha Excel
↓
Tratamento dos dados com Pandas
↓
Cálculo de atrasos
↓
Dashboard interativo
↓
Geração automática de follow-up


## ▶ Como executar

### Clone o repositório

```bash
git clone https://github.com/bmat11/followup-compras.git


