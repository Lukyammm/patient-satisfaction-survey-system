# 📊 Pesquisa de Satisfação — HUC

Sistema web para coleta, gestão e análise de pesquisas de satisfação de pacientes do Hospital Universitário do Ceará (HUC).

---

## 📝 Descrição objetiva

O **PESQUISA_SATISFACAO** é um WebApp desenvolvido em **Google Apps Script + Google Sheets + HTML/CSS/JS** que permite registrar, visualizar e exportar pesquisas de satisfação aplicadas às clínicas e setores hospitalares.

O sistema calcula automaticamente indicadores de **NPS (Net Promoter Score)**, **taxa de satisfação por setor** e **tabulação de manifestações** (sugestões, reclamações, comentários e elogios), com filtros dinâmicos e painel executivo em tempo real.

---

## 🚀 Funcionalidades

- Dashboard com KPIs: total respondido, taxa de satisfação, NPS e manifestações
- Filtros por setor, tipo, sexo e período de data
- Gráficos de satisfação por setor e classificação NPS (Promotores / Neutros / Detratores)
- Formulário para nova pesquisa com avaliações por grupo: Acolhimento, Assistência e Serviços
- Tabela de registros com busca e edição
- Relatório por clínica/setor com tabulação geral
- Exportação em CSV

---

## 🛠️ Tecnologias

- Google Apps Script (backend)
- Google Sheets (banco de dados — aba `MATRIZ`)
- HTML / CSS / JavaScript (frontend WebApp)

---

## ⚙️ Como publicar

1. Crie uma planilha Google com a aba `MATRIZ` (linhas a partir da linha 5).
2. Em `Extensões > Apps Script`, cole os arquivos `Code.gs` e `index.html`.
3. Implante em `Implantar > Nova implantação` como **Aplicativo da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Sua organização** (ou conforme necessidade)
4. Copie o URL gerado e compartilhe com a equipe.
