# Dashboard de Perfil Microbiano (CCIH)

Dashboard web para **vigilância microbiológica** (Abril–Junho/2025), desenvolvido para **padronizar** e **automatizar** a leitura dos dados do laboratório e transformar planilhas em **indicadores e gráficos** para apoio à decisão (CCIH).

## 🎯 Objetivo
Unificar a visualização do **perfil microbiano trimestral**, gerando:
- **KPIs** (amostras totais, culturas positivas, taxa de positividade)
- **Distribuição por topografia** (origem das amostras)
- **Ranking de microrganismos isolados** (Top N)
- **Alerta de resistência** (MDR/XDR) e contagem de eventos críticos

## ✅ Principais funcionalidades
- Web App (Google Apps Script) com **dashboard responsivo**
- **Menu no Google Sheets**:
  - Abrir dashboard (modal)
  - Exibir link externo do Web App
  - Regravar ID da planilha (para cenários de cópia/duplicação)
- **Atualização automática** dos gráficos a partir do Sheets (`google.script.run`)
- **Fallback seguro** (`DEFAULT_PAYLOAD`) para evitar tela em branco caso o backend não esteja disponível

## 🧱 Arquitetura (alto nível)
- **Frontend**: `index.html`
  - Tailwind CSS (via CDN)
  - Chart.js (via CDN)
  - Renderização dos KPIs e gráficos (Canvas)
  - Atualização periódica (polling) a cada 10s

- **Backend**: `Code.gs`
  - `doGet()` como entrypoint do Web App
  - `getDashboardData()` para leitura, transformação e retorno do payload (JSON)
  - Leitura estruturada das abas do Sheets com nomes fixos (matrizes)

## 📊 Fontes de dados (Google Sheets)
O backend espera as abas com os nomes **exatos**:
1. `1. Produção Laboratorial`
2. `2. Distribuição por Topografia`
3. `3. Microrganismos Isolados (Frequência)`
4. `4. Perfil de Resistência (MDR/XDR)`

> Observação: o script usa `Script Properties` para guardar o ID da planilha (necessário quando o Web App roda sem “active spreadsheet”).

## 🧪 Indicadores e gráficos
- **Produção mensal**: culturas positivas, negativas e contaminadas (barras empilhadas)
- **Topografia**: comparativo por origem (barras horizontais)
- **Microrganismos**: Top 6 por frequência (barras)
- **Resistência (MDR/XDR)**: distribuição por microrganismo/mecanismo (barras)
- **KPIs**: total de amostras, positivas, positividade (%), eventos críticos

## 🚀 Como executar (Google Apps Script)
1. Crie uma planilha (Google Sheets) e configure as abas conforme os nomes acima.
2. Abra **Extensões > Apps Script** e cole:
   - `Code.gs` no arquivo principal
   - `index.html` como arquivo HTML
3. Volte ao Sheets e recarregue a página para aparecer o menu **📊 Dashboard**.
4. Teste dentro do Sheets: **📊 Dashboard > 🖥️ Abrir Dashboard**
5. Para publicar externamente:
   - **Implantar > Nova implantação > App da Web**
   - Execute como: você
   - Acesso: conforme política da unidade
6. Use **📊 Dashboard > 🔗 Ver Link Externo** para copiar a URL do Web App.

## 🔒 Boas práticas e segurança
- Evite dados sensíveis no dashboard público.
- Controle o acesso do Web App (recomendado: restrito a usuários autorizados).
- Padronize as abas e colunas para evitar inconsistência no parsing.

## 🧰 Tecnologias
- Google Apps Script (`.gs`)
- Google Sheets
- HTML
- Tailwind CSS (CDN)
- Chart.js
