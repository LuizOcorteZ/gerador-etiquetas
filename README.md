# 🪲 Gerador Automático de Etiquetas Entomológicas (Padrão CEMT)

Uma aplicação web gratuita desenvolvida em Python/Streamlit para automatizar a formatação e geração de etiquetas entomológicas no padrão CEMT (Coleção Entomológica da Universidade Federal de Mato Grosso, baseado nas orientações do Prof. Fernando Vaz-de-Mello).

🔗 **Acesse a aplicação online:** [https://etiqueta-automatica.streamlit.app/](https://etiqueta-automatica.streamlit.app/)

---

## 📋 Pré-requisitos para a Planilha de Entrada

Para garantir que o algoritmo processe e formate os dados corretamente, siga estas orientações ao preparar a sua planilha:

### 1. Formato de Ficheiro Suportado
* O aplicativo aceita ficheiros do **Microsoft Excel (`.xlsx`)** e ficheiros de texto delimitado **(`.csv`)**.
* **Nota sobre o CSV:** O sistema detecta automaticamente o separador de colunas (vírgula ou ponto e vírgula) e lida com codificações padrão (UTF-8) e latino-americanas nativamente (Latin1).

### 2. Estrutura do Cabeçalho
* A planilha deve possuir **apenas uma linha de cabeçalho** (nomes das colunas).
* **Não utilize células mescladas** nas linhas de dados ou no cabeçalho.
* Indique no menu do aplicativo em qual linha está localizado o seu cabeçalho (padrão: Linha 1).

### 3. Formato das Datas
* As datas devem estar num formato reconhecível (ex: `DD/MM/AAAA` ou `AAAA-MM-DD`).
* O sistema converte automaticamente o mês para **algarismos romanos** (Exemplo: `15/03/2024` → `15.iii.2024`).

### 4. Formato das Coordenadas Geográficas
* O recurso de limpeza de milésimos de segundo busca coordenadas geográficas no formato sexagesimal com indicação explícita de segundos marcada por aspas simples ou duplas (ex: `18°30'12.34"S` ou `18°30'12'S`).

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Streamlit** (Interface Web e motor de Cache)
* **Pandas** (Manipulação e filtragem de dados estruturados)
* **python-docx** (Geração e formatação de documentos Word)
* **Google Apps Script API** (Armazenamento de feedbacks em tempo real)

---

## 💬 Feedback e Contribuições

Este projeto está em contínuo desenvolvimento para agilizar as rotinas de laboratório. Se encontrou algum erro, tem dúvidas ou sugestões de melhoria para atender outros grupos taxonômicos, sinta-se à vontade para enviar a sua mensagem pela **aba de Feedback no menu lateral do aplicativo** ou abrir uma *Issue* neste repositório.