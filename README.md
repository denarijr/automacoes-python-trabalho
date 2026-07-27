# 🛠️ Automações em Python para Inteligência de Mercado e Processos

Este repositório reúne scripts em Python desenvolvidos para automatizar rotinas operacionais, auditoria de dados e monitoramento de mercado.

---

## 📋 Projetos Incluídos

### 1. 📊 Conciliação e Batimento de Vendas (ERP vs. Plataforma BI)
* **Arquivo:** `comparador_vendas.py`
* **Objetivo:** Automatizar a verificação diária de consistência de vendas entre o sistema ERP e a plataforma de dashboards.
* **Funcionamento:**
  * Lê e cruza os arquivos de vendas (`CSV`).
  * Trata divergências de valores e nomenclaturas de lojas.
  * Gera um relatório em Excel estilizado (`.xlsx`) com Dashboard/KPIs e aba detalhada com alertas de divergência.

---

### 2. 🏷️ Monitor de Preços de Concorrentes & Formatador de Arquivo
* **Arquivo:** `coletor_precos_concorrentes.py`
* **Objetivo:** Processar e formatar dados de pesquisa de preços de concorrentes para integração com sistemas internos.
* **Funcionamento:**
  * Realiza a leitura e estruturação dos preços e produtos monitorados.
  * Gera um arquivo de texto (`.TXT`) posicional e padronizado:
    * **Primeiros 20 caracteres:** Código EAN (preenchido com zeros à esquerda quando necessário).
    * **Últimos 5 caracteres:** Preço do produto formatado.

---

## 🚀 Tecnologias Utilizadas
* **Python 3**
* **Pandas** (Tratamento e manipulação de dados)
* **OpenPyXL** (Geração e estilização de planilhas Excel)

---

## 💡 Impacto Prático
* **Ganho de Tempo:** Substituição de tarefas manuais repetitivas por rotinas que rodam em segundos.
* **Confiabilidade:** Eliminação de erros manuais na digitação de preços e na verificação de divergências de vendas.
