# 📊 UiPath - Web Scraping & Monitoramento de Mercado (Cotações)

Este repositório contém uma solução de **Robotic Process Automation (RPA)** desenvolvida em **UiPath Studio** (linguagem VB.NET) voltada para a captura automatizada de dados financeiros na web (**Web Scraping**), tratamento de strings com **Expressões Regulares (Regex)**, validação de regras de negócio e geração de relatórios/alertas.

---

## 🎯 **Objetivo do Projeto**

Automatizar a coleta diária de cotações de moedas diretamente de portais financeiros (Google Finanças), tratar inconsistências de formatação dos valores extraídos e alertar as áreas de negócio caso os valores ultrapassem os limites operacionais estipulados.

---

## 🏗️ **Arquitetura da Solução & Módulos (.xaml)**

O projeto foi estruturado seguindo boas práticas de desenvolvimento modular, dividindo as responsabilidades em fluxos de trabalho independentes:

```
UiPath_WebScraping_Cotacoes/
├── Main.xaml                      # Orquestrador principal do projeto
├── Extracao_Dados_Web.xaml        # Abertura do navegador e extração via Table Extraction
├── Processamento_Dados.xaml       # Higienização via Regex e validação da regra de negócio
├── Envio_Relatorio_Email.xaml     # Persistência no Excel e disparo de alertas
└── Cotacoes_Diarias.xlsx          # Planilha de histórico gerada automaticamente
```

### **Detalhamento dos Módulos:**

1. **`Extracao_Dados_Web.xaml`**
   - Conecta-se ao portal web via container `Use Application/Browser`.
   - Executa a extração automatizada da tabela de cotações utilizando `Table Extraction`.
   - Retorna a `DataTable` através do argumento `out_dtCotacoes`.

2. **`Processamento_Dados.xaml`**
   - Recebe a `DataTable` (`in_dtCotacoes`).
   - Percorre os registros com `For Each Row in Data Table`.
   - Aplica Expressão Regular (`System.Text.RegularExpressions.Regex`) para higienizar o texto e isolar a parte numérica, tornando a conversão para `Double` imune a caracteres especiais como `R$`, `%`, sinais de variação (`▲`/`▼`) e símbolos regionais.
   - Avalia o valor contra o limite configurado (ex: `> 5.30`) e define a flag `out_boolAlertaDisparado`.

3. **`Envio_Relatorio_Email.xaml`**
   - Grava/acumula o histórico das cotações no arquivo Excel local (`Cotacoes_Diarias.xlsx`) através da atividade `Write Range Workbook`.
   - Caso `in_boolAlertaDisparado` seja `True`, registra e aciona a notificação de alerta por e-mail.

4. **`Main.xaml`**
   - Gerencia o ciclo de vida completo executando os fluxos via `Invoke Workflow File` e realizando a passagem de argumentos entre os módulos.

---

## 🚀 **Tecnologias e Conceitos Utilizados**

- **UiPath Studio** (Linguagem VB.NET - *Option Strict On*)
- **UI Automation Modern Experience** (`Use Application/Browser`, `Table Extraction`)
- **Manipulação Avançada de Dados** (`DataTable`, `For Each Row`, `CDbl`)
- **Expressões Regulares (Regex)** (`System.Text.RegularExpressions.Regex.Match`)
- **Excel Automation** (`Write Range Workbook`)
- **Controle de Versão** (Git & GitHub com `.gitignore` customizado para UiPath)

---

## 🔧 **Como Executar o Projeto**

1. Clone este repositório para o seu ambiente local.
2. Abra a pasta do projeto no **UiPath Studio**.
3. Certifique-se de ter o navegador **Google Chrome** (ou Edge) instalado.
4. Execute o arquivo **`Main.xaml`** (pressione `F5` ou clique em `Run`).

---

📌 *Projeto desenvolvido para fins de consolidação de conceitos de Web Scraping e RPA no UiPath.*
