# Python RPA SPED-Excel & Analytics IA

Este projeto facilita a conversão de arquivos SPED (.txt) para Excel (.xlsx) e vice-versa, além de incluir uma **versão experimental com análise inteligente via IA** para perguntas sobre os dados fiscais.

## Funcionalidades

- **Conversor SPED para Excel:**  
  Organiza os dados do SPED em planilhas Excel, separando cada tipo de registro em uma aba.
- **Conversor Excel para SPED:**  
  Reverte o processo, gerando o TXT oficial a partir do Excel.
- **Fórmulas de validação:**  
  Opcionalmente, insere fórmulas de validação RFB no Excel gerado.
- **Análise Inteligente (Experimental):**  
  Permite fazer perguntas em linguagem natural sobre os dados fiscais, utilizando modelos de IA (LLM).
- **Interface gráfica:**  
  Disponível em Tkinter (desktop) e Gradio (web).

## Como usar

1. **Instale os requisitos:**
   ```bash
   pip install -r requirements.txt

2. **Execute a interface desejada:**

- **Tkinter:**
   python sped_txt_to_excel.py

- **Gradio (web):**
   python Experimental_Chat-AI.py

   Acesse http://127.0.0.1:7860 no navegador.

3. **Escolha a funcionalidade:**

- **Converter TXT para Excel**
- **Converter Excel para TXT**
- **Fazer perguntas inteligentes sobre os dados (experimental)**

**Versão Experimental com IA**
- **A pasta experimental contém o módulo Experimental_Chat-AI.py, que integra análise inteligente via IA (LLM) aos dados SPED.**
⚠️ Esta funcionalidade está em desenvolvimento e pode apresentar resultados variados.

- **Exemplo de uso da análise inteligente**
   Faça upload do Excel gerado.
   Digite perguntas como:
   "Qual nota fiscal tem o maior valor?"
   "Quantos fornecedores diferentes existem?"
   "Liste os valores das notas do fornecedor X."

**Autor**
Renato - 2022-2025
