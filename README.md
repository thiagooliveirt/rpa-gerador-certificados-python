# 📄 RPA - Gerador Automático de Certificados

> Automação em Python para geração em lote de certificados personalizados (Word/PDF) com compactação automática.

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-150458?style=for-the-badge&logo=pandas&logoColor=white)
![RPA](https://img.shields.io/badge/RPA-Automation-orange?style=for-the-badge)

## 🎯 Sobre o Projeto

Este projeto foi desenvolvido para resolver um problema comum em RH e Treinamentos: a **criação manual de centenas de certificados**. 

A solução lê uma base de dados (Excel/Forms), preenche um modelo Word (`.docx`) preservando toda a formatação original (estilos, fontes, logos) e gera um pacote `.zip` individual para cada colaborador, contendo seus certificados e um log de auditoria.

### 🚀 Principais Funcionalidades
* **Leitura de Dados:** Integração com planilhas Excel geradas via Microsoft Forms.
* **Manipulação de Word:** Substituição inteligente de tags (`{{NOME}}`, `{{CPF}}`) mantendo negritos e estilos.
* **Organização Automática:** Criação de pastas padronizadas (sem acentos/espaços) para cada usuário.
* **Compactação:** Geração automática de arquivos `.zip` para envio fácil.
* **Auditoria:** Geração de logs (`relatorio.json`) detalhando o status de cada arquivo gerado.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* `python-docx`: Para manipulação de documentos Word.
* `pandas` & `openpyxl`: Para leitura e tratamento de dados do Excel.
* `zipfile` & `json`: Bibliotecas nativas para gestão de arquivos e logs.

---

## ⚙️ Como Executar

### Pré-requisitos
Certifique-se de ter o Python instalado. Em seguida, instale as dependências:

```bash
pip install -r requirements.txt
