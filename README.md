# 🤖 Gerador Automático de Certificados

**O Problema:** Preencher certificados manualmente um por um demora muito e gera erros.
**A Solução:** Esse robô lê uma planilha do Excel e gera centenas de certificados em Word e PDF automaticamente em segundos.

---

### 📂 O que ele faz?
1. **Lê os dados:** Pega nomes, CPFs e cursos de uma planilha Excel (vinda do MS Forms).
2. **Preenche o modelo:** Abre o arquivo Word do certificado e troca `{{NOME}}` pelo nome da pessoa.
3. **Organiza tudo:** Cria uma pasta para cada pessoa.
4. **Empacota:** Gera um arquivo `.zip` pronto para enviar por e-mail.

---

### 🚀 Como usar no seu computador

**Passo 1: Prepare as pastas**
O projeto precisa estar organizado assim:
* Pasta `entradas`: Coloque aqui o Excel com os dados (`respostas_forms.xlsx`).
* Pasta `templates`: Coloque aqui o modelo do certificado no Word (`NR06.docx`).

**Passo 2: Configure o Word**
No seu arquivo Word, onde você quer que o nome da pessoa apareça, escreva exatamente assim:
* `{{NOME}}`
* `{{CPF}}`
* `{{DATA}}`

**Passo 3: Rode o robô**
Abra o terminal e digite:
```bash
python main.py
