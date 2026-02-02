import os
import pandas as pd
from docx import Document
import unicodedata
import re

def substituir_texto(container, dicionario_dados):
    """Substitui texto em parágrafos e tabelas mantendo formatação."""
    # 1. Parágrafos
    for p in container.paragraphs:
        for busca, substituto in dicionario_dados.items():
            if busca in p.text:
                for run in p.runs:
                    if busca in run.text:
                        run.text = run.text.replace(busca, str(substituto))
    
    # 2. Tabelas
    if hasattr(container, 'tables'):
        for tabela in container.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    substituir_texto(celula, dicionario_dados)

def limpar_nome_pasta(nome):
    """
    Transforma 'João da Silva' em 'joao_da_silva' 
    (Requisito do Workana: minúsculas, sem acento, underscore)
    """
    # Normaliza para remover acentos
    nfkd = unicodedata.normalize('NFKD', nome)
    nome_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    
    # Remove caracteres especiais e troca espaço por _
    nome_limpo = re.sub(r'[^a-zA-Z0-9 ]', '', nome_sem_acento)
    return nome_limpo.lower().replace(' ', '_')

def processar_lote():
    arquivo_excel = "entradas/respostas_forms.xlsx"
    
    if not os.path.exists(arquivo_excel):
        print("ERRO: Excel não encontrado. Rode o script gerar_excel.py primeiro.")
        return

    print("Lendo planilha...")
    df = pd.read_excel(arquivo_excel)

    for index, linha in df.iterrows():
        # Dados da linha atual
        nome = linha['Nome Completo']
        cpf = linha['CPF']
        curso = linha['Curso']
        data = linha['Data de Conclusão']
        
        print(f"Processando: {nome} - {curso}...")

        # Mapeamento (De -> Para)
        dados_certificado = {
            "{{NOME}}": nome,
            "{{CPF}}": cpf,
            "{{CURSO}}": curso,
            "{{DATA}}": data
        }

        # Definição dos caminhos
        nome_pasta = limpar_nome_pasta(nome)
        caminho_pasta_saida = os.path.join("saidas", nome_pasta)
        
        # Cria pasta do colaborador
        if not os.path.exists(caminho_pasta_saida):
            os.makedirs(caminho_pasta_saida)
        
        # Carrega o template (Assume que o nome do curso é igual ao nome do arquivo .docx)
        caminho_template = f"templates/{curso}.docx"
        
        # Verifica se o template existe (Ex: se temos NR35.docx)
        if not os.path.exists(caminho_template):
            # Fallback: Se não achar o NR específico, usa o NR06 como genérico para teste
            print(f"   Aviso: Template {curso} não encontrado. Usando NR06.docx como base.")
            caminho_template = "templates/NR06.docx"

        if os.path.exists(caminho_template):
            doc = Document(caminho_template)
            substituir_texto(doc, dados_certificado)
            
            nome_arquivo = f"{nome_pasta}_{curso}.docx"
            doc.save(os.path.join(caminho_pasta_saida, nome_arquivo))
        else:
            print(f"   ERRO CRÍTICO: Nem o template original nem o NR06 existem.")

if __name__ == "__main__":
    processar_lote()