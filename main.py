import os
from docx import Document

def substituir_texto(container, dicionario_dados):
    """
    Percorre parágrafos e tabelas substituindo chaves por dados reais.
    Mantém a formatação original (runs).
    """
    # 1. Procura em parágrafos simples
    for p in container.paragraphs:
        for busca, substituto in dicionario_dados.items():
            if busca in p.text:
                for run in p.runs:
                    if busca in run.text:
                        run.text = run.text.replace(busca, str(substituto))

    # 2. Procura dentro de tabelas (exigência do projeto)
    if hasattr(container, 'tables'):
        for tabela in container.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    substituir_texto(celula, dicionario_dados)

# Teste rápido para validar o ambiente
if __name__ == "__main__":
    # Dados fictícios
    dados_teste = {
        "{{NOME}}": "THIAGO OLIVEIRA",
        "{{CPF}}": "123.456.789-00",
        "{{CURSO}}": "NR06",
        "{{DATA}}": "02/02/2026"
    }
    
    template = "templates/NR06.docx"
    saida = "saidas/teste_thiago.docx"
    
    print("Iniciando teste...")
    if os.path.exists(template):
        doc = Document(template)
        substituir_texto(doc, dados_teste)
        doc.save(saida)
        print(f"Sucesso! Arquivo salvo em: {saida}")
    else:
        print(f"ERRO: Não encontrei o arquivo '{template}'. Crie-o no Word primeiro!")