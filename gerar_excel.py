import pandas as pd
import os

# Dados fictícios imitando o Microsoft Forms
dados = {
    'Nome Completo': ['Ana Silva', 'Carlos Souza', 'Beatriz Costa', 'João Pedro'],
    'CPF': ['111.222.333-44', '555.666.777-88', '999.888.777-66', '000.111.222-33'],
    'Curso': ['NR06', 'NR10', 'NR06', 'NR35'],
    'Data de Conclusão': ['10/01/2026', '12/01/2026', '15/01/2026', '20/01/2026']
}

df = pd.DataFrame(dados)

# Cria a pasta se não existir
if not os.path.exists('entradas'):
    os.makedirs('entradas')

# Salva
df.to_excel('entradas/respostas_forms.xlsx', index=False)
print("Arquivo 'entradas/respostas_forms.xlsx' criado com sucesso!")