import pandas as pd
import numpy as np

def limpar_e_reconstruir_ean(valor):
    if pd.isna(valor):
        return "00000000000000000000"
    
    # Remove aspas, espaços em branco e limpa o texto
    val_str = str(valor).replace('"', '').replace("'", "").strip()
    
    # Caso ainda exista algum resíduo de notação científica
    if 'E+' in val_str or 'e+' in val_str or 'E' in val_str or 'e' in val_str:
        try:
            val_normalizado = val_str.replace(',', '.')
            num_inteiro = int(round(float(val_normalizado)))
            val_str = str(num_inteiro)
        except Exception:
            val_str = val_str.replace(',', '').replace('.', '')
            
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
        
    val_str = val_str.replace(',', '').replace('.', '')
    return val_str.zfill(20)

def formatar_preco(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return None
    
    valor_limpo = str(valor).replace('"', '').replace("'", "").strip()
    valor_limpo = valor_limpo.replace(',', '.')
    
    try:
        float_val = float(valor_limpo)
        if float_val <= 0:
            return None
            
        centavos = int(round(float_val * 100))
        return str(centavos).zfill(5)
    except ValueError:
        return None

# --- EXECUÇÃO DO SCRIPT ---

print("Lendo o arquivo base.xlsx...")
# Lendo o arquivo Excel diretamente como TEXTO (dtype=str)
df = pd.read_excel('base.xlsx', dtype=str)

coluna_ean = df.columns[0]
print(f"Coluna de EAN identificada: '{coluna_ean}'")

# Limpa e padroniza a coluna de EAN
df[coluna_ean] = df[coluna_ean].apply(limpar_e_reconstruir_ean)

# Identifica os concorrentes
concorrentes = df.columns[1:]

for concorrente in concorrentes:
    nome_arquivo = concorrente.strip().replace(' ', '_').replace('-', '_').replace('/', '_') + '.txt'
    
    df_temp = pd.DataFrame({
        'EAN': df[coluna_ean],
        'Preco_Bruto': df[concorrente]
    })
    
    df_temp['Preco_Formatado'] = df_temp['Preco_Bruto'].apply(formatar_preco)
    
    # Remove todas as linhas onde o preço ficou vazio/zerado
    df_filtrado = df_temp.dropna(subset=['Preco_Formatado'])
    
    if df_filtrado.empty:
        print(f" ! Concorrente '{concorrente}' não possui nenhum preço cadastrado. Pulando...")
        continue
        
    # Junta EAN (20 chars) + Preço (5 chars)
    linhas_combinadas = df_filtrado['EAN'] + df_filtrado['Preco_Formatado']
    
    conteudo_arquivo = "\n".join(linhas_combinadas.tolist())
    
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.write(conteudo_arquivo)
    
    print(f" -> Arquivo gerado ({len(linhas_combinadas)} produtos): {nome_arquivo}")

print("\n--- Processo finalizado com sucesso! ---")