import os
import pandas as pd
import numpy as np

def limpar_e_reconstruir_ean(valor):
    """
    Higieniza o código EAN removendo aspas, símbolos e notação científica.
    Preenche com zeros à esquerda até atingir exatamente 20 caracteres.
    """
    if pd.isna(valor):
        return "00000000000000000000"
    
    # Remove aspas, espaços em branco e limpa o texto
    val_str = str(valor).replace('"', '').replace("'", "").strip()
    
    # Tratamento para notação científica derivada de exportações no Excel
    if any(e in val_str for e in ['E+', 'e+', 'E', 'e']):
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
    """
    Converte o valor monetário para o formato de centavos inteiros.
    Retorna uma string preenchida com zeros à esquerda em exatamente 5 caracteres (ex: R$ 15,90 -> 01590).
    """
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

def processar_pesquisa_precos():
    arquivo_entrada = 'dados_pesquisa_precos.xlsx'
    
    if not os.path.exists(arquivo_entrada):
        print(f"Erro: O arquivo '{arquivo_entrada}' não foi encontrado na pasta do script.")
        return

    print(f"Lendo a base de dados: '{arquivo_entrada}'...")
    # Lendo o arquivo Excel como string para evitar perdas de zeros no EAN
    df = pd.read_excel(arquivo_entrada, dtype=str)

    coluna_ean = df.columns[0]
    print(f"Coluna de EAN identificada: '{coluna_ean}'")

    # Limpa e padroniza a coluna de EAN
    df[coluna_ean] = df[coluna_ean].apply(limpar_e_reconstruir_ean)

    # Identifica as colunas de concorrentes/estabelecimentos (a partir da 2ª coluna)
    estabelecimentos = df.columns[1:]

    for local in estabelecimentos:
        # Sanitiza o nome do arquivo de saída .txt
        nome_arquivo = "export_precos_" + local.strip().replace(' ', '_').replace('-', '_').replace('/', '_') + '.txt'
        
        df_temp = pd.DataFrame({
            'EAN': df[coluna_ean],
            'Preco_Bruto': df[local]
        })
        
        df_temp['Preco_Formatado'] = df_temp['Preco_Bruto'].apply(formatar_preco)
        
        # Remove registros sem preço válido
        df_filtrado = df_temp.dropna(subset=['Preco_Formatado'])
        
        if df_filtrado.empty:
            print(f" ! Estabelecimento '{local}' não possui preços válidos cadastrados. Pulando...")
            continue
            
        # Concatena EAN (20 posições) + Preço em centavos (5 posições)
        linhas_combinadas = df_filtrado['EAN'] + df_filtrado['Preco_Formatado']
        
        conteudo_arquivo = "\n".join(linhas_combinadas.tolist())
        
        with open(nome_arquivo, 'w', encoding='utf-8') as f:
            f.write(conteudo_arquivo)
        
        print(f" -> Arquivo posicional gerado ({len(linhas_combinadas)} produtos): {nome_arquivo}")

    print("\n--- Processo de exportação finalizado com sucesso! ---")

if __name__ == "__main__":
    processar_pesquisa_precos()
