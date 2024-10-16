import os
import pandas as pd

# Verificar se o arquivo existe
def verificar_arquivo(caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo não encontrado: {caminho_arquivo}")
        return False
    return True

# Função para carregar arquivo com detecção automática do cabeçalho
def carregar_arquivo(caminho_arquivo, sheet):
    for skip in [2, 3, 4, 5, 6]:  # Testa diferentes linhas de cabeçalho
        try:
            # Tenta carregar o arquivo com diferentes valores de 'skiprows'
            df = pd.read_excel(caminho_arquivo, skiprows=skip, sheet_name=sheet)
            
            # Verifica se o DataFrame contém colunas essenciais
            if df.columns.size > 0:  # Verifica se há colunas no DataFrame
                return df
        except Exception as e:
            continue  # Se falhar, tenta o próximo valor de skiprows

    raise ValueError(f"Não foi possível carregar o arquivo corretamente: {caminho_arquivo}")

# Salvar cabeçalhos
def salvar_cabecalho(arquivo_entrada, sheet, arquivo_saida):
    if verificar_arquivo(arquivo_entrada):
        # Carregar o arquivo usando detecção automática de cabeçalho
        df = carregar_arquivo(arquivo_entrada, sheet)
        colunas = df.columns
        df_cabecalho = pd.DataFrame(columns=colunas)
        df_cabecalho.to_excel(arquivo_saida, index=False)
        print(f"Cabeçalhos salvos em: {arquivo_saida}")
        return colunas
    return None

def main():
    # Mercado Livre
    arquivo_entrada_ml = './Entrada/ml.xlsx'
    arquivo_saida_ml = './Saida/ml_saida.xlsx'

    colunas_ml = salvar_cabecalho(arquivo_entrada_ml, sheet='Anúncios', arquivo_saida=arquivo_saida_ml)

    if colunas_ml is not None:
        print("Colunas do Mercado Livre:")
        print(colunas_ml)

    # Shopee
    arquivo_entrada_shopee = './Entrada/shopee.xlsx'
    arquivo_saida_shopee = './Saida/shopee_saida.xlsx'

    colunas_shopee = salvar_cabecalho(arquivo_entrada_shopee, sheet='Sheet1', arquivo_saida=arquivo_saida_shopee)

    if colunas_shopee is not None:
        print("Colunas do Shopee:")
        print(colunas_shopee)

if __name__ == "__main__":
    main()
