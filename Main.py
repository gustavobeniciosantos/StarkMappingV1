import pandas as pd
from Processamento import carregar_csv, tratar_dados, criar_tabelas_dinamicas, salvar_excel

def main():
    df = carregar_csv("Stark 09.24.csv")

    df_tratado = tratar_dados(df)

    tabela_dinamica_valores, tabela_dinamica_fees = criar_tabelas_dinamicas(df_tratado)

    salvar_excel(df_tratado, tabela_dinamica_valores, tabela_dinamica_fees, "Mapeamento Stark 09.24.xlsx")

if __name__ == "__main__":
    main()
