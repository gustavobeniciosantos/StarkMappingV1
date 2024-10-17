import pandas as pd

def carregar_csv(arquivo_csv):
    df = pd.read_csv(arquivo_csv, header=None, dtype={0: str, 1: str, 2: str, 3: str, 4: str, 5: str})
    
    df.iloc[1:, 5] = pd.to_numeric(df.iloc[1:, 5], errors='coerce')
    df.iloc[1:, 0] = pd.to_numeric(df.iloc[1:, 0], errors='coerce')
    df = df.drop(0)

    return df

def tratar_dados(df):

    df.insert(0, 'Chave de conciliação', "")
    df.insert(1, 'Data tratada', "")

    
    df = df.rename(columns={0: 'amount', 1: 'balance', 2: 'created', 3: 'description', 4: 'externalId', 5: 'fee', 6: 'id', 7: 'receiverId', 8: 'senderId', 9: 'source', 10: 'tags'})

    df['Data tratada'] = pd.to_datetime(df['created'], errors='coerce').dt.strftime('%d/%m/%Y')

    df.loc[df['description'].str.contains('Devolução:', case=False, na=False), 'Chave de conciliação'] = 'Devolução'
    df.loc[df['description'].str.contains('Estorno:', case=False, na=False), 'Chave de conciliação'] = 'Estorno'
    df.loc[df['description'].str.contains('KIWIFY EDUCACAO E TECNOLOGIA LTDA (36.149.947/0001-06)', case=False, na=False, regex=False), 'Chave de conciliação'] = 'Transf Banco Itaú'
    df.loc[(df['source'].str.contains('invoice/', case=False, na=False)) & (df['Chave de conciliação'] == ''), 'Chave de conciliação'] = 'Venda com Pix'
    df.loc[(df['source'].str.contains('boleto-issuing', case=False, na=False)) & (df['Chave de conciliação'] == ''), 'Chave de conciliação'] = 'Custo Boleto'
    df.loc[(df['source'].str.contains('invoice-issuing', case=False, na=False)) & (df['Chave de conciliação'] == ''), 'Chave de conciliação'] = 'Custo Pix'
    df.loc[(df['source'].str.contains('pix-table', case=False, na=False)) & (df['Chave de conciliação'] == ''), 'Chave de conciliação'] = 'Custo Ajuste Pix'
    df.loc[(df['source'].str.contains('boleto/', case=False, na=False)), 'Chave de conciliação'] = 'Venda com Boleto'
    df.loc[(df['source'].str.contains('deposit/')) & (~df['description'].str.contains('KIWIFY EDUCACAO E TECNOLOGIA LTDA (36.149.947/0001-06)', case=False, na=False, regex=False)), 'Chave de conciliação'] = 'Recebimento TED Comprador'
    df.loc[(df['tags'].str.contains('withdraw-request,withdraw', case=False, na=False)), 'Chave de conciliação'] = 'Saque do Produtor'
    df.loc[(df['tags'].str.contains('payment-request/', case=False, na=False)) & (df['Chave de conciliação'] == ''), 'Chave de conciliação'] = 'Trasnferência Pix Manual'

    return df

def criar_tabelas_dinamicas(df):

    tabela_dinamica_valores = pd.pivot_table(df, values='amount',
                                             index='Data tratada',
                                             columns='Chave de conciliação',
                                             aggfunc='sum',
                                             fill_value=0)
    
    tabela_dinamica_fees = pd.pivot_table(df, values='fee',
                                          index='Data tratada',
                                          columns='Chave de conciliação',
                                          aggfunc='sum',
                                          fill_value=0)
    return tabela_dinamica_valores, tabela_dinamica_fees

def salvar_excel(df, tabela_valores, tabela_fees, nome_arquivo):

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados Tratados')
        tabela_valores.to_excel(writer, sheet_name='Dinâmica Valor')
        tabela_fees.to_excel(writer, sheet_name='Dinâmica Taxas')

    print('Tabelas dinâmicas adicionadas ao arquivo Excel.')
