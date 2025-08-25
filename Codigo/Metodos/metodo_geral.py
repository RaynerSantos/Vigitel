import pandas as pd

def Automacao_Vigitel_Geral(REPLICAS_FIXO, bd_fixo, REPLICAS_CEL, bd_cel):

    bd_todas_cidades = []
    CONTAGEM_LINHAS = 4
    SOMA_LINHAS_FIXO = 0
    SOMA_LINHAS_CEL = 0
    COLUNAS = ['Réplica', 'Réplica Ativa', 'TOTAL', 'Total Elegíveis', '20', '5', '6', '4', '44', '66', '7', '8', '9', '10', '88 Eleg', 
                'Total Não Eleg', '1', '2', '3', '7 N.E', '22', 'Out. Cid. 33', '88 Não Eleg', 
                'Tx. Elegível', 'Tx. Sucesso', 'Recusa Total', 'Recusa Agenda', 'Recusa Entrev.', 'Virgens', 'Total Tentativas']

    df_total_geral = pd.DataFrame(columns=[
        'Réplica', 'TOTAL', 'Total Elegíveis', '20', '5', '6', '4', '44', '66', '7', '8', '9', '10', 
        '88 Eleg', 'Total Não Eleg', '1', '2', '3', '7 N.E', '22', 'Out. Cid. 33', '88 Não Eleg', 
        'Tx. Elegível', 'Tx. Sucesso', 'Recusa Total', 'Recusa Agenda', 'Recusa Entrev.', 'Virgens', 'Total Tentativas'])

    total_tentativas_fixo = pd.DataFrame(columns=[
    'TOTAL', 'Total Elegíveis', '20', '5', '6', '4', '44', '66', '7', '8', '9', '10', 
    '88 Eleg', 'Total Não Eleg', '1', '2', '3', '7 N.E', '22', 'Out. Cid. 33', '88 Não Eleg', 
    'Tx. Elegível', 'Tx. Sucesso', 'Recusa Total', 'Recusa Agenda', 'Recusa Entrev.', 'Virgens', 'Total Tentativas'])

    total_tentativas_cel = pd.DataFrame(columns=[
        'TOTAL', 'Total Elegíveis', '20', '5', '6', '4', '44', '66', '7', '8', '9', '10', 
        '88 Eleg', 'Total Não Eleg', '1', '2', '3', '7 N.E', '22', 'Out. Cid. 33', '88 Não Eleg', 
        'Tx. Elegível', 'Tx. Sucesso', 'Recusa Total', 'Recusa Agenda', 'Recusa Entrev.', 'Virgens', 'Total Tentativas'])

    def buscar_cidades(SOMA_LINHAS=0, SOMA_FIXA=4, REPLICAS=[], bd=pd.DataFrame()):
        lista_cidades = []
        for i, rep in enumerate(REPLICAS):
            if i == 0:
                print(bd.iloc[0, 0])
                lista_cidades.append(bd.iloc[0, 0])
            elif i == 1:
                print(bd.iloc[int(rep) + SOMA_FIXA, 0])
                lista_cidades.append(bd.iloc[int(rep) + SOMA_FIXA, 0])
                SOMA_LINHAS += int(rep) + SOMA_FIXA
            else:
                print(bd.iloc[SOMA_LINHAS + int(rep) + SOMA_FIXA, 0])
                lista_cidades.append(bd.iloc[SOMA_LINHAS + int(rep) + SOMA_FIXA, 0])
                SOMA_LINHAS += int(rep) + SOMA_FIXA
        return lista_cidades

    CIDADES = buscar_cidades(SOMA_LINHAS=0, SOMA_FIXA=4, REPLICAS=REPLICAS_FIXO, bd=bd_fixo)

    bd_fixo.columns = COLUNAS
    bd_cel.columns = COLUNAS
    for j, cidade in enumerate(CIDADES):
        #===== Selecionando o dataframe para cada cidade =====#
        if j == 0:
            # Fixo
            tabela_fixo = bd_fixo.iloc[(2 + SOMA_LINHAS_FIXO):(int(REPLICAS_FIXO[j])+2), 2:(len(COLUNAS))].reset_index()
            total_tentativas_fixo.loc[j] = bd_fixo.iloc[(int(REPLICAS_FIXO[j])+3), 2:(len(COLUNAS))].values
            SOMA_LINHAS_FIXO += int(REPLICAS_FIXO[j]) + CONTAGEM_LINHAS
            # Celular
            tabela_cel = bd_cel.iloc[(2 + SOMA_LINHAS_CEL):(int(REPLICAS_CEL[j])+2), 2:(len(COLUNAS))].reset_index()
            total_tentativas_cel.loc[j] = bd_cel.iloc[(int(REPLICAS_CEL[j])+3), 2:(len(COLUNAS))].values
            SOMA_LINHAS_CEL += int(REPLICAS_CEL[j]) + CONTAGEM_LINHAS
        
        else:
            # Fixo
            tabela_fixo = bd_fixo.iloc[(2 + SOMA_LINHAS_FIXO):(SOMA_LINHAS_FIXO + int(REPLICAS_FIXO[j])+2), 2:(len(COLUNAS))].reset_index().sort_index()
            total_tentativas_fixo.loc[j] = bd_fixo.iloc[(SOMA_LINHAS_FIXO + int(REPLICAS_FIXO[j])+3), 2:(len(COLUNAS))].values
            SOMA_LINHAS_FIXO += int(REPLICAS_FIXO[j]) + CONTAGEM_LINHAS
                
            # Celular
            tabela_cel = bd_cel.iloc[(2 + SOMA_LINHAS_CEL):(SOMA_LINHAS_CEL + int(REPLICAS_CEL[j])+2), 2:(len(COLUNAS))].reset_index().sort_index()
            total_tentativas_cel.loc[j, :] = bd_cel.iloc[(SOMA_LINHAS_CEL + int(REPLICAS_CEL[j])+3), 2:(len(COLUNAS))].values
            SOMA_LINHAS_CEL += int(REPLICAS_CEL[j]) + CONTAGEM_LINHAS


        tabela = tabela_fixo.add(tabela_cel, fill_value=0)
        tabela = tabela[COLUNAS[2:(len(COLUNAS))]]
        total_tentativas = total_tentativas_fixo.add(total_tentativas_cel, fill_value=0)


        #===== Tx. Elegível =====#
        for i in range(len(tabela["Tx. Elegível"])):
            print(f'\ntabela.loc[i, "TOTAL"]:\n{tabela.loc[i, "TOTAL"]}')
            print(f'\ntabela.loc[i, "Virgens"]:\n{tabela.loc[i, "Virgens"]}')
            denominador = ( tabela.loc[i, "TOTAL"] - tabela.loc[i, "Virgens"] )
            numerador = tabela.loc[i, "Total Elegíveis"]
            if denominador == 0:
                tabela.loc[i, "Tx. Elegível"] = 0
            else:
                tabela.loc[i, "Tx. Elegível"] = numerador / denominador

        #===== Tx. Sucesso =====#
        for i in range(len(tabela["Tx. Sucesso"])):
            denominador = tabela.loc[i, "Total Elegíveis"]
            numerador = tabela.loc[i, "20"]
            if denominador == 0:
                tabela.loc[i, "Tx. Sucesso"] = 0
            else:
                tabela.loc[i, "Tx. Sucesso"] = numerador / denominador
        
        #===== Recusa Agenda | Recusa Entrev. | Recusa Total =====#
        for i in range(len(tabela["Recusa Agenda"])):
            denominador = tabela.loc[i, "Total Elegíveis"]
            numerador_rec_agenda = tabela.loc[i, "4"]
            numerador_rec_entrev = tabela.loc[i, "44"]
            if denominador == 0:
                tabela.loc[i, "Recusa Agenda"] = 0
                tabela.loc[i, "Recusa Entrev."] = 0
            else:
                tabela.loc[i, "Recusa Agenda"] = numerador_rec_agenda / denominador
                tabela.loc[i, "Recusa Entrev."] = numerador_rec_entrev / denominador
            tabela.loc[i, "Recusa Total"] = (tabela.loc[i, "Recusa Agenda"] + tabela.loc[i, "Recusa Entrev."])
        
        #===== Tratar a última linha de sub Total e rodar todas as formulas para calcular as taxas novamente =====#
        tabela.loc[len(tabela)] = [''] * len(tabela.columns)  # Adiciona uma nova linha vazia
        tabela.iloc[-1, :] = tabela.iloc[:-1, :].sum()

        for i in range(len(tabela["Tx. Elegível"])):
            denominador = ( tabela.loc[i, "TOTAL"] - tabela.loc[i, "Virgens"] )
            numerador = tabela.loc[i, "Total Elegíveis"]
            if denominador == 0:
                tabela.loc[i, "Tx. Elegível"] = 0
            else:
                tabela.loc[i, "Tx. Elegível"] = numerador / denominador


        for i in range(len(tabela["Tx. Sucesso"])):
            denominador = tabela.loc[i, "Total Elegíveis"]
            numerador = tabela.loc[i, "20"]
            if denominador == 0:
                tabela.loc[i, "Tx. Sucesso"] = 0
            else:
                tabela.loc[i, "Tx. Sucesso"] = numerador / denominador


        for i in range(len(tabela["Recusa Agenda"])):
            denominador = tabela.loc[i, "Total Elegíveis"]
            numerador_rec_agenda = tabela.loc[i, "4"]
            numerador_rec_entrev = tabela.loc[i, "44"]
            if denominador == 0:
                tabela.loc[i, "Recusa Agenda"] = 0
                tabela.loc[i, "Recusa Entrev."] = 0
            else:
                tabela.loc[i, "Recusa Agenda"] = numerador_rec_agenda / denominador
                tabela.loc[i, "Recusa Entrev."] = numerador_rec_entrev / denominador
            tabela.loc[i, "Recusa Total"] = (tabela.loc[i, "Recusa Agenda"] + tabela.loc[i, "Recusa Entrev."])

        
        # Copiar a última linha do DataFrame original
        linha_subtotal = tabela.iloc[int(REPLICAS_CEL[j]), :].to_frame().T

        # Concatenar a linha com o DataFrame vazio
        df_total_geral = pd.concat([df_total_geral, linha_subtotal], ignore_index=True)


        replica_ativa = []
        for i in range(len(tabela["Tx. Elegível"])):
            if tabela.loc[i, "Tx. Elegível"] > 0:
                replica_ativa.append("SIM")
            else:
                replica_ativa.append("NÃO")
        tabela["Réplica Ativa"] = replica_ativa
        tabela["Réplica"] = [str(valor) for valor in list(range(1, len(tabela)+1))]
        # tabela.loc[:, 'Total Tentativas'] = ''
        tabela = tabela[COLUNAS]
        tabela.loc[len(tabela)-1, 'Réplica Ativa'] = ''
        tabela.iloc[len(tabela)-1, 0] = "Sub Total:"


        # Concatenar a tabela com a linha do total de tentativas
        # Copiar a última linha do DataFrame original
        linha_total_tentativas = total_tentativas.iloc[j, :].to_frame().T
        tabela = pd.concat([tabela, linha_total_tentativas], axis=0, ignore_index=False)
        tabela.iloc[len(tabela)-1, 0] = "Total Tentativas:"
        

        tabela.columns = pd.MultiIndex.from_product([[cidade], tabela.columns])
        
        bd_todas_cidades.append(tabela)
        
        print(bd_todas_cidades[j])
    

    #===== Última tabela =====#
    df_total_geral = df_total_geral.sum().to_frame().T
    df_total_geral['Réplica'] = 'Total geral:'
    df_total_geral.insert(1, 'Réplica Ativa', '')
    #===== Tx. Elegível =====#
    denominador = ( df_total_geral.loc[0, "TOTAL"] - df_total_geral.loc[0, "Virgens"] )
    numerador = df_total_geral.loc[0, "Total Elegíveis"]
    if denominador == 0:
        df_total_geral.loc[0, "Tx. Elegível"] = 0
    else:
        df_total_geral.loc[0, "Tx. Elegível"] = numerador / denominador

    #===== Tx. Sucesso =====#
    denominador = df_total_geral.loc[0, "Total Elegíveis"]
    numerador = df_total_geral.loc[0, "20"]
    if denominador == 0:
        df_total_geral.loc[0, "Tx. Sucesso"] = 0
    else:
        df_total_geral.loc[0, "Tx. Sucesso"] = numerador / denominador

    #===== Recusa Agenda | Recusa Entrev. | Recusa Total =====#
    denominador = df_total_geral.loc[0, "Total Elegíveis"]
    numerador_rec_agenda = df_total_geral.loc[0, "4"]
    numerador_rec_entrev = df_total_geral.loc[0, "44"]
    if denominador == 0:
        df_total_geral.loc[0, "Recusa Agenda"] = 0
        df_total_geral.loc[0, "Recusa Entrev."] = 0
    else:
        df_total_geral.loc[0, "Recusa Agenda"] = numerador_rec_agenda / denominador
        df_total_geral.loc[0, "Recusa Entrev."] = numerador_rec_entrev / denominador
    df_total_geral.loc[0, "Recusa Total"] = (df_total_geral.loc[0, "Recusa Agenda"] + df_total_geral.loc[0, "Recusa Entrev."])

    nova_coluna_df_total_geral = [''] * 30
    df_total_geral.columns = nova_coluna_df_total_geral

    bd_todas_cidades.append(df_total_geral)
    return bd_todas_cidades
        
# bd_todas_cidades = Automacao_Vigitel_Geral(REPLICAS_FIXO, bd_fixo, REPLICAS_CEL, bd_cel)