import pandas as pd

# 1 - Carregar arquivo xlxs
try:

    df = pd.read_excel("camisas_umhb.xlsx")
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nome', 'Tamanho','Total pago', 'Falta'])
# Garantir que as colunas 'Falta' e 'Total pago' sejam do tipo numérico (int ou float)
df['Falta'] = pd.to_numeric(df['Falta'], errors='coerce').fillna(0)
df['Total pago'] = pd.to_numeric(df['Total pago'], errors='coerce').fillna(0)

# Função para atualizar os totais utilizados no relatório
def atualizar_totais(dataframe):
    total_pessoas = len(dataframe)
    total_quitados = len(dataframe[dataframe['Falta'] == 0])
    total_nao_pagaram = len(dataframe[dataframe['Total pago'] == 0])
    total_parcial = len(dataframe[(dataframe['Total pago'] > 0) & (dataframe['Falta'] > 0)])
    total_caixa = dataframe['Total pago'].sum()
    total_a_receber = dataframe['Falta'].sum()
    return total_pessoas, total_quitados, total_nao_pagaram, total_parcial, total_caixa, total_a_receber

# Totais iniciais
total_pessoas, total_quitados, total_nao_pagaram, total_parcial, total_caixa, total_a_receber = atualizar_totais(df)

# Criar e salvar uma planilha apenas com quem tem pendências (Falta > 0)
planilha_devedora = df[df['Falta'] > 0]
planilha_devedora.to_excel("pendentes.xlsx", index=False)

# Menu interativo para mostrar os resultados
while True:
    print("\n" + "=" * 36)
    print("=== SISTEMA DE CONTROLE DE CAMISAS ===")
    print("=" * 36)
    print("1. Ver resumo financeiro")
    print("2. Gerar planilha de pendentes")
    print("3. Dar baixa / registrar pagamento")
    print("4. Cadastar novo irmão")
    print("5. Sair")
    print("=" * 36)

    opcao = input("Escolha uma opção: ")

    if opcao == "1":
        print("\n=== RELATÓRIO DE PAGAMENTO ===")
        print(f"Total de pessoas na lista: {total_pessoas}")
        print(f"Já quitaram tudo: {total_quitados}")
        print(f"Pagamentos parciais: {total_parcial}")
        print(f"Ainda não pagaram nada: {total_nao_pagaram}")
        print("=" * 34)
        print(f"💰 TOTAL EM CAIXA: R$ {total_caixa:.2f}")
        print(f"📉 TOTAL A RECEBER: R$ {total_a_receber:.2f}")
        print("=" * 34)

    elif opcao == "2":
        planilha_devedora = df[df['Falta'] > 0]
        planilha_devedora.to_excel("pendentes.xlsx", index=False)
        print("\n[OK] Arquivo 'pendentes.xlsx' gerado com sucesso!")

    elif opcao == "3":
        #NOVA VERIFICAÇÃO:  Se o Dataframe estiver vazio, cancela a operação
        if df.empty:
            print("\n[AVISO] Ainda não há nenhum irmão cadastrado no sistema.")
            print("Por favor, use a Opção 4 para cadastrar o primeiro irmão.")
            continue #Volta direto para o menu principal.3

        index_localizado = None

        while True:
            nome_busca = input("\nDigite o nome do irmão para buscar: ")
            resultado = df[df['Nome'].str.contains(nome_busca, case=False, na=False)]

            if resultado.empty:
                print("\n[AVISO] Nenhum irmão encontrado com esse nome.")
                continue

            if len(resultado) > 1:
                print("\n[Aviso] Mais de um irmão foi encontrado com esse termo:")
                print(resultado[['Nome', 'Total pago', 'Falta']].to_string(index=False))
                print("\nPor favor, refaça a busca digitando o nome de forma mais específica.")
                continue
            else:
                index_localizado = resultado.index[0]
                break

        print("\nIrmão localizado:")
        print(df.loc[[index_localizado], ['Nome', 'Total pago', 'Falta']].to_string(index=False))

        print("\n[1] Informar um valor pago (Abatimento Parcial)")
        print("\n[2] Dar baixa total (Quitar saldo devedor)")
        tipo_baixa = input("Escolha o tipo de baixa: ")

        if tipo_baixa =="1":
            while True:
                try:
                    valor_pago = float(input("\nDigite o Valor pago pelo irmão: R$ "))
                    break
                except ValueError:
                    print("\n[ERRO] Digite um valor numérico Válido (ex: 40 ou 40.50).")
            df.at[index_localizado, 'Total pago'] += valor_pago
            df.at[index_localizado, 'Falta'] -= valor_pago

        elif tipo_baixa == "2":
            falta_atual = float(df.at[index_localizado, 'Falta'])
            pago_atual = float(df.at[index_localizado, 'Total pago'])
            df.at[index_localizado, 'Total pago'] = pago_atual + falta_atual
            df.at[index_localizado, 'Falta'] = 0.0
            print(f"\n[OK] baixa total realizada! R$ {falta_atual:.2f} transferido para Total pago")

        else: 
            print("\n[AVISO] Opção de baixa invalida. Operação cancelada.")
        df.to_excel("camisas_umhb.xlsx", index=False)
        print("[OK] Planilha atualizada com sucesso!")
        total_pessoas, total_quitados, total_nao_pagaram, total_parcial, total_caixa, total_a_receber = atualizar_totais(df)
    
    elif opcao == "4":
        print("\n=== CADASTRAR NOVO IRMÃO ===")
        novo_nome = input("Nome completo: ").strip().title()
        novo_tamanho = input("Tamanho da camisa (ex: P, M, G, GG, XGG): ").strip().upper()

        while True:
            try:
                valor_camisa = float(input("Valor total da camisa: R$ "))
                break
            except ValueError:
                print("[ERRO] Digite um valor numérico válido.")
        #monta uma nova linha respeitando as colunas existentesna sua planilha
        nova_linha = {
            'Nome': novo_nome,
            'Tamanho': novo_tamanho,
            'Total pago': 0.0, #começa sem pagar noda por padrão
            'Falta': valor_camisa #Valor total entra como falta inicial
        }
        # Adicionar nova linha no final do dataframe
        df.loc[len(df)] = nova_linha
        df.to_excel("camisas_umhb.xlsx", index=False)

        print(f"\n[OK] {novo_nome} cadastrado com sucesso!")
        total_pessoas, total_quitados, total_nao_pagaram, total_parcial, total_caixa, total_a_receber = atualizar_totais(df)

    elif opcao == "5":
        print("\nSaindo do sistema... Até logo! 😃")
        break

    else:
        print("\n[ERRO] Opção inválida! Tente novamente.")
