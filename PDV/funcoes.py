import pandas as pd
import os

def cadastrar_etapa1():
    caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
    etapa_um_assinados = r"D:\Pastas - SharePoint\SP - PDV\eletronuclear.gov.br\PDV - Documentos\PDV 2024\Documentos\1 - Inscrição\2 - Com assinatura"
    df_acompanhamento = pd.read_excel(caminho_planilha)

    arquivos = [arquivo for arquivo in os.listdir(etapa_um_assinados) if arquivo != "desktop.ini"]

    for arquivo in arquivos:

        matricula = str(arquivo[2:8])

        if matricula not in df_acompanhamento["matricula"].astype(str).values:
            novo_registro = {"matricula": matricula, "etapa1": "Registro criado"}

            df_acompanhamento = pd.concat([df_acompanhamento, pd.DataFrame([novo_registro])], ignore_index=True)

            print(f'Matricula {matricula} incluída.\n')

    df_acompanhamento.to_excel(caminho_planilha, index=False)
    print('Etapa 1 finalizada.\n')


def incluir_informacoes():
    caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
    consulta_informacoes = r"C:\Users\jbtesch\Desktop\consulta_email.xlsx"

    df_acompanhamento = pd.read_excel(caminho_planilha)
    df_informacoes = pd.read_excel(consulta_informacoes)

    df_acompanhamento['nome'] = df_acompanhamento['nome'].astype(str)
    df_acompanhamento['email'] = df_acompanhamento['email'].astype(str)

    for i, linha in df_acompanhamento.iterrows():
        matricula = linha['matricula']
        filtro = df_informacoes[df_informacoes['matricula'] == matricula]

        if not filtro.empty and df_acompanhamento.at[i, 'etapa1'] == "Registro criado":
            nome = str(filtro.iloc[0]['nome'])  # Garante que o valor é uma string
            email = str(filtro.iloc[0]['email'])  # Garante que o valor é uma string

            # Atualiza diretamente as células da linha correspondente
            df_acompanhamento.at[i, 'nome'] = nome
            df_acompanhamento.at[i, 'email'] = email
            df_acompanhamento.at[i, 'etapa1'] = "Aguardando envio de email"
            print(f'\nMatrícula {matricula} acaba de ser alimentada.\n@@@@@@@@@\n')

        elif not filtro.empty and df_acompanhamento.at[i, 'etapa1'] != "Registro criado":
            print(f'Matrícula {matricula} já foi preenchida.\n')

        elif filtro.empty:
            print(f"Matrícula {matricula} não encontrada em df_informacoes.\n")


    df_acompanhamento.to_excel(r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx", index=False)
    print('Etapa de inclusão de informações finalizada')

def cadastrar_etapa2():
    caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
    df_acompanhamento = pd.read_excel(caminho_planilha)
    etapa_dois_assinados = r"D:\Pastas - SharePoint\SP - PDV\eletronuclear.gov.br\PDV - Documentos\PDV 2024\Documentos\2 - TCGC\1 - Com assinatura"

    arquivos = [arquivo for arquivo in os.listdir(etapa_dois_assinados) if arquivo != "desktop.ini"]

    df_acompanhamento['etapa2'] = df_acompanhamento['etapa2'].fillna("").astype(str)

    for arquivo in arquivos:
        matricula = str(arquivo[4:10])

        if matricula in df_acompanhamento["matricula"].astype(str).values:

            indice = df_acompanhamento[df_acompanhamento['matricula'] == int(matricula)].index[0]

            if df_acompanhamento.at[indice, 'etapa2'] == "":
                df_acompanhamento.at[indice, 'etapa2'] = "Aguardando envio de email"
                print(f"Matrícula {matricula} registrada na etapa 2\n")

            elif df_acompanhamento.at[indice, 'etapa2'] == "Aguardando envio de email":
                print(f'A matrícula {matricula} está aguardando o envio de email\n@@@@@@@@\n')

            elif df_acompanhamento.at[indice, 'etapa2'] == "Email enviado":
                pass

        else:
            print(f"Matrícula {matricula} não encontrada para o cadastramento da etapa 2\n")

    df_acompanhamento.to_excel(caminho_planilha, index=False)
    print('Etapa 2 finalizada\n')

def cadastrar_etapa3():
    caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
    df_acompanhamento = pd.read_excel(caminho_planilha)

    etapa_tres_assinados = r"D:\Pastas - SharePoint\SP - PDV\eletronuclear.gov.br\PDV - Documentos\PDV 2024\Documentos\3 - Adesão\2 - Com assinatura"
    arquivos = [arquivo for arquivo in os.listdir(etapa_tres_assinados) if arquivo != "desktop.ini"]

    df_acompanhamento['etapa3'] = df_acompanhamento['etapa3'].fillna("").astype(str)

    for arquivo in arquivos:
        matricula = str(arquivo[2:8])

        if matricula in df_acompanhamento["matricula"].astype(str).values:

            indice = df_acompanhamento[df_acompanhamento['matricula'] == int(matricula)].index[0]

            if df_acompanhamento.at[indice, 'etapa3'] == "":
                df_acompanhamento.at[indice, 'etapa3'] = "Aguardando envio de email"
                print(f"Matrícula {matricula} registrada na etapa 3\n")

            elif df_acompanhamento.at[indice, 'etapa3'] == "Aguardando envio de email":
                print(f'A matrícula {matricula} está aguardando o envio de email\n@@@@@@@@\n')

            elif df_acompanhamento.at[indice, 'etapa3'] == "Email enviado":
                pass

        else:
            print(f"Matrícula {matricula} não encontrada para o cadastramento da etapa 2\n")

    df_acompanhamento.to_excel(caminho_planilha, index=False)
    print('Etapa 3 finalizada\n')

