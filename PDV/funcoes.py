import pandas as pd
import os
import win32com.client as win32

usuario_ativo = os.getlogin()

def cadastrar_etapa1():
    caminho_planilha = rf"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
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
    caminho_planilha = rf"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
    consulta_informacoes = rf"C:\Users\{usuario_ativo}\Desktop\consulta_email.xlsx"

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


    df_acompanhamento.to_excel(fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx", index=False)
    print('Etapa de inclusão de informações finalizada')

def cadastrar_etapa2():
    caminho_planilha = fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
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
                print(f"Matrícula {matricula} registrada na etapa 2\n ")

            elif df_acompanhamento.at[indice, 'etapa2'] == "Aguardando envio de email":
                print(f'A matrícula {matricula} está aguardando o envio de email\n@@@@@@@@\n')

            elif df_acompanhamento.at[indice, 'etapa2'] == "Email enviado":
                pass

        else:
            print(f"Matrícula {matricula} não encontrada para o cadastramento da etapa 2\n")

    df_acompanhamento.to_excel(caminho_planilha, index=False)
    print('Etapa 2 finalizada\n')

def cadastrar_etapa3():
    caminho_planilha = fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
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

def enviar_email_etapa1(): #adicionar , remetente

    teste =  fr"C:\Users\{usuario_ativo}\Desktop\teste_disparo_email.xlsx"
    df_teste = pd.read_excel(teste)
    outlook = win32.Dispatch('outlook.application')
    #caminho_planilha = fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
    #df_acompanhamento = pd.read_excel(caminho_planilha)

    #for i, linha in df_acompanhamento.iterrows():
    for i, linha in df_teste.iterrows():
        etapa1 = linha['etapa1']

        if etapa1 == "Aguardando envio de email":

            mail = outlook.CreateItem(0)

            mail.To = linha['email']
            mail.SentOnBehalfOfName = 'pdv2024@eletronuclear.gov.br'
            mail.Subject = 'Etapa de Inscrição - PDV2024'
            mail.HTMLBody = f"""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Etapa de Inscrição - Conclusão</title>
            </head>
            <body>
                <p><strong>Prezado(a) {linha['nome']},</strong></p>
                <p>Informamos que a <strong>Etapa de Inscrição</strong> foi concluída com sucesso. Agradecemos pela sua dedicação até o momento. Agora, pedimos que aguarde o contato referente aos próximos passos.</p>
                <p>Como parte do encerramento da Etapa de Inscrição, pedimos que preencha o formulário abaixo:</p>
                <p><a href="https://forms.office.com/r/CURjbwz7ff"><strong>Formulário de Conclusão da Etapa de Inscrição</strong></a></p>
                <p>O preenchimento é essencial para a continuidade do processo.</p>
                <p>Agradecemos pela sua atenção e colaboração. Caso tenha alguma dúvida ou necessite de suporte, estamos à disposição.</p>
            </body>
            </html>                 
            """

            mail.Send()
            df_teste.at[i, 'etapa1'] = "Enviado"
            print(f"E-mail sobre etapa 1 enviado para {linha['email']} com sucesso!")

    df_teste.to_excel(teste, index=False)

def enviar_email_etapa2(): #adicionar , remetente

    teste =  fr"C:\Users\{usuario_ativo}\Desktop\teste_disparo_email.xlsx"
    df_teste = pd.read_excel(teste)
    outlook = win32.Dispatch('outlook.application')
    #caminho_planilha = fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
    #df_acompanhamento = pd.read_excel(caminho_planilha)

    #for i, linha in df_acompanhamento.iterrows():
    for i, linha in df_teste.iterrows():
        etapa1 = linha['etapa2']

        if etapa1 == "Aguardando envio de email":

            mail = outlook.CreateItem(0)

            mail.To = linha['email']
            mail.SentOnBehalfOfName = 'pdv2024@eletronuclear.gov.br'
            mail.Subject = 'Etapa de Passagem de Conhecimento - PDV2024'
            mail.HTMLBody = f"""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
            </head>
            <body>
                <p><strong>Prezado(a) {linha['nome']},</strong></p>
                <p>Informamos que a <strong>Etapa de Passagem de conhecimento (TCGC)</strong> foi concluída com sucesso. Agradecemos pela sua dedicação e comprometimento durante o processo.</p>
                <h3>Próximos passos:</h3>
                <ul>
                    <li>Caso ainda haja alguma pendência de assinatura relacionada à <strong>Etapa de Adesão</strong>, pedimos que aguarde o contato do departamento responsável.</li>
                    <li>Se o documento de Adesão já foi assinado, consideramos o processo totalmente concluído.</li>
                </ul>
                <p>Agradecemos pela sua atenção e colaboração em cada etapa. Caso tenha dúvidas ou necessite de suporte, nossa equipe está à disposição para auxiliá-lo(a).</p>
            </body>
            </html>
   
            """

            mail.Send()
            df_teste.at[i, 'etapa2'] = "Enviado"
            print(f"E-mail sobre etapa 2 enviado para {linha['email']} com sucesso!")

    df_teste.to_excel(teste, index=False)

def enviar_email_etapa3():  # adicionar , remetente

    teste = fr"C:\Users\{usuario_ativo}\Desktop\teste_disparo_email.xlsx"
    df_teste = pd.read_excel(teste)
    outlook = win32.Dispatch('outlook.application')
    # caminho_planilha = fr"C:\Users\{usuario_ativo}\Desktop\acompanhamento_email_PDV.xlsx"
    # df_acompanhamento = pd.read_excel(caminho_planilha)

    # for i, linha in df_acompanhamento.iterrows():
    for i, linha in df_teste.iterrows():
        etapa1 = linha['etapa3']

        if etapa1 == "Aguardando envio de email":
            mail = outlook.CreateItem(0)

            mail.To = linha['email']
            mail.SentOnBehalfOfName = 'pdv2024@eletronuclear.gov.br'
            mail.Subject = 'Etapa de Adesão - PDV2024'
            mail.HTMLBody = f"""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
            </head>
            <body>
                <p><strong>Prezado(a) {linha['nome']},</strong></p>
                <p>Informamos que a <strong>Etapa de Adesão</strong> foi concluída com sucesso. Agradecemos pela sua dedicação e comprometimento durante o processo.</p>
                <h3>Próximos passos:</h3>
                <ul>
                    <li>Caso ainda haja alguma pendência de assinatura relacionada à <strong>Etapa de Passagem de conhecimento (TCGC)</strong>, pedimos que aguarde o contato do departamento responsável.</li>
                    <li>Se o documento de Passagem de Conhecimento já foi assinado, consideramos o processo totalmente concluído.</li>
                </ul>
                <p>Agradecemos pela sua atenção e colaboração em cada etapa. Caso tenha dúvidas ou necessite de suporte, nossa equipe está à disposição para auxiliá-lo(a).</p>
            </body>
            </html>


            """

            mail.Send()
            df_teste.at[i, 'etapa3'] = "Enviado"
            print(f"E-mail sobre etapa 3 enviado para {linha['email']} com sucesso!")

    df_teste.to_excel(teste, index=False)



    #@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
#to do:
        # substituir o caminho para o arquivo original, invés do teste.
        # também substituir o remetente: botar o email do pdv2024.