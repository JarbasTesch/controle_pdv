import pandas as pd
import os
teste = r'''
# Caminhos dos arquivos
etapa_um_assinados = r"D:\Pastas - SharePoint\SP - PDV\eletronuclear.gov.br\PDV - Documentos\PDV 2024\Documentos\1 - Inscrição\2 - Com assinatura"
caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
consulta_informacoes = r"C:\Users\jbtesch\Desktop\consulta_email.xlsx"

# Carregando as planilhas
df_acompanhamento = pd.read_excel(caminho_planilha)
df_consulta_informacoes = pd.read_excel(consulta_informacoes)

# Limpando espaços dos nomes das colunas, caso haja
df_consulta_informacoes.columns = df_consulta_informacoes.columns.str.strip()

# Ignorando arquivos ocultos e listando apenas os arquivos válidos
arquivos = [arquivo for arquivo in os.listdir(etapa_um_assinados) if arquivo != "desktop.ini"]

for arquivo in arquivos:
    # Extraindo matrícula
    matricula = arquivo[2:8]
    print(f"Matrícula extraída: {matricula}")  # Imprimindo a matrícula para verificar

    # Verificando se a matrícula está no acompanhamento
    if matricula not in df_acompanhamento["matricula"].astype(str).values:
        # Buscando informações da matrícula na consulta
        info_matricula = df_consulta_informacoes[df_consulta_informacoes["matricula"] == matricula]

        # Verificando se a matrícula foi encontrada na consulta
        if not info_matricula.empty:
            print(f"Informações encontradas para a matrícula {matricula}.")  # Verificando se as informações foram encontradas
            nome = info_matricula.iloc[0]["nome"]
            email = info_matricula.iloc[0]["email"]

            # Criando um novo registro
            novo_registro = {"matricula": matricula, "etapa1": "Registrado", "nome": nome, "email": email}
            df_acompanhamento = pd.concat([df_acompanhamento, pd.DataFrame([novo_registro])], ignore_index=True)
        else:
            print(f"Matrícula {matricula} não encontrada no arquivo de consulta.")

# Salvando o DataFrame atualizado
df_acompanhamento.to_excel(caminho_planilha, index=False)

print("Planilha atualizada com os novos registros.")
os.startfile(caminho_planilha)

# Caminhos dos arquivos
etapa_um_assinados = r"D:\Pastas - SharePoint\SP - PDV\eletronuclear.gov.br\PDV - Documentos\PDV 2024\Documentos\1 - Inscrição\2 - Com assinatura"
caminho_planilha = r"C:\Users\jbtesch\Desktop\acompanhamento_email_PDV.xlsx"
consulta_informacoes = r"C:\Users\jbtesch\Desktop\consulta_email.xlsx"

# Carregando as planilhas
df_acompanhamento = pd.read_excel(caminho_planilha)
df_consulta_informacoes = pd.read_excel(consulta_informacoes)

# Ignorando arquivos ocultos e listando apenas os arquivos válidos
arquivos = [arquivo for arquivo in os.listdir(etapa_um_assinados) if arquivo != "desktop.ini"]

for arquivo in arquivos:
    # Extraindo matrícula
    matricula = arquivo[2:8]

    # Verificando se a matrícula está no acompanhamento
    if matricula not in df_acompanhamento["matricula"].astype(str).values:

        info_matricula = df_consulta_informacoes[df_consulta_informacoes["matricula"] == matricula]

        nome = info_matricula.iloc[0]["nome"]
        email = info_matricula.iloc[0]["email"]

        novo_registro = {"matricula": matricula, "etapa1": "Registrado", "nome": nome, "email": email}
        df_acompanhamento = pd.concat([df_acompanhamento, pd.DataFrame([novo_registro])], ignore_index=True)

# Salvando o DataFrame atualizado
df_acompanhamento.to_excel(caminho_planilha, index=False)

print("Planilha atualizada com os novos registros.")
os.startfile(caminho_planilha)'''