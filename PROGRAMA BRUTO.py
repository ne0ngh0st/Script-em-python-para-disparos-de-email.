import pandas as pd
import win32com.client as win32
import os

# Ler a planilha de clientes
df = pd.read_excel("carteira.xlsx")

# Remover espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# Verificar se as colunas necessárias existem no DataFrame
colunas_necessarias = ["SERVIÇOS", "EMPRESA", "EMAIL"]
if not all(coluna in df.columns for coluna in colunas_necessarias):
    raise ValueError(f"A planilha deve conter as colunas: {colunas_necessarias}")

# Solicitar o segmento desejado
segmento_desejado = input("Digite o segmento desejado: ")

# Filtrar a planilha pelo segmento na coluna "SERVIÇOS"
df_segmento = df[df["SERVIÇOS"].str.contains(segmento_desejado, case=False, na=False)]

if df_segmento.empty:
    print(f"Nenhuma empresa encontrada para o segmento: {segmento_desejado}")
else:
    print(f"Encontradas {len(df_segmento)} empresas para o segmento: {segmento_desejado}")

    # Criar conexão com o Outlook
    outlook = win32.Dispatch("Outlook.Application")

    # Caminho do catálogo
    caminho_catalogo = r"C:\Users\antonio.barbosa\Desktop\PESSOAL\Prog\EnvioEmails\CATALOGO 2025.pdf"
    if not os.path.exists(caminho_catalogo):
        raise FileNotFoundError(f"O arquivo do catálogo não foi encontrado em: {caminho_catalogo}")

    for index, row in df_segmento.iterrows():
        nome_empresa = row["EMPRESA"]
        email_destino = row["EMAIL"]

        # Verificar se o e-mail é válido
        if pd.notna(email_destino) and "@" in email_destino:
            try:
                # Criar o e-mail
                mail = outlook.CreateItem(0)
                mail.To = email_destino
                mail.CC = "sandra.schwab@autopel.com"
                mail.Subject = f"Solução em Etiquetas, Lacres e Rótulos para {nome_empresa}"

                corpo_email = f"""
                Olá!!

                Meu nome é Antonio, e represento a Autopel Soluções, fornecedora especializada em etiquetas, rótulos e lacres de segurança para indústrias como a {nome_empresa}.

                Sabemos que a identificação correta de produtos é essencial para garantir rastreabilidade, segurança e conformidade com as normas do setor. Trabalhamos com materiais de alta qualidade e personalizados para atender às suas necessidades, incluindo:

                ✅ Etiquetas e rótulos adesivos para embalagens e logística  
                ✅ Lacres de segurança para controle e autenticidade  
                ✅ Etiquetas personalizadas para diferentes aplicações industriais  

                Gostaria de entender melhor suas demandas e avaliar como podemos contribuir com soluções eficientes para sua empresa. Podemos agendar uma conversa nos próximos dias?

                Fico à disposição para qualquer dúvida e aguardo seu retorno.

                Atenciosamente,  
                Antonio  
                Autopel Soluções  

                (Veja o catálogo de nossos produtos em anexo.)
                """

                mail.Body = corpo_email

                # Adicionar o catálogo como anexo
                mail.Attachments.Add(caminho_catalogo)

                # Enviar e-mail (se quiser revisar antes de enviar, troque Send() por Display())
                mail.Display()  # Use mail.Send() para enviar diretamente

                print(f"E-mail enviado para {email_destino} com sucesso!")
            except Exception as e:
                print(f"Erro ao enviar e-mail para {email_destino}: {e}")
        else:
            print(f"E-mail inválido para {nome_empresa}: {email_destino}")

    print("Processo de envio de e-mails concluído!")