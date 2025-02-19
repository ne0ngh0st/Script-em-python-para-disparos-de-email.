import pandas as pd
import win32com.client as win32
import os
import base64

# Definir os templates para cada segmento
TEMPLATES = {
    "teste": {
        "subject": "Soluções Personalizadas em Papel e Suprimentos - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Meu nome é Antonio e represento a Autopel Soluções, fornecedora especializada em produtos essenciais para o setor de TESTE, como papéis e suprimentos personalizados. Atendemos empresas como a <strong>{nome_empresa}</strong>, oferecendo soluções de qualidade para facilitar a operação e garantir conformidade regulatória.</p>

<p>Trabalhamos com itens específicos para o seu segmento, incluindo:</p>

<ul>
    <li>✅ Papel para impressoras de alta qualidade para setores hospitalares e clínicos</li>
    <li>✅ Bobinas de senha para hospitais e unidades de saúde</li>
    <li>✅ Suprimentos personalizados para organização e rastreabilidade de processos internos</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos oferecer para atender às suas necessidades.</p>

<p>Gostaria de entender melhor suas demandas e discutir como podemos fornecer soluções eficientes para sua empresa. Podemos agendar uma conversa nos próximos dias?</p>

<p>Fico à disposição para qualquer dúvida e aguardo seu retorno.</p>

<p>Atenciosamente,</p>

<img src="https://7ejh2uf3df.execute-api.us-east-2.amazonaws.com/PROD/rastreamento?email={{email}}" width="1" height="1" style="display:none;" />



        """
    },

    # Adicione mais segmentos conforme necessário
    "DEFAULT": {
        "subject": "Solução em Etiquetas, Lacres e Rótulos para {nome_empresa}",
        "body": """Olá!!

        Meu nome é Antonio, e represento a Autopel Soluções, fornecedora especializada em etiquetas, rótulos e lacres de segurança para indústrias como a {nome_empresa}.

        Trabalhamos com materiais de alta qualidade e personalizados para atender às suas necessidades, incluindo:

        ✅ Etiquetas e rótulos adesivos para embalagens e logística  
        ✅ Lacres de segurança para controle e autenticidade  
        ✅ Soluções personalizadas para diferentes aplicações  

        Gostaria de agendar uma conversa para apresentar nossas soluções?

        Atenciosamente,  
        Antonio  
        Autopel Soluções  
        """
    }

}

# Ler a planilha de clientes
df = pd.read_excel("C:\\Users\\antonio.barbosa\\Desktop\\PESSOAL\\Prog\\EnvioEmails\\Data\\empresasteste.xlsx")
df.columns = df.columns.str.strip()

# Verificar colunas necessárias
colunas_necessarias = ["SERVIÇOS", "EMPRESA", "EMAIL"]
if not all(coluna in df.columns for coluna in colunas_necessarias):
    raise ValueError(f"A planilha deve conter as colunas: {colunas_necessarias}")

# Solicitar o segmento desejado
segmento_desejado = input("Digite o segmento desejado: ").strip()

# Normalizar o segmento digitado (remover espaços e converter para minúsculas)
segmento_desejado_normalizado = segmento_desejado.lower().strip()

# Escolher o template com base no segmento (case-insensitive)
template_key = None
for key in TEMPLATES.keys():
    if key.lower() == segmento_desejado_normalizado:
        template_key = key
        break

# Usar o template DEFAULT se o segmento não for encontrado
template = TEMPLATES.get(template_key, TEMPLATES["DEFAULT"])

# Filtrar empresas do segmento
df_segmento = df[df["SERVIÇOS"].str.contains(segmento_desejado, case=False, na=False)]

if df_segmento.empty:
    print(f"Nenhuma empresa encontrada para o segmento: {segmento_desejado}")
else:
    print(f"Encontradas {len(df_segmento)} empresas para o segmento: {segmento_desejado}")

    # Configurar Outlook e caminho do catálogo
    outlook = win32.Dispatch("Outlook.Application")
    caminho_catalogo = r"C:\Users\antonio.barbosa\Desktop\PESSOAL\Prog\EnvioEmails\Data\CATALOGO 2025.pdf"
    caminho_assinatura = r"C:\Users\antonio.barbosa\Desktop\PESSOAL\Prog\EnvioEmails\Data\assinatura.png"

    if not os.path.exists(caminho_catalogo):
        raise FileNotFoundError(f"Arquivo do catálogo não encontrado: {caminho_catalogo}")
    if not os.path.exists(caminho_assinatura):
        raise FileNotFoundError(f"Arquivo da assinatura não encontrado: {caminho_assinatura}")

    # Perguntar ao usuário se deseja analisar ou enviar direto
    opcao = input("Deseja analisar os e-mails antes de enviar? (S/N): ").strip().upper()
    analisar_antes = opcao == "S"

    # Ler a imagem da assinatura e codificar em Base64
    with open(caminho_assinatura, "rb") as img_file:
        img_base64 = base64.b64encode(img_file.read()).decode("utf-8")

    # Adicionar a assinatura ao corpo do e-mail
    assinatura_html = f"""
    <br><br>
    <img src="data:image/jpg;base64,{img_base64}" alt="Assinatura">
    """

    # Enviar e-mails
    for index, row in df_segmento.iterrows():
        nome_empresa = row["EMPRESA"]
        email_destino = row["EMAIL"]

        if pd.notna(email_destino) and "@" in email_destino:
            try:
                mail = outlook.CreateItem(0)
                mail.To = email_destino
                mail.CC = ""
                mail.Subject = template["subject"].format(nome_empresa=nome_empresa)
                mail.HTMLBody = template["body"].format(nome_empresa=nome_empresa) + assinatura_html
                mail.Attachments.Add(caminho_catalogo)

                if analisar_antes:
                    mail.Display()  # Exibe o e-mail para análise
                    print(f"E-mail para {email_destino} (Segmento: {segmento_desejado}) exibido para análise.")
                else:
                    mail.Send()  # Envia o e-mail diretamente
                    print(f"E-mail para {email_destino} (Segmento: {segmento_desejado}) enviado com sucesso!")
            except Exception as e:
                print(f"Erro ao enviar para {email_destino}: {str(e)}")
        else:
            print(f"E-mail inválido para {nome_empresa}: {email_destino}")

print("Processo concluído!")
