import pandas as pd
import win32com.client as win32
import os

# Definir os templates para cada segmento
TEMPLATES = {
    "Saúde": {
        "subject": "Soluções Personalizadas em Papel e Suprimentos - {nome_empresa}",
        "body": """
        Olá,

Meu nome é Antonio e represento a Autopel Soluções, fornecedora especializada em produtos essenciais para o setor de saúde, como papéis e suprimentos personalizados. Atendemos empresas como a {nome_empresa}, oferecendo soluções de qualidade para facilitar a operação e garantir conformidade regulatória.

Trabalhamos com itens específicos para o seu segmento, incluindo:

✅ Papel para impressoras de alta qualidade para setores hospitalares e clínicos
✅ Bobinas de senha para hospitais e unidades de saúde
✅ Suprimentos personalizados para organização e rastreabilidade de processos internos

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos oferecer para atender às suas necessidades.

Gostaria de entender melhor suas demandas e discutir como podemos fornecer soluções eficientes para sua empresa. Podemos agendar uma conversa nos próximos dias?

Fico à disposição para qualquer dúvida e aguardo seu retorno.

Atenciosamente,
        """
    },
    "Educação": {
        "subject": "Soluções Personalizadas em Papel e Suprimentos - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou super animado para apresentar as soluções que podemos oferecer para {nome_empresa}. Somos especializados em produtos que facilitam o dia a dia de instituições educacionais, e sabemos o quão importante é ter materiais de alta qualidade para garantir o bom funcionamento da sua rotina administrativa.

Temos o que você precisa para otimizar processos, aumentar a produtividade e garantir eficiência! Confira alguns dos itens que podemos fornecer:

✅ Papéis de alta qualidade para impressão, cópias e uso diário
✅ Itens de escritório essenciais, como canetas, pastas, e mais
✅ Bobinas para impressoras, perfeitas para gestão de documentos e controle

Em anexo, envio nosso catálogo com uma variedade de soluções que com certeza vão contribuir para o sucesso das suas atividades.

Que tal marcarmos uma conversa para que eu possa entender suas necessidades e juntos encontrarmos as melhores soluções para sua instituição? Tenho certeza de que podemos fazer a diferença!

Fico à disposição para esclarecer qualquer dúvida e aguardo seu retorno.

Atenciosamente,
        """
    },
     "Logística": {
        "subject": "Soluções em Etiquetas e Suprimentos para Logística - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que, no setor de logística, eficiência e organização são essenciais. Por isso, estamos aqui para fornecer produtos que vão otimizar o seu trabalho e garantir o sucesso da operação!

Oferecemos:

✅ Etiquetas adesivas de alta qualidade para facilitar a identificação e rastreabilidade dos seus produtos
✅ Sulfite de alta gramatura e qualidade para garantir impressões nítidas e duradouras
✅ Soluções personalizadas para atender às demandas específicas da sua logística

Em anexo, envio nosso catálogo com todos os produtos que podem melhorar a eficiência e organização da sua operação.

Que tal agendarmos uma conversa para entender melhor como podemos ajudar a {nome_empresa}? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "TECH": {
        "subject": "Soluções em Bobinas e Suprimentos para Tecnologia - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que no setor de tecnologia, a eficiência operacional e a precisão são fundamentais. Por isso, temos os produtos ideais para garantir que suas operações diárias sejam ainda mais eficientes!

Oferecemos:

✅ Bobinas de ponto eletrônico para registrar dados com precisão e agilidade
✅ Outras bobinas para equipamentos de impressão e controle de processos internos
✅ Papel de alta qualidade para impressões e relatórios técnicos

Em anexo, envio nosso catálogo com todos os produtos que podem ajudar a otimizar os processos da sua empresa.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para esclarecer qualquer dúvida e encontrar as melhores soluções para {nome_empresa}.

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Indústria": {
        "subject": "Solução em Etiquetas, Lacres e Rótulos para {nome_empresa}",
        "body": """
        Olá!!

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
    },
     "Comércio Varejista": {
        "subject": "Soluções em Bobinas Térmicas e Suprimentos para Varejo - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito animado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos como o setor de varejo depende da agilidade e eficiência para garantir que cada transação ocorra sem contratempos. Por isso, temos os produtos ideais para facilitar os processos do seu negócio!

Oferecemos:

✅ Bobinas térmicas de alta qualidade para seu sistema de PDV
✅ Papel de excelente gramatura para impressões nítidas e duradouras
✅ Soluções para otimizar o controle de vendas e garantir o bom funcionamento das operações

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos oferecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Restaurante": {
        "subject": " Soluções em Bobinas Térmicas e Suprimentos para Restaurantes - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que no setor de restaurantes, a agilidade e a precisão são fundamentais para garantir a satisfação dos clientes e o bom funcionamento da operação. Por isso, temos os produtos ideais para garantir que tudo corra bem!

Oferecemos:

✅ Bobinas térmicas para impressoras de PDV e comanda
✅ Papel de alta qualidade para cardápios, impressões de pedidos e recibos
✅ Soluções para otimizar os processos internos e melhorar a experiência do cliente

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu restaurante!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Serviços": {
        "subject": "Suprimentos Eficientes para Empresas de Serviços - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que no setor de serviços, a eficiência e a organização são essenciais para garantir uma operação ágil e sem erros. Por isso, temos os produtos ideais para ajudar a otimizar seus processos!

Oferecemos:

✅ Bobinas para sistemas de impressão e controle de dados
✅ Papel de alta qualidade para impressões de documentos, recibos e relatórios
✅ Soluções para melhorar a agilidade e a precisão nos seus processos internos

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Concessionária": {
        "subject": "Soluções em Bobinas e Suprimentos para Concessionárias - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou animado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que as concessionárias precisam de agilidade e precisão nos processos de venda e documentação, e é por isso que temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para sistemas de impressão de recibos e vendas
✅ Papel de alta gramatura para contratos e documentos diversos
✅ Soluções para otimizar a gestão e organização interna

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
Antonio [Seu Sobrenome]
Autopel Soluções
        """
    },
     "Banco": {
        "subject": "Soluções em Bobinas e Suprimentos para Bancos - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que os bancos lidam com grande volume de documentos e transações diariamente, e por isso, oferecemos produtos que ajudam a manter a operação eficiente, segura e sem erros!

Oferecemos:

✅ Bobinas para impressão de comprovantes, extratos e recibos
✅ Papel de alta qualidade para contratos, relatórios e documentos bancários
✅ Soluções personalizadas para otimizar os processos internos e garantir a conformidade

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu banco!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Grafica": {
        "subject": "Soluções em Bobinas e Papel para Gráficas - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que as gráficas dependem de materiais de alta qualidade para garantir impressões perfeitas, e é por isso que temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para impressão de alta qualidade e eficiência em seus processos gráficos
✅ Papel de excelente gramatura e acabamento para garantir impressões nítidas e duradouras
✅ Soluções personalizadas para atender às demandas específicas de impressão da sua gráfica

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Estacionamento": {
        "subject": "Soluções em Bobinas e Suprimentos para Estacionamentos - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou animado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que os estacionamentos lidam com um grande volume de registros diários e precisam de materiais que garantam agilidade e organização. Por isso, temos os produtos ideais para facilitar o seu trabalho!

Oferecemos:

✅ Bobinas para sistemas de controle de entradas e saídas de veículos
✅ Papel de alta qualidade para emissão de recibos, ticket de estacionamento e relatórios
✅ Soluções para garantir uma gestão eficiente e otimização dos processos

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu estacionamento!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Segurança": {
        "subject": "Soluções em Etiquetas, Bobinas e Suprimentos para Segurança - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor de segurança exige materiais confiáveis e eficientes para garantir a rastreabilidade, controle e autenticidade das operações. Por isso, temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Etiquetas personalizadas para controle de ativos e rastreabilidade
✅ Bobinas para impressoras de registro de dados e relatórios de segurança
✅ Papel de alta qualidade para relatórios, documentos e registros de ocorrências
✅ Soluções para otimizar o controle e garantir a segurança das operações

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Energia": {
        "subject": "Soluções em Bobinas e Suprimentos para Empresas de Energia - {nome_empresa}",
        "body": """
        Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor de energia lida com grande volume de dados e informações, e por isso é fundamental ter materiais de alta qualidade para garantir a eficiência e segurança das operações. Temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para sistemas de registro e impressão de dados técnicos e relatórios
✅ Papel de alta qualidade para relatórios operacionais, documentos e contratos
✅ Soluções personalizadas para otimizar processos e garantir a precisão nas suas operações

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Eventos": {
        "subject": " Soluções em Bobinas e Suprimentos para Empresas de Eventos - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor de eventos exige materiais de alta qualidade para garantir a agilidade, organização e o sucesso das operações. É por isso que temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para sistemas de emissão de ingressos, bilhetes e recibos
✅ Papel de alta qualidade para convites, programas de eventos e material promocional
✅ Soluções para otimizar o controle e organização dos processos de registro e atendimento ao público

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu evento!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Hotelaria": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor de Hotelaria - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor de hotelaria precisa de materiais que ajudem a garantir uma experiência de qualidade e organização para os hóspedes, além de otimizar os processos internos. É por isso que temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para sistemas de registro de check-in e check-out, recibos e faturamento
✅ Papel de alta qualidade para formulários, cartões de boas-vindas e outros materiais impressos
✅ Soluções para otimizar o controle de reservas e documentos administrativos

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu hotel!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
     "Farmacêutica": {
        "subject": "Soluções em Etiquetas, Bobinas e Suprimentos para Indústria Farmacêutica - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que a indústria farmacêutica exige materiais de alta qualidade para garantir a conformidade regulatória, rastreabilidade e segurança dos produtos. Por isso, temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Etiquetas personalizadas para identificação de produtos e conformidade regulatória
✅ Bobinas para impressão de recibos, etiquetas e outros documentos regulatórios
✅ Papel de alta qualidade para rótulos, bulas e documentos de controle de produção
✅ Soluções para otimizar a organização e o controle dos processos de fabricação e distribuição

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
    "Automotivo": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor Automotivo - {nome_empresa}",
        "body": """
        Olá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor automotivo exige materiais de alta qualidade e eficiência para garantir o bom andamento dos processos e o controle de informações. Temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para sistemas de registro e emissão de documentos, como ordens de serviço e comprovantes de pagamento
✅ Papel de alta qualidade para fichas técnicas, manuais, comprovantes e outros materiais de suporte
✅ Soluções personalizadas para otimizar os processos internos e garantir a organização

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,
        """
    },
    "Imobiliário": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor Imobiliário - {nome_empresa}",
        "body": """
        XXOlá,

Tudo bem?

Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para {nome_empresa}. Sabemos que o setor imobiliário lida com grande volume de documentos e registros que exigem materiais de alta qualidade e organização. É por isso que temos os produtos ideais para atender às suas necessidades!

Oferecemos:

✅ Bobinas para emissão de recibos, comprovantes e contratos de aluguel ou venda
✅ Papel de alta qualidade para documentos imobiliários, contratos e formulários de registro
✅ Soluções para otimizar o controle de documentos e garantir a agilidade nos processos

Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para {nome_empresa}.

Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!

Fico aguardando seu retorno.

Atenciosamente,X
        """
    },

    # Adicione mais segmentos conforme necessário
    "DEFAULT": {
        "subject": "Solução em Etiquetas, Lacres e Rótulos para {nome_empresa}",
        "body": """
        Olá!!

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
df = pd.read_excel("carteira.xlsx")
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
    caminho_catalogo = r"C:\Users\antonio.barbosa\Desktop\PESSOAL\Prog\EnvioEmails\CATALOGO 2025.pdf"
    
    if not os.path.exists(caminho_catalogo):
        raise FileNotFoundError(f"Arquivo do catálogo não encontrado: {caminho_catalogo}")

    # Ler a planilha de clientes
df = pd.read_excel("carteira.xlsx")
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
    caminho_catalogo = r"C:\Users\antonio.barbosa\Desktop\PESSOAL\Prog\EnvioEmails\CATALOGO 2025.pdf"
    
    if not os.path.exists(caminho_catalogo):
        raise FileNotFoundError(f"Arquivo do catálogo não encontrado: {caminho_catalogo}")
        
        # Perguntar ao usuário se deseja analisar ou enviar diretamente
    opcao = input("Deseja analisar os e-mails antes de enviar? (S/N): ").strip().upper()
    analisar_antes = opcao == "S"


    # Enviar e-mails
    for index, row in df_segmento.iterrows():
        nome_empresa = row["EMPRESA"]
        email_destino = row["EMAIL"]

        if pd.notna(email_destino) and "@" in email_destino:
            try:
                mail = outlook.CreateItem(0)
                mail.To = email_destino
                mail.CC = "sandra.schwab@autopel.com"
                mail.Subject = template["subject"].format(nome_empresa=nome_empresa)
                mail.Body = template["body"].format(nome_empresa=nome_empresa)
                mail.Attachments.Add(caminho_catalogo)

                mail.Display()  # Troque para mail.Send() para enviar diretamente
                print(f"E-mail para {email_destino} (Segmento: {segmento_desejado}) criado!")
            except Exception as e:
                print(f"Erro ao enviar para {email_destino}: {str(e)}")
        else:
            print(f"E-mail inválido para {nome_empresa}: {email_destino}")
if analisar_antes:
    mail.Display()
    print(f"E-mail para {email_destino} exibido para análise.")
else:
    mail.Send()
    print(f"E-mail para {email_destino} enviado com sucesso!")
print("Processo concluído!")
