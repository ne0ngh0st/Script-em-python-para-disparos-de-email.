import pandas as pd
import win32com.client as win32
import os
import base64

# Definir os templates para cada segmento
TEMPLATES = {
    "Saúde": {
        "subject": "Soluções Personalizadas em Papel e Suprimentos - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Meu nome é Antonio e represento a Autopel Soluções, fornecedora especializada em produtos essenciais para o setor de saúde, como papéis e suprimentos personalizados. Atendemos empresas como a <strong>{nome_empresa}</strong>, oferecendo soluções de qualidade para facilitar a operação e garantir conformidade regulatória.</p>

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

        """
    },
    "Educação": {
        "subject": "Soluções Personalizadas em Papel e Suprimentos - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou super animado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
<p>Somos especializados em produtos que facilitam o dia a dia de instituições educacionais, e sabemos o quão importante é ter materiais de alta qualidade para garantir o bom funcionamento da sua rotina administrativa.</p>

<p>Temos o que você precisa para otimizar processos, aumentar a produtividade e garantir eficiência! Confira alguns dos itens que podemos fornecer:</p>

<ul>
    <li>✅ Papéis de alta qualidade para impressão, cópias e uso diário</li>
    <li>✅ Itens de escritório essenciais, como canetas, pastas, e mais</li>
    <li>✅ Bobinas para impressoras, perfeitas para gestão de documentos e controle</li>
</ul>

<p>Em anexo, envio nosso catálogo com uma variedade de soluções que com certeza vão contribuir para o sucesso das suas atividades.</p>

<p>Que tal marcarmos uma conversa para que eu possa entender suas necessidades e juntos encontrarmos as melhores soluções para sua instituição? Tenho certeza de que podemos fazer a diferença!</p>

<p>Fico à disposição para esclarecer qualquer dúvida e aguardo seu retorno.</p>

<p>Atenciosamente,</p>

        """
    },
     "Logística": {
        "subject": "Soluções em Etiquetas e Suprimentos para Logística - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que, no setor de logística, eficiência e organização são essenciais. Por isso, estamos aqui para fornecer produtos que vão otimizar o seu trabalho e garantir o sucesso da operação!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Etiquetas adesivas de alta qualidade para facilitar a identificação e rastreabilidade dos seus produtos</li>
    <li>✅ Sulfite de alta gramatura e qualidade para garantir impressões nítidas e duradouras</li>
    <li>✅ Soluções personalizadas para atender às demandas específicas da sua logística</li>
</ul>

<p>Em anexo, envio nosso catálogo com todos os produtos que podem melhorar a eficiência e organização da sua operação.</p>

<p>Que tal agendarmos uma conversa para entender melhor como podemos ajudar a <strong>{nome_empresa}</strong>? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu negócio!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "TECH": {
        "subject": "Soluções em Bobinas e Suprimentos para Tecnologia - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p>
<p> Sabemos que no setor de tecnologia, a eficiência operacional e a precisão são fundamentais. Por isso, temos os produtos ideais para garantir que suas operações diárias sejam ainda mais eficientes!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas de ponto eletrônico para registrar dados com precisão e agilidade</li>
    <li>✅ Outras bobinas para equipamentos de impressão e controle de processos internos</li>
    <li>✅ Papel de alta qualidade para impressões e relatórios técnicos</li>
</ul>

<p>Em anexo, envio nosso catálogo com todos os produtos que podem ajudar a otimizar os processos da sua empresa.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para esclarecer qualquer dúvida e encontrar as melhores soluções para <strong>{nome_empresa}</strong>.</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Indústria": {
        "subject": "Solução em Etiquetas, Lacres e Rótulos para {nome_empresa}",
        "body": """<p>Olá!!</p>

<p>Meu nome é Antonio, e represento a Autopel Soluções, fornecedora especializada em etiquetas, rótulos e lacres de segurança para indústrias como a <strong>{nome_empresa}</strong>.</p>

<p>Trabalhamos com materiais de alta qualidade e personalizados para atender às suas necessidades, incluindo:</p>

<ul>
    <li>✅ Etiquetas e rótulos adesivos para embalagens e logística</li>
    <li>✅ Lacres de segurança para controle e autenticidade</li>
    <li>✅ Soluções personalizadas para diferentes aplicações</li>
</ul>

<p>Gostaria de agendar uma conversa para apresentar nossas soluções?</p>

<p>Atenciosamente,<br>
Antonio<br>
Autopel Soluções</p>
        """
    },
     "Comércio Varejista": {
        "subject": "Soluções em Bobinas Térmicas e Suprimentos para Varejo - {nome_empresa}",
        "body": """O<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou muito animado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
<p>Sabemos como o setor de varejo depende da agilidade e eficiência para garantir que cada transação ocorra sem contratempos. Por isso, temos os produtos ideais para facilitar os processos do seu negócio!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas térmicas de alta qualidade para seu sistema de PDV</li>
    <li>✅ Papel de excelente gramatura para impressões nítidas e duradouras</li>
    <li>✅ Soluções para otimizar o controle de vendas e garantir o bom funcionamento das operações</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos oferecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Restaurante": {
        "subject": " Soluções em Bobinas Térmicas e Suprimentos para Restaurantes - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
<p>Sabemos que no setor de restaurantes, a agilidade e a precisão são fundamentais para garantir a satisfação dos clientes e o bom funcionamento da operação. Por isso, temos os produtos ideais para garantir que tudo corra bem!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas térmicas para impressoras de PDV e comanda</li>
    <li>✅ Papel de alta qualidade para cardápios, impressões de pedidos e recibos</li>
    <li>✅ Soluções para otimizar os processos internos e melhorar a experiência do cliente</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu restaurante!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Serviços": {
        "subject": "Suprimentos Eficientes para Empresas de Serviços - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p>
<p>Sabemos que no setor de serviços, a eficiência e a organização são essenciais para garantir uma operação ágil e sem erros. Por isso, temos os produtos ideais para ajudar a otimizar seus processos!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas para sistemas de impressão e controle de dados</li>
    <li>✅ Papel de alta qualidade para impressões de documentos, recibos e relatórios</li>
    <li>✅ Soluções para melhorar a agilidade e a precisão nos seus processos internos</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Concessionária": {
        "subject": "Soluções em Bobinas e Suprimentos para Concessionárias - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou animado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
<p>Sabemos que as concessionárias precisam de agilidade e precisão nos processos de venda e documentação, e é por isso que temos os produtos ideais para atender às suas necessidades!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas para sistemas de impressão de recibos e vendas</li>
    <li>✅ Papel de alta gramatura para contratos e documentos diversos</li>
    <li>✅ Soluções para otimizar a gestão e organização interna</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu negócio!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
<p>Antonio [Seu Sobrenome]</p>
<p>Autopel Soluções</p>
        """
    },
     "Banco": {
        "subject": "Soluções em Bobinas e Suprimentos para Bancos - {nome_empresa}",
        "body": """Olá,

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
        "body": """Olá,

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
        "body": """Olá,

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
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
<p>Sabemos que os bancos lidam com grande volume de documentos e transações diariamente, e por isso, oferecemos produtos que ajudam a manter a operação eficiente, segura e sem erros!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas para impressão de comprovantes, extratos e recibos</li>
    <li>✅ Papel de alta qualidade para contratos, relatórios e documentos bancários</li>
    <li>✅ Soluções personalizadas para otimizar os processos internos e garantir a conformidade</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu banco!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Energia": {
        "subject": "Soluções em Bobinas e Suprimentos para Empresas de Energia - {nome_empresa}",
        "body": """<p> Olá, tudo bem?<p>

<p>Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que o setor de energia lida com grande volume de dados e informações, e por isso é fundamental ter materiais de alta qualidade para garantir a eficiência e segurança das operações. Temos os produtos ideais para atender às suas necessidades!</p>

    <p><strong>Oferecemos:</strong></p>
    <ul>
        <li>✅ Bobinas para sistemas de registro e impressão de dados técnicos e relatórios</li>
        <li>✅ Papel de alta qualidade para relatórios operacionais, documentos e contratos</li>
        <li>✅ Soluções personalizadas para otimizar processos e garantir a precisão nas suas operações</li>
    </ul>

    <p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

    <p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>

    <p>Fico aguardando seu retorno.</p>

    <p>Atenciosamente,</p>
    <p><strong>Antonio</strong><br>Autopel Soluções</p>
        """
    },
     "Eventos": {
        "subject": " Soluções em Bobinas e Suprimentos para Empresas de Eventos - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>

<p>Sou Antonio, da Autopel Soluções, e estou super empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que o setor de eventos exige materiais de alta qualidade para garantir a agilidade, organização e o sucesso das operações. É por isso que temos os produtos ideais para atender às suas necessidades!</p>

<p>Oferecemos:</p>

<ul>
    <li>✅ Bobinas para sistemas de emissão de ingressos, bilhetes e recibos</li>
    <li>✅ Papel de alta qualidade para convites, programas de eventos e material promocional</li>
    <li>✅ Soluções para otimizar o controle e organização dos processos de registro e atendimento ao público</li>
</ul>

<p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

<p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e encontrar as melhores soluções para o seu evento!</p>

<p>Fico aguardando seu retorno.</p>

<p>Atenciosamente,</p>
        """
    },
     "Hotelaria": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor de Hotelaria - {nome_empresa}",
        "body": """<p>Olá,</p>

<p>Tudo bem?</p>
        <p>Sou Antonio, da Autopel Soluções, e estou muito empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que o setor de hotelaria precisa de materiais que ajudem a garantir uma experiência de qualidade e organização para os hóspedes, além de otimizar os processos internos. É por isso que temos os produtos ideais para atender às suas necessidades!</p>

    <p><strong>Oferecemos:</strong></p>
    <ul>
        <li>✅ Bobinas para sistemas de registro de check-in e check-out, recibos e faturamento</li>
        <li>✅ Papel de alta qualidade para formulários, cartões de boas-vindas e outros materiais impressos</li>
        <li>✅ Soluções para otimizar o controle de reservas e documentos administrativos</li>
    </ul>

    <p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

    <p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu hotel!</p>

    <p>Fico aguardando seu retorno.</p>

    <p>Atenciosamente,</p>
    <p><strong>Antonio</strong><br>Autopel Soluções</p>
        """
    },
     "Farmacêutica": {
        "subject": "Soluções em Etiquetas, Bobinas e Suprimentos para Indústria Farmacêutica - {nome_empresa}",
        "body": """Olá,

Tudo bem?

<p>Olá,</p>
    
    <p>Tudo bem?</p>
    
    <p>Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>.<p> 
    <p>Sabemos que a indústria farmacêutica exige materiais de alta qualidade para garantir a conformidade regulatória, rastreabilidade e segurança dos produtos. Por isso, temos os produtos ideais para atender às suas necessidades!</p>
    
    <p>Oferecemos:</p>
    
    <ul>
        <li>✅ Etiquetas personalizadas para identificação de produtos e conformidade regulatória</li>
        <li>✅ Bobinas para impressão de recibos, etiquetas e outros documentos regulatórios</li>
        <li>✅ Papel de alta qualidade para rótulos, bulas e documentos de controle de produção</li>
        <li>✅ Soluções para otimizar a organização e o controle dos processos de fabricação e distribuição</li>
    </ul>
    
    <p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>
    
    <p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>
    
    <p>Fico aguardando seu retorno.</p>
    
    <p>Atenciosamente,</p>
    <p>Antonio</p>
        """
    },
    "Automotivo": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor Automotivo - {nome_empresa}",
        "body": """<p>Olá,</p>

    <p>Tudo bem?</p>

    <p>Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que o setor automotivo exige materiais de alta qualidade e eficiência para garantir o bom andamento dos processos e o controle de informações. Temos os produtos ideais para atender às suas necessidades!</p>

    <p>Oferecemos:</p>
    <ul>
        <li>✅ Bobinas para sistemas de registro e emissão de documentos, como ordens de serviço e comprovantes de pagamento</li>
        <li>✅ Papel de alta qualidade para fichas técnicas, manuais, comprovantes e outros materiais de suporte</li>
        <li>✅ Soluções personalizadas para otimizar os processos internos e garantir a organização</li>
    </ul>

    <p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

    <p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>

    <p>Fico aguardando seu retorno.</p>

    <p>Atenciosamente,</p>
        """
    },
    "Imobiliário": {
        "subject": "Soluções em Bobinas e Suprimentos para o Setor Imobiliário - {nome_empresa}",
        "body": """<p>Olá,</p>

    <p>Tudo bem?</p>

    <p>Sou Antonio, da Autopel Soluções, e estou empolgado para apresentar as soluções que podemos oferecer para <strong>{nome_empresa}</strong>. Sabemos que o setor imobiliário lida com grande volume de documentos e registros que exigem materiais de alta qualidade e organização. É por isso que temos os produtos ideais para atender às suas necessidades!</p>

    <p>Oferecemos:</p>
    <ul>
        <li>✅ Bobinas para emissão de recibos, comprovantes e contratos de aluguel ou venda</li>
        <li>✅ Papel de alta qualidade para documentos imobiliários, contratos e formulários de registro</li>
        <li>✅ Soluções para otimizar o controle de documentos e garantir a agilidade nos processos</li>
    </ul>

    <p>Em anexo, envio nosso catálogo com mais detalhes sobre os produtos que podemos fornecer para <strong>{nome_empresa}</strong>.</p>

    <p>Que tal agendarmos uma conversa para entender melhor suas necessidades? Estou à disposição para tirar qualquer dúvida e ajudar a encontrar as melhores soluções para o seu negócio!</p>

    <p>Fico aguardando seu retorno.</p>

    <p>Atenciosamente,</p>
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
df = pd.read_excel("C:\\Users\\antonio.barbosa\\Desktop\\PESSOAL\\Prog\\EnvioEmails\\Data\\carteira.xlsx")
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
                mail.CC = "sandra.schwab@autopel.com"
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