import discord
from discord.ext import commands
import pandas as pd
import asyncio

# Defina suas configurações do bot
TOKEN = 'COLOQUE O SEU TOKEN'
CHANNEL_NAME = 'teste'
FERIAS_FILE_PATH = r'H:\Drives compartilhados\0.0 SUPORTE\Acessos\PowerBi\POWERBI\FUNCIONARIOS\Afastamentos - InicioRetornoMotivo - Empregados - (RUAN).xls'
RAMAL_EMAIL_FILE_PATH = r'H:\Drives compartilhados\0.0 SUPORTE\Acessos\PowerBi\POWERBI\FUNCIONARIOS\EMPREGADOS -  CONTROLE.xls'
EMPRESAS_FILE_PATH = r'H:\Drives compartilhados\0.0 SUPORTE\Acessos\PowerBi/Relação de Empresas Vinculadas Sistemas.xls'

# Definir as intents necessárias
intents = discord.Intents.default()
intents.typing = False
intents.presences = False
intents.members = True
intents.messages = True

# Crie o bot
bot = commands.Bot(command_prefix='!', intents=intents)

# bot é iniciado
@bot.event
async def on_ready():
    print(f'{bot.user.name} has connected to Discord!')
    # Mude o status do bot
    await bot.change_presence(activity=discord.Game('Não, Trabalhando...'))
    
    # Encontre o canal onde enviaremos as notificações
    for guild in bot.guilds:
        channels = guild.text_channels
        channel = discord.utils.get(channels, name=CHANNEL_NAME)
        if channel:
            print(f"Bot conectado no canal '{CHANNEL_NAME}'")
            return
    print(f"Não foi possível encontrar o canal '{CHANNEL_NAME}'")

# Função para dividir o texto em pedaços menores
def split_message(message, max_length=2000):
    parts = []
    while len(message) > max_length:
        split_pos = message.rfind('\n', 0, max_length)
        if split_pos == -1:
            split_pos = max_length
        parts.append(message[:split_pos])
        message = message[split_pos:]
    parts.append(message)
    return parts

# Comando para verificar quem está de férias
@bot.command(name='ferias', aliases=['FERIAS'])
async def check_ferias(ctx):
    try:
        # Ler a planilha Excel, pulando a primeira linha
        df = pd.read_excel(FERIAS_FILE_PATH, skiprows=1)
        
        # Supondo que as colunas sejam 'NOME', 'INICIO FE' e 'RETORNO'
        today = pd.Timestamp.today()
        em_ferias = df[(df['INICIO FE'] <= today) & (df['RETORNO FE'] >= today)]
        futuras_ferias = df[df['INICIO FE'] > today]

        response = ""

        if not em_ferias.empty:
            response += "As seguintes pessoas estão de férias:\n"
            for index, row in em_ferias.iterrows():
                response += f"- {row['NOME']} (de {row['INICIO FE'].strftime('%d/%m/%Y')} a {row['RETORNO FE'].strftime('%d/%m/%Y')})\n"
        else:
            response += "Nenhuma pessoa está de férias no momento.\n"

        response += "\n"

        if not futuras_ferias.empty:
            response += "As seguintes pessoas entrarão de férias:\n"
            for index, row in futuras_ferias.iterrows():
                response += f"- {row['NOME']} (início em {row['INICIO FE'].strftime('%d/%m/%Y')})\n"
        else:
            response += "Nenhuma pessoa entrará de férias em breve."

        for part in split_message(response):
            await ctx.author.send(part, delete_after=48)  # Envia a resposta via DM e apaga após 58 segundos
    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao verificar as férias: {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos

# Comando para listar os ramais
@bot.command(name='ramal', aliases=['RAMAL'])
async def list_ramais(ctx):
    try:
        # Ler a planilha Excel
        df = pd.read_excel(RAMAL_EMAIL_FILE_PATH)
        
        # Converta a coluna 'Ramal' para string, remova o .0 dos números inteiros e trate valores NaN
        df['Ramal'] = df['Ramal'].apply(lambda x: str(int(x)) if pd.notna(x) else '')

        # Supondo que as colunas sejam 'Nome' e 'Ramal'
        response = "Lista de Ramais:\n"
        for index, row in df.iterrows():
            response += f"- {row['Nome']}: {row['Ramal']}\n"

        for part in split_message(response):
            await ctx.author.send(part, delete_after=48)  # Envia a resposta via DM e apaga após 58 segundos
    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao verificar os ramais: {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos

# Comando para listar os emails
@bot.command(name='email', aliases=['EMAIL'])
async def list_email(ctx):
    try:
        # Ler a planilha Excel
        df = pd.read_excel(RAMAL_EMAIL_FILE_PATH)
        
        # Supondo que as colunas sejam 'Nome' e 'Email'
        response = "Lista de Emails:\n"
        for index, row in df.iterrows():
            response += f"- {row['Nome']}: {row['Email']}\n"

        for part in split_message(response):
            await ctx.author.send(part, delete_after=58)  # Envia a resposta via DM e apaga após 58 segundos
    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao verificar os emails: {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos

# Comando para pesquisar por empresas
@bot.command(name='empresa', aliases=['EMPRESA'])
async def search_company(ctx, *, nome: str):
    try:
        # Ler a planilha Excel de empresas
        df_empresas = pd.read_excel(EMPRESAS_FILE_PATH)

        # Verificar e limpar nomes das colunas
        df_empresas.columns = df_empresas.columns.str.strip()

        # Verificar se a coluna 'USUÁRIO RESPONSÁVEL' existe
        if 'USUÁRIO RESPONSÁVEL' not in df_empresas.columns:
            raise ValueError("A coluna 'USUÁRIO RESPONSÁVEL' não foi encontrada na planilha.")
        
        # Buscar informações por nome na planilha de empresas
        empresa_info = df_empresas[df_empresas['Empresa'].str.contains(nome, case=False, na=False)]
        
        if not empresa_info.empty:
            # Agrupar as informações por 'Código E', 'Empresa', e 'Inscrição'
            grouped_info = empresa_info.groupby(['Código E', 'Empresa', 'Inscrição', 'Regime Federal I'], as_index=False).agg({
                'USUÁRIO RESPONSÁVEL': lambda x: ', '.join(x)
            })
            
            response = "Informações da Empresa:\n"
            for index, row in grouped_info.iterrows():
                response += f"Código E: {row['Código E']}\n"
                response += f"Empresa: {row['Empresa']}\n"
                response += f"Inscrição: {row['Inscrição']}\n"
                response += f"USUÁRIO RESPONSÁVEL: {row['USUÁRIO RESPONSÁVEL']}\n"
                response += f"Regime Federal I: {row['Regime Federal I']}\n\n"
        else:
            response = f"Nenhuma empresa encontrada com o nome {nome}."
        
        for part in split_message(response):
            await ctx.author.send(part, delete_after=48)  # Envia a resposta via DM e apaga após 58 segundos

    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao pesquisar por {nome}: {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos

# Comando para pesquisar por nome
@bot.command(name='nome', aliases=['NOME'])
async def search_name(ctx, *, nome: str):
    try:
        # Ler as planilhas Excel
        df_ferias = pd.read_excel(FERIAS_FILE_PATH, skiprows=1)
        df_ramal_email = pd.read_excel(RAMAL_EMAIL_FILE_PATH)
        
        # Buscar informações por nome nas planilhas
        info = []

        # Buscar informações de férias/aniversário
        if not df_ferias.empty:
            ferias_info = df_ferias[df_ferias['NOME'].str.contains(nome, case=False, na=False)]
            if not ferias_info.empty:
                for index, row in ferias_info.iterrows():
                    info.append(f"Nome: {row['NOME']}")
                    info.append(f"Está de férias: de {row['INICIO FE'].strftime('%d/%m/%Y')} a {row['RETORNO FE'].strftime('%d/%m/%Y')}")
            else:
                info.append(f"{nome} não está de férias.")

        # Buscar informações de ramal/email e departamento
        if not df_ramal_email.empty:
            ramal_email_info = df_ramal_email[df_ramal_email['Nome'].str.contains(nome, case=False, na=False)]
            if not ramal_email_info.empty:
                for index, row in ramal_email_info.iterrows():
                    info.append(f"Nome: {row['Nome']}")
                    info.append(f"Ramal: {str(int(row['Ramal'])) if pd.notna(row['Ramal']) else ''}")
                    info.append(f"Email: {row['Email']}")
                    info.append(f"Departamento: {row.get('Departamento', 'Não especificado')}")
            else:
                info.append(f"Informações de ramal/email/departamento não encontradas para {nome}.")

        # Enviar informações encontradas
        if info:
            response = "\n".join(info)
        else:
            response = f"Nenhuma informação encontrada para {nome}."

        for part in split_message(response):
            await ctx.author.send(part, delete_after=48)  # Envia a resposta via DM e apaga após 58 segundos

    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao pesquisar por {nome}: {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos

# Comando para buscar informações sobre empresas por usuário responsável
@bot.command(name='info', aliases=['INFO'])
async def info_empresa(ctx, usuario_responsavel: str):
    try:
        # Ler a planilha Excel de empresas
        df_empresas = pd.read_excel(EMPRESAS_FILE_PATH)

        # Verificar e limpar nomes das colunas
        df_empresas.columns = df_empresas.columns.str.strip()

        # Filtrar as informações pelo usuário responsável fornecido
        empresas_usuario = df_empresas[df_empresas['USUÁRIO RESPONSÁVEL'].str.contains(usuario_responsavel, case=False, na=False)]

        if not empresas_usuario.empty:
            # Contar o número de empresas por regime
            contagem_empresas = empresas_usuario['Regime Federal I'].value_counts()

            # Formatar a resposta incluindo o nome do autor do comando
            response = f"Informações para o usuário responsável '{usuario_responsavel}' (solicitado por {ctx.author.name}):\n"
            for regime, quantidade in contagem_empresas.items():
                response += f"- {quantidade} empresa(s) {regime}\n"
        else:
            response = f"Nenhuma empresa encontrada para o usuário responsável '{usuario_responsavel}'."

        for part in split_message(response):
            await ctx.author.send(part, delete_after=48)  # Envia a resposta via DM e apaga após 48 segundos

    except Exception as e:
        await ctx.author.send(f"Ocorreu um erro ao buscar informações para o usuário responsável '{usuario_responsavel}': {e}", delete_after=30)  # Envia a resposta via DM e apaga após 30 segundos


# Comando para mostrar a lista de comandos
@bot.command(name='comandos', aliases=['COMANDOS'])
async def show_commands(ctx):
    response = """
Comandos disponíveis:
- !ferias: Verifica quem está de férias.
- !ramal: Lista os ramais.
- !email: Lista os emails.
- !empresa <nomeempresa>: Pesquisa informações de empresas pelo nome.
- !nome <nome>: Pesquisa por informações de ramal, email, férias/aniversário e departamento para o nome especificado.
- !info <nome>: Mostra quantas empresas por regime tem para o usuário responsável especificado.
"""
    for part in split_message(response):
        await ctx.author.send(part, delete_after=58)  # Envia a resposta via DM e apaga após 58 segundos




# Evento para capturar todas as mensagens

@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    # Verifica se a mensagem é um comando válido
    if message.content.startswith(bot.command_prefix):
        # Processa comandos se houver
        await bot.process_commands(message)
        # Adiciona um atraso de alguns segundos antes de apagar a mensagem original do comando
        await asyncio.sleep(24)  # Ajuste o tempo de atraso conforme necessário (em segundos)
        
        if not isinstance(message.channel, discord.DMChannel):
        # Apagar a mensagem original do comando
            await message.delete()
        # Apagar as últimas 5 mensagens do canal, incluindo a mensagem do comando
            async for msg in message.channel.history(limit=16):
            
                try:
                    await msg.delete()
                    await asyncio.sleep(14)  # Atraso entre as exclusões para evitar limitação de taxa
                except discord.HTTPException:
                    pass
    else:
        if message.channel.name == CHANNEL_NAME:
            response = """
Comandos disponíveis:
- !ferias: Verifica quem está de férias.
- !ramal: Lista os ramais.
- !email: Lista os emails.
- !empresa <nomeempresa>: Pesquisa informações de empresas pelo nome.
- !nome <nome>: Pesquisa por informações de ramal, email, férias/aniversário e departamento para o nome especificado.
- !info <nome>: Mostra quantas empresas por regime tem para o usuário responsável especificado.
"""
            for part in split_message(response):
                await message.channel.send(part, delete_after=48)  # Envia a resposta no canal e apaga após 58 segundos
        else:
            # Processa comandos se houver
            await bot.process_commands(message)

# Inicie o bot
bot.run(TOKEN)
