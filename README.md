# DiscordBotContabilNtw
Este projeto é um bot para Discord desenvolvido em Python, usando a biblioteca discord.py, projetado para ajudar no gerenciamento de informações organizacionais contábeis do escritório. O bot oferece uma série de comandos úteis para consulta e gerenciamento de dados como férias de funcionários, ramais, e-mails e informações sobre empresas.

## Funcionalidades

O bot possui os seguintes comandos:

- `!ferias`: Verifica quem está de férias no momento e quem entrará em férias em breve.
- `!ramal`: Lista todos os ramais cadastrados.
- `!email`: Lista todos os e-mails cadastrados.
- `!empresa <nomeempresa>`: Pesquisa informações sobre empresas pelo nome.
- `!nome <nome>`: Pesquisa por informações de ramal, e-mail, férias/aniversário e departamento para o nome especificado.
- `!info <nome>`: Mostra quantas empresas por regime existem para o usuário responsável especificado.
- `!comandos`: Mostra uma lista de todos os comandos disponíveis.

## Configuração

### Requisitos

- Python 3.x
- Bibliotecas Python:
  - `discord.py`
  - `pandas`
  - `openpyxl` (para ler arquivos Excel `.xlsx`)

### Instalação

1. Clone o repositório:

   ```bash
   git clone https://github.com/seuusuario/discord-bot.git
   cd discord-bot

