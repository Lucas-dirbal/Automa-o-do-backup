# Automacao de Relatorios + WhatsApp

Projeto local para:

- baixar e processar relatorios via Selenium;
- gerar planilhas de pendencias;
- separar arquivos por representante;
- enviar mensagens e anexos pelo WhatsApp usando Baileys;
- acompanhar tudo por um painel web em tempo real.

O projeto foi pensado principalmente para Windows.

## Visao geral

O fluxo principal funciona assim:

1. O servidor Node.js sobe o painel web em `http://localhost:2602`.
2. Esse servidor mantem a conexao do WhatsApp via Baileys.
3. Pelo painel, voce informa usuario e senha do sistema.
4. O servidor chama o script Python `baixaRel.py`.
5. O Python acessa o sistema, gera os arquivos e pede ao servidor o envio no WhatsApp.

## Componentes principais

- `servidor_baileys.js`: servidor Express + Socket.IO + integracao com WhatsApp.
- `baixaRel.py`: automacao principal com Selenium, leitura de PDF e geracao de planilhas.
- `public/`: interface web do painel.
- `auth_info_baileys/`: sessao autenticada do WhatsApp.
- `Relatorios/`: saidas principais da automacao.
- `representantes_separados/`: planilhas separadas por representante.
- `uploads/`: arquivos temporarios usados no envio.

## Requisitos

- Windows.
- Node.js instalado.
- Python 3 instalado.
- Dependencias Node instaladas com `npm install`.
- Dependencias Python instaladas no ambiente virtual.
- Brave ou Google Chrome disponivel na maquina.

Dependencias Python usadas pelo projeto:

- `selenium`
- `pandas`
- `pdfplumber`
- `requests`
- `openpyxl`

## Instalacao

### 1. Instalar dependencias Node

```powershell
npm install
```

### 2. Criar e preparar o ambiente Python

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install selenium pandas pdfplumber requests openpyxl
```

Se o projeto ja tiver a pasta `.venv` pronta, basta reutilizar.

## Como iniciar

### Servidor completo

```powershell
npm start
```

Depois abra:

```text
http://localhost:2602
```

No painel, voce pode:

- conectar o WhatsApp lendo o QR Code;
- iniciar a automacao informando usuario e senha;
- acompanhar logs, metricas e arquivos gerados;
- testar envio manual de mensagem.

### Rodar so o Python

Para depuracao, tambem e possivel executar o script diretamente:

```powershell
.\.venv\Scripts\python.exe .\baixaRel.py
```

Observacao: para o fluxo completo de envio no WhatsApp funcionar, o servidor Node precisa estar ativo.

## Variaveis de ambiente uteis

O `baixaRel.py` reconhece estas variaveis:

- `AUTOMACAO_USUARIO`: usuario do sistema.
- `AUTOMACAO_SENHA`: senha do sistema.
- `AUTOMACAO_SERVER_URL`: URL do servidor Node. Padrao: `http://localhost:2602`.

Exemplo:

```powershell
$env:AUTOMACAO_USUARIO="seu_usuario"
$env:AUTOMACAO_SENHA="sua_senha"
.\.venv\Scripts\python.exe .\baixaRel.py
```

## Inicio automatico com o Windows

O projeto ja tem scripts para registrar o servidor no Agendador de Tarefas do Windows.

Instalar:

```powershell
npm run windows:startup:install
```

Iniciar manualmente em modo oculto:

```powershell
npm run windows:start-hidden
```

Remover a inicializacao automatica:

```powershell
npm run windows:startup:remove
```

Nome da tarefa criada:

```text
AutomacaoServidorBaileys
```

## Arquivos e pastas importantes

- `mapeamento_representantes.xlsx`: precisa conter as colunas `Representante` e `Telefone`.
- `PENDENCIAS.xlsx`: planilha consolidada gerada no fluxo.
- `server.out.log`: log padrao do servidor Node.
- `server.err.log`: log de erro do servidor Node.
- `startup.log`: log do script de inicializacao automatica no Windows.

## Observacoes importantes

- A URL do sistema alvo esta definida diretamente no `baixaRel.py` em `URL_SISTEMA`.
- O projeto usa recursos especificos de Windows, como `msvcrt`, PowerShell e `taskkill`.
- A pasta `auth_info_baileys/` guarda a sessao do WhatsApp e deve ser tratada como dado sensivel.
- Na primeira execucao do Selenium, pode ser necessario que o ambiente consiga resolver o driver do navegador.

## Solucao de problemas

Se o painel abrir mas o WhatsApp nao conectar:

- verifique se o QR apareceu no painel;
- confira os logs em `server.out.log` e `server.err.log`;
- confirme se a sessao em `auth_info_baileys/` nao esta corrompida.

Se a automacao nao abrir o navegador:

- confira se Brave ou Chrome estao instalados;
- valide se o Selenium esta instalado no `.venv`;
- verifique se a versao do navegador esta compativel com o driver.

Se o envio no WhatsApp falhar:

- confirme que o servidor esta de pe em `http://localhost:2602`;
- confira se o numero esta no formato com DDI, DDD e numero;
- revise o arquivo `mapeamento_representantes.xlsx`.

## Estado atual

Hoje o projeto possui:

- painel web em tempo real;
- disparo da automacao Python pelo navegador;
- envio de mensagem e anexo no WhatsApp;
- scripts para iniciar o servidor junto com o Windows.
