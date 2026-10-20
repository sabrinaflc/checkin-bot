# 🗓️ Automação de Presença em Atendimentos

[![forthebadge](http://forthebadge.com/images/badges/made-with-python.svg)](http://forthebadge.com)
[![forthebadge](http://forthebadge.com/images/badges/built-with-love.svg)](http://forthebadge.com)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**Script em Python para automatizar o preenchimento mensal de planilhas de presença a partir da agenda de um sistema web de clínica.**

O projeto utiliza **Playwright** para navegar pelo sistema e **openpyxl** para atualizar uma planilha Excel, identificando os atendimentos realizados em cada dia do mês.

Em vez de conferir manualmente cada data e cada paciente, o script executa as tarefas repetitivas e registra as presenças automaticamente.

> [!WARNING]
> **Privacidade, LGPD e sigilo profissional**
>
> Este projeto pode processar dados pessoais sensíveis relacionados à saúde. Nunca publique no GitHub planilhas reais, nomes de pacientes, capturas de tela, credenciais ou outros dados confidenciais. Consulte a seção [Privacidade e segurança](#-privacidade-e-segurança).

---

## 📑 Sumário

- [✨ Funcionalidades](#-funcionalidades)
- [⚙️ Como funciona](#️-como-funciona)
- [🧰 Tecnologias e requisitos](#-tecnologias-e-requisitos)
- [🚀 Instalação](#-instalação)
- [📊 Preparando a planilha](#-preparando-a-planilha)
- [🔧 Configuração](#-configuração)
- [▶️ Como executar](#️-como-executar)
- [🧪 Teste recomendado](#-teste-recomendado)
- [🛠️ Solução de problemas](#️-solução-de-problemas)
- [🔄 Adaptação para outros sistemas](#-adaptação-para-outros-sistemas)
- [🔒 Privacidade e segurança](#-privacidade-e-segurança)
- [⚠️ Limitações](#️-limitações)
- [📄 Licença](#-licença)

---

## ✨ Funcionalidades

- 📅 Consulta a agenda **dia a dia**, dentro do mês configurado.
- 🌐 Navegação automatizada com Chromium e Playwright.
- 🔐 Login manual, sem leitura ou armazenamento de senha pelo script.
- ✅ Identificação dos registros com status de atendimento realizado.
- 📋 Preenchimento automático da planilha Excel.
- 🔎 Comparação de nomes ignorando acentos, maiúsculas, minúsculas e espaços extras.
- 🔁 Registro de atendimentos repetidos no mesmo dia: `X`, `XX` e assim por diante.
- ➕ Inclusão automática de pacientes encontrados no sistema que ainda não constam na planilha.
- 💾 Salvamento ao final de cada dia processado.
- 🔄 Nova tentativa quando ocorre uma falha durante a consulta.
- ⚠️ Relatório dos dias que precisam de conferência manual.

## ⚙️ Como funciona

O fluxo de execução é simples:

1. O script abre o Chromium na página do sistema da clínica.
2. Você realiza o login manualmente.
3. Após confirmar o login, o script acessa **Atendimento → Horário Marcado**.
4. A especialidade configurada é selecionada.
5. O script percorre os dias do mês, consulta a agenda e lê a tabela **Realizados/Cancelados**.
6. Apenas registros com status de atendimento realizado são considerados.
7. Cada paciente é localizado na planilha e sua presença é registrada na coluna correspondente ao dia.
8. A planilha é salva ao final de cada dia processado.
9. Caso algum dia não possa ser consultado, o script tenta novamente uma vez e informa quais datas precisam de revisão.

### Como as presenças são registradas

| Situação | Registro |
|---|---|
| Um atendimento no dia | `X` |
| Dois atendimentos no mesmo dia | `XX` |
| Três atendimentos no mesmo dia | `XXX` |
| Paciente não atendido | Nenhuma marcação automática |

> **Importante:** a ausência de uma marcação não significa necessariamente que o paciente faltou. O dia pode não ter sido processado corretamente e deve ser conferido quando houver indicação de falha.

---

## 🧰 Tecnologias e requisitos

| Tecnologia | Finalidade |
|---|---|
| [Python](https://www.python.org/) 3.9 ou superior | Execução do script |
| [Playwright](https://playwright.dev/python/) | Automação do navegador |
| [openpyxl](https://openpyxl.readthedocs.io/) | Leitura e escrita de arquivos Excel |
| Chromium | Navegador utilizado na automação |

É necessário ter acesso ao sistema web, permissão para consultar a agenda e uma planilha `.xlsx` preparada no formato esperado.

## 🚀 Instalação

### 1. Clone o repositório

Substitua o endereço abaixo pelo endereço real do seu repositório.

```bash
git clone https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
cd SEU-REPOSITORIO
```

### 2. Crie um ambiente virtual

Recomendado para manter as dependências isoladas do restante do Python.

**Windows — Prompt de Comando:**

```bat
python -m venv .venv
.venv\Scripts\activate
```

**Windows — PowerShell:**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Linux/macOS:**

```bash
python -m venv .venv
source .venv/bin/activate
```

### 3. Instale as dependências

Se o repositório já possuir um arquivo `requirements.txt`:

```bash
pip install -r requirements.txt
```

Caso ainda não exista, instale os pacotes diretamente:

```bash
pip install playwright openpyxl
```

### 4. Instale o Chromium

```bash
playwright install chromium
```

Após concluir a instalação, o ambiente estará pronto para executar o script.

---

## 📊 Preparando a planilha

A planilha `.xlsx` deve seguir uma estrutura simples: nomes dos pacientes na primeira coluna e dias do mês nas colunas seguintes.

### Estrutura esperada

| Coluna A — Nome | Coluna B — Dia 1 | Coluna C — Dia 2 | Coluna D — Dia 3 |
|---|:---:|:---:|:---:|
| *(Cabeçalho)* | 1 | 2 | 3 |
| Paciente de exemplo | X | | XX |
| Outro paciente | | X | |

*Os nomes apresentados são fictícios e servem apenas para ilustrar o formato.*

### Regras da planilha

- **Linha 1:** cabeçalho; os nomes começam a ser lidos na linha 2.
- **Coluna A:** nomes dos pacientes.
- **Colunas B em diante:** uma coluna para cada dia do mês.
- **Dia 1:** coluna B.
- **Dia 2:** coluna C.
- **Dia 3:** coluna D.

A posição dos nomes e dos dias pode ser alterada nas variáveis `COL_NOME` e `COL_DIA_1`.

> [!TIP]
> Faça os primeiros testes em uma **cópia da planilha**. O script modifica e salva o arquivo informado na configuração.

---

## 🔧 Configuração

As principais opções ficam no bloco `CONFIGURAÇÃO`, no início do arquivo `preencher_presenca.py`.

| Variável | Descrição | Exemplo |
|---|---|---|
| `ANO` | Ano a processar | `2026` |
| `MES` | Mês a processar | `10` |
| `ARQUIVO` | Caminho da planilha | `"pacientes.xlsx"` |
| `COL_NOME` | Coluna dos nomes (A = 1) | `1` |
| `COL_DIA_1` | Coluna correspondente ao dia 1 (B = 2) | `2` |
| `ESPECIALIDADE` | Nome exato da especialidade no sistema | `"NOME DA ESPECIALIDADE"` |
| `DIA_INICIO` | Primeiro dia a processar | `1` |
| `DIA_FIM` | Último dia a processar; `None` percorre o mês inteiro | `None` |
| `URL` | Endereço do sistema web | `"https://..."` |

### Exemplo de configuração

```python
ANO = 2026
MES = 10

ARQUIVO = "pacientes.xlsx"

COL_NOME = 1
COL_DIA_1 = 2

ESPECIALIDADE = "NOME DA ESPECIALIDADE"

DIA_INICIO = 1
DIA_FIM = 2  # Teste inicial
```

Os nomes acima são exemplos: mantenha a configuração compatível com as variáveis realmente utilizadas no seu script.

### Configurando a URL por variável de ambiente

Para não deixar o endereço do sistema diretamente no código, utilize a variável de ambiente `SISTEMA_URL`.

**Windows — PowerShell:**

```powershell
$env:SISTEMA_URL = "https://endereco-do-seu-sistema"
```

**Linux/macOS:**

```bash
export SISTEMA_URL="https://endereco-do-seu-sistema"
```

O script precisa estar preparado para ler essa variável, por exemplo, utilizando `os.getenv("SISTEMA_URL")`.

---

## ▶️ Como executar

No terminal, com o ambiente virtual ativado, execute:

```bash
python preencher_presenca.py
```

### Passo a passo

1. Aguarde a abertura do Chromium.
2. Faça login normalmente no sistema da clínica.
3. Retorne ao terminal e pressione `Enter`, conforme solicitado pelo script.
4. Aguarde o processamento das datas.
5. Confira as mensagens exibidas no terminal.
6. Abra a planilha e valide os resultados.

### Exemplo de saída no terminal

```text
Dia 1: 8 pacientes atendidos (11 registros na tabela)
Dia 2: 6 pacientes atendidos (9 registros na tabela) | repetidos (XX): 1
Dia 3: 7 pacientes atendidos (10 registros na tabela)

Dias NÃO lidos (confira manualmente): [15]
```

A saída é ilustrativa. O formato exato depende da implementação do script.

Por privacidade, os logs devem mostrar contagens e informações operacionais, sem nomes ou outros dados identificáveis de pacientes.

---

## 🧪 Teste recomendado

Antes de processar o mês inteiro, valide o funcionamento com uma cópia da planilha.

- [ ] Faça uma cópia de segurança da planilha.
- [ ] Configure `DIA_INICIO = 1`.
- [ ] Configure `DIA_FIM = 2`.
- [ ] Execute o script.
- [ ] Compare manualmente os dois dias no sistema.
- [ ] Confirme se os status e os nomes foram interpretados corretamente.
- [ ] Verifique se os atendimentos repetidos foram registrados como esperado.
- [ ] Confira se novos pacientes foram adicionados corretamente.
- [ ] Verifique se a planilha foi salva sem perda de informações.

Após validar os resultados, configure `DIA_FIM = None` para processar o mês inteiro.

---

## 🛠️ Solução de problemas

| Problema | Causa provável | Solução |
|---|---|---|
| `NENHUM select com essa opcao...` | O texto da especialidade não corresponde à opção do sistema | Copie o nome exato exibido no menu |
| Dia não lido por erro de SQL | Instabilidade ou erro retornado pelo sistema | Confira o dia e tente processá-lo novamente |
| Tabela do dia não encontrada | Carregamento lento, ausência de agenda ou mudança na página | Confira o sistema e ajuste os seletores, se necessário |
| Pacientes não marcados | Status diferente do esperado ou falha na correspondência dos nomes | Confira a função `eh_atendido` e a normalização dos nomes |
| Paciente duplicado na planilha | Variação de grafia ou duplicidade já existente | Revise e padronize a lista antes de executar |
| Erro ao salvar o arquivo | Planilha aberta em outro programa ou falta de permissão | Feche o Excel e verifique as permissões do arquivo |
| Chromium não abre | Navegador não instalado ou dependências ausentes | Execute `playwright install chromium` |

### Reprocessando um único dia

Se apenas o dia 15 apresentar falha, por exemplo, configure:

```python
DIA_INICIO = 15
DIA_FIM = 15
```

Execute novamente e confira o resultado no sistema e na planilha.

> **Atenção:** antes de reprocessar uma data, verifique como o script trata marcações existentes. Ele deve evitar duplicar `X` ou `XX` em execuções repetidas.

---

## 🔄 Adaptação para outros sistemas

O script foi desenvolvido para uma estrutura específica de sistema web. Para adaptá-lo a outra plataforma, os principais pontos a revisar são:

| Componente | O que adaptar |
|---|---|
| `abrir_tela` | Links do menu, navegação e seletor do botão de busca (`#btnBuscar`) |
| `buscar_dia` | Calendário, seleção da data, botão de busca e título da tabela |
| `JS_LER_TABELA` | Identificação da tabela e das colunas de paciente e status |
| `eh_atendido` | Regras para identificar os status que representam atendimentos realizados |

### Dica de desenvolvimento

O Playwright oferece uma ferramenta para gravar interações com o navegador e auxiliar na identificação dos seletores:

```bash
playwright codegen https://endereco-do-seu-sistema
```

Utilize um ambiente autorizado e evite gravar ou compartilhar sessões, dados pessoais, tokens ou informações confidenciais durante a inspeção.

---

## 🔒 Privacidade e segurança

Como o projeto pode envolver informações relacionadas à saúde, a proteção dos dados deve fazer parte de todo o fluxo de utilização.

### Boas práticas

- **Não publique dados reais:** não envie planilhas de pacientes, relatórios identificáveis ou capturas de tela para repositórios públicos.
- **Proteja as credenciais:** faça login manualmente e nunca inclua senhas, cookies ou tokens no código.
- **Evite dados pessoais nos logs:** registre contagens, datas e mensagens técnicas, sem nomes ou informações clínicas.
- **Proteja os arquivos locais:** restrinja o acesso às planilhas e às cópias de segurança.
- **Utilize somente acessos autorizados:** respeite as regras da clínica, as condições de uso do sistema e as orientações profissionais aplicáveis.
- **Valide os resultados:** confira manualmente os dias que apresentarem falha antes de considerar o controle concluído.
- **Avalie a autorização para automação:** confirme com a clínica ou responsável pelo sistema se esse tipo de acesso automatizado é permitido.

### Arquivo `.gitignore`

Inclua regras para evitar o envio acidental de arquivos sensíveis ao GitHub.

```gitignore
# Ambientes virtuais
.venv/
venv/

# Configurações locais e segredos
.env
.env.*
!.env.example

# Planilhas e dados exportados
*.xlsx
*.xls
*.csv

# Capturas de tela
screenshots/
screenshots_*/
*.png
*.jpg
*.jpeg

# Cache do Python
__pycache__/
*.py[cod]
```

**Importante:** revise as regras de acordo com os arquivos do projeto. O `.gitignore` impede novos rastreamentos de arquivos não versionados, mas não remove arquivos que já foram adicionados ao Git.

Se um arquivo com dados reais for enviado por engano, removê-lo em um commit posterior não elimina o conteúdo do histórico. É necessário tratar a exposição, remover os dados do histórico quando apropriado e avaliar as medidas de resposta ao incidente com os responsáveis pela proteção de dados.

### LGPD e sigilo profissional

O uso do script deve respeitar a Lei Geral de Proteção de Dados Pessoais (LGPD), as políticas da instituição e as regras de sigilo profissional aplicáveis.

A automação não substitui a responsabilidade de conferir a integridade e a exatidão dos registros.

---

## ⚠️ Limitações

- Depende da estrutura e dos seletores atuais do sistema web.
- Pode precisar de ajustes após atualizações da plataforma.
- Processa uma especialidade por execução.
- Registra presenças, mas não classifica automaticamente faltas, cancelamentos ou valores.
- Requer interface gráfica, pois o Chromium é aberto de forma visível.
- Pode precisar de conferência manual quando há falhas de carregamento ou inconsistências nos dados.
- A correspondência de nomes depende da qualidade e da padronização das informações.
- O salvamento por dia reduz o risco de perder todo o progresso, mas não substitui backups e validações.

---

**Desenvolvido para reduzir tarefas repetitivas, organizar o controle mensal de atendimentos e facilitar a conferência de presença.**

*Automação com responsabilidade, segurança e atenção ao sigilo profissional.*
