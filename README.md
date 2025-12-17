# Listar processos SEI e exportar para Excel

Mini-projeto para listar **todos** os processos do SEI (MG) na unidade configurada e gerar um arquivo `.xlsx` com metadados.

## O que este repositório faz

- Faz login no SEI com `SEI_USER`/`SEI_PASS`
- Garante a unidade configurada em `SEI_UNIDADE` (tenta trocar automaticamente após o login)
- Lista processos **Recebidos** e **Gerados** (com paginação automática)
- Exporta um Excel em `./saida/processos.xlsx` (ou no caminho passado em `--saida`)

## O que ele não faz (por design)

- Não aplica filtros (sempre exporta tudo de Recebidos + Gerados)
- Não gera PDF
- Não abre cada processo para enriquecer dados “internos” (o Excel é baseado no que aparece no Controle de Processos)

## Pré-requisitos

- Python >= 3.13
- (Opcional) `uv` instalado: https://github.com/astral-sh/uv

## Como usar

### Opção A: com `uv` (recomendado)

1. Instalar dependências:
   ```bash
   uv sync
   ```

2. Configurar credenciais e unidade (via `.env`):
   ```bash
   cp .env.example .env
   ```
   Preencha no `.env` as variáveis obrigatórias:
   - `SEI_USER`
   - `SEI_PASS`
   - `SEI_ORGAO`
   - `SEI_UNIDADE`

3. Rodar a listagem e exportar Excel:
   ```bash
   uv run listar_processos_sei.py
   ```
   Saída padrão: `./saida/processos.xlsx`

Opcional (alterar caminho do Excel):
```bash
uv run listar_processos_sei.py --saida ./saida/processos.xlsx
```

### Opção B: sem `uv` (pip + venv)

1. Criar e ativar um virtualenv:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # macOS/Linux
   ```
   No Windows (PowerShell):
   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Instalar dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Configurar credenciais:
   ```bash
   cp .env.example .env
   ```

4. Rodar:
   ```bash
   python3 listar_processos_sei.py
   ```
   No Windows (PowerShell):
   ```powershell
   py listar_processos_sei.py
   ```

## Saída (colunas do Excel)

O arquivo `.xlsx` contém 1 linha por processo e colunas (nomes internos):

- `numero_processo`, `categoria`, `visualizado`
- `titulo`, `tipo_especificidade`
- `responsavel_nome`, `responsavel_cpf`
- `marcadores`
- `tem_documentos_novos`, `tem_anotacoes`
- `id_procedimento`, `hash`, `url`

## Variáveis de ambiente

Obrigatórias:

- `SEI_USER` / `SEI_PASS`
- `SEI_ORGAO` (código do órgão)
- `SEI_UNIDADE` (nome da unidade no SEI)

Opcionais:

- `SEI_DEBUG=1` habilita logs detalhados
- `SEI_SAVE_DEBUG_HTML=1` salva HTMLs úteis para depuração em `data/debug/`
- `SEI_DATA_DIR=data` troca o diretório base dos artefatos locais (debug HTML)

## Debug

- `SEI_DEBUG=1` habilita logs detalhados
- `SEI_SAVE_DEBUG_HTML=1` salva HTMLs em `data/debug/` (útil se o SEI mudar o layout)

## Troubleshooting rápido

- Erro de login/credenciais: confira `SEI_USER`/`SEI_PASS` e se sua conta não está bloqueada.
- Unidade não encontrada: confirme se `SEI_UNIDADE` está exatamente como aparece na lista do SEI (o script normaliza maiúsculas/espaços, mas precisa do mesmo texto).
- Mudança de layout do SEI: habilite `SEI_SAVE_DEBUG_HTML=1` e inspecione os HTMLs em `data/debug/`.
  - Se a extração parar de funcionar, o que geralmente muda são IDs/classes das tabelas ou o formulário de paginação.

## Segurança

- Nunca commite o `.env` (este repositório já ignora por padrão via `.gitignore`).
