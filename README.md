# Listar processos SEI e exportar para Excel

Mini-projeto para listar **todos** os processos do SEI (MG) na unidade configurada e gerar um arquivo `.xlsx` com metadados.

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

## Debug

- `SEI_DEBUG=1` habilita logs detalhados
- `SEI_SAVE_DEBUG_HTML=1` salva HTMLs em `data/debug/` (útil se o SEI mudar o layout)
