# Capital BH Consórcios — Gestão Financeira (Desktop)

Aplicativo desktop em Python (CustomTkinter) para gestão financeira de uma empresa de consórcios.

## Como executar

Crie um ambiente virtual (recomendado) e instale as dependências:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Execute:

```bash
python main.py
```

## Arquivo mestre (Excel)

O app cria/usa o arquivo `Financeiro_Capital_BH.xlsx` na mesma pasta do projeto, com as abas:
- `Vendas`
- `Gastos`
- `Retiradas`

## Importação de CSV

Cada tela (Vendas/Gastos/Retiradas) permite importar CSV e o app tenta mapear automaticamente colunas comuns.
Se alguma coluna não for reconhecida, o registro ainda pode ser salvo manualmente.

