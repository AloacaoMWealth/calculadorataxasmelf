# Motor de Fee M Wealth

## Rodar

```bash
pip install -r requirements_fee_mwealth.txt
streamlit run app_fee_mwealth.py
```

## Fluxo
1. Suba o **Controle de Clientes MWealth**.
2. Escolha o mês/ano da cobrança.
3. O app pega automaticamente o último `PL dd/mm/aaaa` anterior ao mês selecionado.
4. Soma o PL por **Grupo Familiar** (offshore convertido pela PTAX) e encontra a faixa da **Tabela Fee** de cada conta.
5. Suba os PLs:
   - XP: um relatório com todos os dias;
   - BTG: um arquivo `PL Total DD.MM.xlsx` por dia útil;
   - Charles Schwab: um `CS Total DD.MM.csv` por dia útil;
   - Safra / XP US: por enquanto, formato mensal padronizado com `Conta` e `Valor` (opcionais: `Data`, `Moeda`).
6. Clique em **Calcular cobrança**.

## Fórmulas
- Diário: `Fee = PL_BRL × Fee a.a. / 252`
- Mensal: `Fee = PL_BRL × Fee a.a. / 12`
- Offshore: `PL_BRL = PL_USD × PTAX venda`

## Saídas
- Cadastro Fee
- Grupos
- Resumo Cobrança
- Memória de Cálculo
- Uma aba por corretora
- Pendências de match
- CSV separado por corretora

## Regras importantes
- `Sem cobrança` vira taxa 0%.
- Labels `Tesouraria (0,35)` e `Tesouraria (0,40%)` são interpretadas como taxa fixa anual.
- Contas são normalizadas sem máscara/zeros à esquerda para casar BTG e Controle.
- Dias faltantes **não são preenchidos automaticamente**: a tela mostra cobertura dos arquivos diários.
