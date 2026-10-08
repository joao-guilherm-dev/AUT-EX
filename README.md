# AUT-EX

Automação de planilha de vendas em Excel com Python. Você digita os produtos no terminal e o script cuida do resto: cria o arquivo, registra as vendas, calcula os totais e deixa tudo formatado.

## O que ele faz

- Cria `vendas.xlsx` na primeira execução, ou continua de onde parou se o arquivo já existe
- Registra vários produtos em sequência, sem precisar rodar o script de novo
- Calcula o total de cada venda com **fórmula do Excel** (`=B2*C2`), então a planilha continua calculando sozinha se você editar um valor
- Mostra o **total geral** e a quantidade de **itens vendidos** ao lado da tabela
- Cabeçalho estilizado, valores em reais (R$), filtro, linha de cabeçalho congelada e larguras ajustadas
- Valida as entradas: aceita `10,50`, `10.50` e `1.500,50`, e recusa texto, números negativos e quantidade quebrada
- Se o arquivo estiver aberto no Excel, avisa e espera você fechar em vez de dar erro

## Como usar

```bash
pip install -r requirements.txt
python aut.py
```

Exemplo:

```
Nova planilha criada!
Produto: Mouse
Quantidade: 2
Preço: 49,90
  OK: Mouse | 2 x R$ 49.90 = R$ 99.80 (linha 2)
Adicionar outro produto? [S/n] n
1 produto(s) salvo(s) em 'vendas.xlsx'.
```

## Estrutura da planilha

| Produto | Quantidade | Preço    | Total (fórmula) |
|---------|------------|----------|-----------------|
| Mouse   | 2          | R$ 49,90 | `=B2*C2`        |

Ao lado: **Total geral** (`=SUM(D:D)`) e **Itens vendidos** (`=SUM(B:B)`).

> Os totais são fórmulas, então os valores aparecem quando o arquivo é aberto no Excel (ou LibreOffice). Se for ler o `.xlsx` por outro programa (como o pandas), abra e salve uma vez no Excel antes.

## Testes

```bash
pip install -r requirements-dev.txt
pytest
```

## Próximos passos

- Exportar relatório em PDF ou gráfico de vendas por produto
- Ler vendas de um CSV em vez de digitar
- Interface gráfica (Tkinter)

## Licença

MIT
