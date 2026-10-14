# Relatório Gerencial — CGO × Filial

Aplicativo Python/Tkinter que cruza dois relatórios CSV por código de produto e gera um Excel com movimentos por filial e uma aba de auditoria para produtos vinculados a múltiplos CGOs.

## Executar

Requer Python 3.10 ou superior com Tkinter.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe Gerencial_itens.py
```

Selecione **CGO SINT**, depois **FILIAL SINT**, e clique em **Gerar Excel**. Os CSVs usam separador `;`, codificação UTF-8 ou Latin-1 e valores numéricos no padrão brasileiro, como `1.234,56`.

## Colunas de entrada

| CSV | Colunas obrigatórias |
| --- | --- |
| CGO SINT | `Código Produto Base`, `Código C.G.O.`, `C.G.O. : Produto Base`, `Qtd Saída Outras`, `Vlr Saída Outras` |
| FILIAL SINT | `Código Produto Base`, `Código Empresa`, `Qtd Saída Outras`, `Vlr Saída Outras` |

Produtos com um CGO seguem para as abas de filiais. Produtos com vários CGOs aparecem em `PRODUTOS CGOS DIFERENTES`. Essa aba repete o movimento da filial para cada CGO relacionado: não some suas linhas como se fossem movimentos independentes.

## Privacidade e limites

O processamento é local, sem conexão com banco, login ou credenciais. Este repositório contém somente código e documentação, sem exportações empresariais ou dados de clientes. Arquivos de entrada e relatórios gerados estão no `.gitignore`.

O layout precisa corresponder aos nomes de colunas esperados. Produtos sem correspondência entre os CSVs não aparecem no cruzamento. Valores que não podem ser convertidos são tratados como zero pelo código original; confira os arquivos antes de usar o resultado. Não há suíte automatizada no projeto original.
