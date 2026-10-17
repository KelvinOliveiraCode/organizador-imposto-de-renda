# Organizador de Declaração de Imposto de Renda

![Excel](https://img.shields.io/badge/Excel-217346?logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)

Planilha em Excel para organizar as informações da declaração de Imposto de Renda de Pessoa Física (IRPF), com abas de rendimentos, deduções, bens e direitos, validações de preenchimento e consolidação automática na aba Resumo.

## Problema / Por que existe

Reunir os dados da declaração de IRPF exige reunir informações espalhadas entre comprovantes de rendimento, recibos de despesa dedutível e extratos de patrimônio. Sem um único lugar para lançar esses valores, fica difícil conferir totais e evitar erros de digitação antes de preencher o programa oficial da Receita Federal. Este projeto cria esse ponto único de lançamento, construído inteiramente no Excel e regenerável por um script Python.

## O que faz

O arquivo `Organizador_IRPF.xlsx` contém sete abas:

- **Início**: capa com botões de navegação (hiperlinks) para todas as abas e instruções de uso.
- **Dados Pessoais**: campos de cadastro do contribuinte — nome completo, CPF, data de nascimento, estado civil, endereço, cidade/UF, CEP, telefone, e-mail, tipo de declaração, exercício e código de agendamento — com valores de exemplo como lugar de preenchimento.
- **Rendimentos**: tabela de 40 linhas para lançar fontes pagadoras (nome, CNPJ/CPF), tipo de rendimento por lista suspensa (Tributável - PJ, Tributável - PF, Isento/Não Tributável, Rendimento de Aplicações, Outros), valor em reais, data de recebimento e observação. Totaliza automaticamente os valores na última linha e destaca células acima de R$ 10.000,00 por formatação condicional.
- **Deduções**: tabela de 30 linhas com categoria por lista suspensa (Previdência Oficial, Saúde, Mobilidade/Educação, Dependente, Pensão Alimentícia, Doações, Outras), descrição, valor, documento e observação. Total automático no fim da tabela.
- **Bens e Direitos**: tabela de 25 linhas com grupo de bem por lista suspensa (Imóveis, Veículos, Aplicações Financeiras, Participação em Empresas, Bens Móveis, Outros), descrição, situação em 31/12 e observação. Total automático no fim da tabela.
- **Resumo**: consolidação automática por fórmulas — total de rendimentos, total de deduções, total de bens, base de cálculo (rendimentos menos deduções), alíquota estimada de referência (15%) e imposto estimado (cálculo genérico e educativo, não substituto da legislação vigente).
- **Links Úteis**: links diretos para sites oficiais da Receita Federal (IRPF, download do programa, consulta à restituição, e-CAC, tabelas práticas) e materiais de apoio do desafio DIO/GitHub.

Cada aba tem uma barra de navegação no topo com hiperlinks para as outras abas, e as tabelas de lançamento têm cabeçalhos congelados.

## Stack

- Microsoft Excel (ou LibreOffice Calc) — arquivo `.xlsx` entregável, com abas, fórmulas, validações de dados, formatação condicional e hiperlinks.
- Python 3 com openpyxl — script `gerar_planilha.py` que recria a planilha do zero (estilo preto, branco e cinzo, validações, fórmulas, formatação condicional e links).

## Como funciona

O script `gerar_planilha.py` cria cada aba com `openpyxl`, define largura de colunas, desliga grades visíveis, insere a barra de navegação com hiperlinks internos (`#'Nome da Aba'!A1`), acrescenta listas suspensas com `Data Validation` nos campos de tipo de rendimento, categoria de dedução e grupo de bem, escreve fórmulas de soma (`=SUM(...)`) e de base de cálculo e imposto estimado na aba Resumo, aplica formatação condicional (`CellIsRule`, maior que 10000) na coluna de valores de rendimentos, congela cabeçalhos (`freeze_panes`) e salva `Organizador_IRPF.xlsx`.

Para usar, abra o xlsx, preencha as abas de lançamento usando as listas suspensas quando disponíveis e confira os totais na aba Resumo.

## Como rodar

1. Instale o openpyxl:

```bash
pip install openpyxl
```

2. Execute o gerador na raiz do projeto:

```bash
python gerar_planilha.py
```

O script sobrescreve `Organizador_IRPF.xlsx` no diretório atual. Abra o arquivo no Excel para preencher a declaração.

## Estrutura do projeto

```
Organizador de Declaração de Imposto de Renda/
├── Organizador_IRPF.xlsx   # Planilha principal (entregável, gerada pelo script)
├── gerar_planilha.py       # Script Python que gera a planilha com openpyxl
├── README.md               # Este arquivo
└── contexto do desafio.txt # Enunciado do desafio DIO que originou o projeto
```

## Trade-offs e decisões

- **Tudo no Excel, gerado por Python**: a planilha é o entregável final e o script Python serve como documentação executável de como recriá-la. Isso evita depender de macros (VBA), que podem ser bloqueadas por configurações de segurança do Excel, mas significa que qualquer edição manual no xlsx precisa ser refeita no script para ser reproduzida.
- **Tema preto, branco e cinzo**: substitui a paleta de referência por uma interface neutra e de alto contraste, sem depender de cores de marca.
- **Imposto estimado simplificado**: a aba Resumo aplica uma alíquota de referência fixa (15%) sobre a base de cálculo. É um exemplo genérico e educativo; não reflete alíquotas progressivas, deduções legais específicas nem isenções da Receita Federal.
- **Quantidade fixa de linhas**: 40 linhas em rendimentos, 30 em deduções e 25 em bens e direitos. Cobrem a maioria dos casos, mas um contribuinte com muitos lançamentos precisa incluir linhas além do intervalo coberto pelas fórmulas de total.

## Limitações conhecidas

- O cálculo do imposto na aba Resumo é genérico e educativo; consulte sempre a legislação vigente e um profissional contábil antes de declarar.
- Validações de lista suspensa e formatação condicional podem se comportar de forma diferente em editores que não o Excel, como o LibreOffice Calc.
- Não há importação automática de extratos, notas fiscais ou do programa oficial da Receita Federal; todo lançamento é manual.
- O script não valida CPF, CNPJ nem datas; apenas restringe a digitação às listas suspensas onde aplicável.

## Licença

Este projeto é um entregável de estudo do desafio DIO "Construindo um Organizador de Declaração de Imposto de Renda no Excel". Não há licença de software definida; use e adapte por conta própria.

## Autor

Kelvin
