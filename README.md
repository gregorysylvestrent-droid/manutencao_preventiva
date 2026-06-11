# Dashboard de Manutenção Preventiva — Apps Script

Projeto em Google Apps Script para acompanhar a saúde das manutenções preventivas de veículos a partir das bases da planilha **manutenção preventiva**. O painel foi desenvolvido para substituir fórmulas auxiliares na planilha: o usuário cola/atualiza as bases e o processamento é feito pelo backend `.gs`, refletindo os resultados no dashboard HTML.

## Objetivo

Centralizar, em uma interface web com identidade visual NorteTech, a análise de vencimentos e aderência das manutenções preventivas por veículo, item, modelo, placa e centro de custo.

O dashboard permite visualizar:

- situação geral da frota;
- veículos atrasados, urgentes, próximos do vencimento e em dia;
- quantidade de preventivas que deveriam ter sido realizadas;
- quantidade de preventivas realizadas e pendentes;
- acompanhamento detalhado por item de manutenção;
- resumo consolidado por veículo;
- evolução mensal do histórico de manutenção;
- veículos sem plano cadastrado por modelo;
- distribuição de vencimentos por período/mês;
- filtros múltiplos por placa, centro de custo, modelo e status.

## Arquivos do projeto

| Arquivo | Finalidade |
| --- | --- |
| `Código.js` | Backend do Google Apps Script. Lê as abas, normaliza dados, calcula visões derivadas, aplica filtros e fornece dados ao dashboard. |
| `Dashboard.html` | Interface do painel, contendo layout, estilos, filtros, modo claro/escuro, tabelas, gráficos e chamadas ao backend via `google.script.run`. |
| `README.md` | Documentação do projeto, objetivo, estrutura esperada e instruções de implantação. |

> No editor do Apps Script, o arquivo `Código.js` deve ser criado como arquivo `.gs` ou colado em um arquivo de script do projeto. O `Dashboard.html` deve ser criado como arquivo HTML chamado `Dashboard`.

## Abas esperadas na planilha

O projeto foi preparado para ler os nomes originais da planilha e também algumas variações com emojis usadas em versões anteriores.

| Aba | Descrição |
| --- | --- |
| `Cadastro` | Base de veículos, placas, modelos, centro de custo e tipo de veículo. |
| `Plano Mnt` | Plano preventivo por modelo, com os itens de manutenção e intervalos previstos. |
| `Histórico MNTPREV` | Histórico de manutenções realizadas por placa, modelo, item, data e quilometragem. |
| `Registro de Abastecimentos` | Base usada para identificar a quilometragem atual/mais recente de cada veículo. |
| `Acompanhamento por Item` | Visão detalhada por item. Pode existir como aba pronta, mas o backend também consegue calculá-la quando ausente ou vazia. |
| `Resumo por Veículo` | Visão consolidada por veículo. Pode existir como aba pronta, mas o backend também consegue calculá-la quando ausente ou vazia. |
| `Meses Vencimento` | Visão resumida de vencimentos por período/mês, quando disponível. |
| `Sem Plano` | Relação de veículos sem plano cadastrado. Pode existir como aba pronta, mas o backend também consegue calculá-la quando ausente ou vazia. |

## Funcionamento dos cálculos

O backend prioriza as bases coladas pelo usuário e reduz a dependência de fórmulas na planilha:

1. Lê `Cadastro`, `Plano Mnt`, `Histórico MNTPREV` e `Registro de Abastecimentos`.
2. Identifica o plano aplicável a cada veículo pelo modelo.
3. Obtém a quilometragem atual pela base de abastecimentos mais recente.
4. Busca a última manutenção realizada por placa e item no histórico.
5. Calcula, por item:
   - KM atual;
   - KM da última manutenção;
   - KM desde a última manutenção;
   - próxima manutenção prevista em KM;
   - KM restante para a próxima manutenção;
   - quantidade que deveria ter sido feita;
   - quantidade realizada;
   - quantidade pendente;
   - status do item.
6. Consolida por veículo:
   - item mais crítico;
   - total de itens monitorados;
   - itens atrasados, urgentes, próximos e OK;
   - total esperado, realizado e pendente;
   - status geral do veículo;
   - mês/período de vencimento.
7. Lista veículos cujo modelo não possui plano cadastrado.

Quando as abas `Acompanhamento por Item`, `Resumo por Veículo` ou `Sem Plano` já existem e possuem dados, o projeto pode reaproveitá-las. Quando não existem ou estão vazias, essas visões são calculadas automaticamente pelo código.

## Status considerados

| Status | Interpretação |
| --- | --- |
| `ATRASADO` | Item/veículo com manutenção vencida ou quantidade pendente. |
| `URGENTE` | Item/veículo muito próximo do vencimento. |
| `PRÓXIMO` | Item/veículo dentro da faixa de atenção para vencimento futuro. |
| `OK` | Item/veículo dentro do prazo. |
| `SEM DADO` | Não foi possível determinar o status por falta de dados suficientes. |

## Recursos do dashboard

- Indicadores principais de frota, atraso, urgência, aderência e pendências.
- Tabela de veículos críticos.
- Tabela paginada de `Resumo por Veículo`.
- Tabela paginada de `Acompanhamento por Item`.
- Visão de histórico mensal de manutenções.
- Visão de veículos sem plano.
- Visão de vencimentos por mês/período.
- Modal de detalhe por veículo.
- Filtros múltiplos por:
  - placa;
  - centro de custo;
  - modelo;
  - status.
- Alternância entre modo claro e modo escuro.
- Layout responsivo para uso em navegador.

## Como implantar no Google Apps Script

1. Abra a planilha de manutenção preventiva no Google Sheets.
2. Acesse **Extensões > Apps Script**.
3. Crie/cole o backend:
   - crie um arquivo de script;
   - cole o conteúdo de `Código.js`;
   - se desejar, nomeie o arquivo como `Código.gs` no Apps Script.
4. Crie o arquivo HTML:
   - clique em **+ > HTML**;
   - nomeie o arquivo como `Dashboard`;
   - cole o conteúdo de `Dashboard.html`.
5. Salve o projeto.
6. Execute/autorize uma função do backend, se solicitado pelo Google, para conceder permissões de leitura da planilha.
7. Publique como app da web:
   - **Implantar > Nova implantação**;
   - selecione **App da Web**;
   - execute como o proprietário/usuário adequado;
   - defina quem pode acessar conforme a política da organização;
   - copie a URL gerada.
8. Acesse a URL do app para visualizar o painel.

## Uso recomendado

1. Atualize/cole as bases nas abas de origem, especialmente:
   - `Cadastro`;
   - `Plano Mnt`;
   - `Histórico MNTPREV`;
   - `Registro de Abastecimentos`.
2. Mantenha os cabeçalhos das colunas com nomes claros e consistentes.
3. Abra ou atualize o dashboard.
4. Use os filtros múltiplos para restringir a análise por placa, centro de custo, modelo ou status.
5. Consulte as páginas/tabelas do painel para identificar prioridades de manutenção.

## Observações sobre qualidade dos dados

Para melhores resultados, mantenha as bases com:

- placas padronizadas;
- modelos escritos de forma idêntica entre `Cadastro` e `Plano Mnt`;
- datas válidas no histórico e abastecimentos;
- quilometragens numéricas;
- itens de manutenção com nomes consistentes entre plano e histórico;
- intervalos de manutenção preenchidos no plano.

Se um veículo aparece em `Sem Plano`, normalmente significa que o modelo cadastrado para ele não foi encontrado no `Plano Mnt`.

## Validações locais realizadas

Durante o desenvolvimento foram utilizadas verificações de sintaxe e um teste local com mock de `SpreadsheetApp` para validar a lógica principal do backend:

```bash
node --check 'Código.js'
```

```bash
# Extração do script inline do Dashboard.html e validação com node --check
```

```bash
node /tmp/test_backend.js
```

Esses testes validam sintaxe e comportamento básico da lógica em ambiente local. A validação final deve ser feita também no Google Apps Script conectado à planilha real, pois permissões, nomes exatos de abas/colunas e dados reais podem variar.

## Limitações conhecidas

- O projeto depende de dados consistentes entre cadastro, plano, histórico e abastecimentos.
- A estimativa de vencimento por mês é uma aproximação baseada na quilometragem restante quando não há aba `Meses Vencimento` pronta.
- Caso a estrutura real da planilha tenha colunas muito diferentes dos nomes esperados, pode ser necessário ajustar os mapeamentos de leitura no backend.
- O painel usa `google.script.run`, portanto a execução completa depende do ambiente do Google Apps Script.
