# SAE - Sistema Apollo Enterprise para Gestão de Investimentos

## Visão geral

Este projeto é um WebApp em Google Apps Script (GAS) para gestão pessoal de clientes e operações de índice/investimento. O SAE utiliza Google Sheets como banco de dados relacional simples, Vue 3 via CDN no frontend e `google.script.run` como ponte segura entre interface e backend.

## Stack obrigatória

- **Frontend:** Vue 3 CDN, HTML single-file, CSS utilitário, mobile-first, Chart.js CDN.
- **Backend:** Google Apps Script V8, HtmlService e funções públicas para `google.script.run`.
- **Banco:** Google Sheets, uma aba por entidade, UUID sistêmico, datas ISO e números normalizados no backend.
- **Sem:** React, JSX, bundlers, Babel runtime ou backend externo.

## Entidades do banco

### `tbl_clientes`

| Campo | Função |
| --- | --- |
| `uuid` | Chave sistêmica oculta ao usuário. |
| `cliente_id` | ID manual lançado pelo usuário e exibido na interface. |
| `data_cadastro` | Data ISO de cadastro. |
| `status` | Ativo, Inativo ou Potencial. |
| `nome`, `idade`, `telefone`, `email`, `cidade`, `estado`, `redes_sociais` | Dados cadastrais. |
| `valor_mensalidade`, `vencimento_dia`, `capital_inicial_contrato` | Dados financeiros/contratuais. |
| `pagamento_valor`, `data_pagamento`, `data_desligamento` | Dados de acompanhamento administrativo. |

### `tbl_operacoes`

| Campo | Função |
| --- | --- |
| `uuid` | Chave sistêmica da operação. |
| `cliente_id` | Relação com `tbl_clientes.cliente_id`. |
| `data_operacao` | Data ISO da operação. |
| `capital_inicial_contrato` | Valor herdado do cliente, editável no lançamento. |
| `n_contratos` | Quantidade de contratos. |
| `valor_por_contrato` | Valor por ponto/contrato, padrão `0.20`. |
| `pontos_pos`, `pontos_neg` | Pontos positivos e negativos. |
| `percentual_ganho` | Calculado no backend. |
| `take` | `pontos_pos * n_contratos * valor_por_contrato`. |
| `stop` | `pontos_neg * n_contratos * valor_por_contrato`. |

### `config`

Armazena metas históricas. A meta inicial padrão é **750 pontos**. Novas metas são append-only para não reescrever dados já lançados.

### `auditoria`

Registra eventos sistêmicos críticos em JSON: criação/edição de clientes, operações e metas.

## Regras principais

1. O backend executa normalização de números, datas, duplicidade e cálculos.
2. O frontend apenas coleta dados, formata para PT-BR e apresenta gráficos/tabelas.
3. Cliente não pode duplicar por `cliente_id`, telefone ou e-mail.
4. Operação sempre se relaciona por `cliente_id` manual do cliente.
5. Take e Stop são valores em BRL derivados dos pontos lançados.
6. A carteira é derivada de operações, sem duplicar saldo em nova aba.

## Fluxo de implantação no GAS

1. Criar projeto Google Apps Script vinculado à planilha.
2. Criar/colar os arquivos `code.gs` e `Index.html`.
3. Executar `setupDatabase()` pelo editor GAS para criar/sincronizar abas.
4. Publicar como WebApp.
5. Usar o botão **Check** no header antes de releases.
