SAE - Sistema Apollo Enterprise: Gestão de Investimentos

Documentação Técnica de Arquitetura e Protocolos
Versão: 2.0 (Fase de Escalonamento)
Responsável: Anderson

1. Visão Geral do Sistema

O SAE Mini Índice é um ERP desenvolvido sobre a infraestrutura do Google Cloud (Google Apps Script) para gestão de carteira de investidores e monitoramento de operações financeiras. O sistema utiliza uma arquitetura Single-File WebApp com persistência de dados em Google Sheets, seguindo os padrões de design Glassmorphism SAE.

2. Arquitetura de Dados (Database)

O banco de dados reside em uma planilha Google, onde cada aba representa uma entidade normalizada.

2.1. Entidade: clientes

Armazena o cadastro mestre dos investidores.

Campo

Tipo

Descrição

Regra de Negócio

uuid

String

Identificador Único Universal

Gerado via Utilities.getUuid().

id_sequencial

Integer

ID de Controle Visual

Incremental automático (Ex: 1, 2, 3...).

data_cadastro

ISO8601

Data de entrada no sistema

Automação: new Date().toISOString().

status

String

Estado do registro

Padrão: "Ativo". Opções: "Ativo", "Inativo".

nome

String

Nome completo

Campo obrigatório.

idade

Integer

Idade do investidor

Utilizado para análise de perfil.

telefone

String

Contato Primário

Chave Única. Bloqueia duplicidade no backend.

email

String

Contato Secundário

Chave Única. Bloqueia duplicidade no backend.

cidade

String

Município de residência

Exibição mesclada com Estado no Frontend.

estado

String (2)

Unidade Federativa

Armazenado em uppercase (Sigla).

redes_sociais

String

Handlers sociais

Ex: @usuario.

valor_mensalidade

Number

Valor de faturamento

Base para o módulo de faturamento (v3.0).

vencimento_dia

Integer

Dia para cobrança

Inteiro de 1 a 31.

capital_inicial

Number

Margem por contrato

Valor base para cálculo de % de ganho.

2.2. Entidade: operacoes

Registro histórico de todas as entradas no mercado.

Relacionamento: N:1 com a tabela clientes via cliente_id.

Cálculos Automáticos: - Lucro Bruto = (Pontos+ - Pontos-) * Contratos * Valor_Ponto

Taxas = Contratos * 0.25

Lucro Líquido = Lucro Bruto - Taxas

3. Regras de Negócio e Segurança (Backend)

3.1. Validação de Integridade

O backend (Código.js) implementa uma camada de segurança antes do appendRow:

Verificação de Existência: Antes de inserir, o sistema varre as colunas telefone e email.

Tratamento de Erros: Se houver colisão, o GAS dispara um Error que é capturado pelo withFailureHandler no Vue 3, exibindo uma mensagem amigável ao usuário e impedindo o registro duplo.

3.2. Formatação PT-BR

Datas: Todas as datas são armazenadas em ISO (YYYY-MM-DDTHH:mm) para garantir ordenação correta, mas são convertidas para toLocaleDateString('pt-BR') na renderização.

Moeda: Valores financeiros seguem o padrão BRL (R$ 0,00).

4. Frontend (UI/UX Protocol)

4.1. Componentes Vue 3

Modal de Cadastro: Utiliza position: fixed com backdrop-filter para foco total na entrada de dados.

Tabela Dinâmica: Implementa filtros em tempo real usando Computed Properties do Vue, garantindo performance mesmo com centenas de clientes.

Mesclagem Visual: No front, cidade e estado são concatenados ({{c.cidade}} / {{c.estado}}) para "economia de espaço", conforme diretriz SAE.

4.2. Estados da Interface

Loading State: Ativado globalmente via variável loading durante qualquer comunicação com o google.script.run.

Empty State: Verificação de clientes.length para exibição de mensagens de "Nenhum investidor encontrado".

5. Próximos Passos (Roadmap)

[ ] Módulo Faturamento: Cruzamento entre vencimento_dia e data atual para gerar alertas de pendência.

[ ] Interligação de Saldos: Atualização automática da aba saldos após cada registro em operacoes.

[ ] Gráficos Avançados: Implementação de evolução patrimonial por investidor.

Documento gerado automaticamente pelo Sistema SAE. Proibida alteração de chaves primárias sem aviso prévio ao arquiteto do sistema.
