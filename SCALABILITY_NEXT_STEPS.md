# SAE — Plano de Escalabilidade (Fase seguinte)

## Leitura inicial concluída
Arquivos-base do MVP analisados: `Index.html`, `code.gs` e `Readme.md`.

## Direção arquitetural (ERP/SaaS em GAS + Sheets)
1. **Camada de domínio no backend (`code.gs`)**
   - Separar regras por módulos lógicos: `ClientesService`, `OperacoesService`, `DashboardService`.
   - Padronizar contratos de entrada/saída (DTOs) para `google.script.run`.
2. **Padronização de dados no Sheets**
   - Garantir tipagem consistente (número, ISO date, UUID) em toda gravação.
   - Introduzir aba `auditoria` para trilha mínima de alterações críticas.
3. **Frontend Vue 3 (single-file)**
   - Evoluir para componentes por template string interna (sem bundler) para reduzir acoplamento.
   - Inserir estados de erro por módulo (dashboard/clientes/operações).
4. **Confiabilidade operacional**
   - Validar pré-condições de infraestrutura (abas existentes, headers obrigatórios).
   - Criar rotina de health-check backend para execução antes de releases.

## Critérios para considerar “pronto para produção inicial”
- Setup idempotente da base com headers obrigatórios.
- Todas as entradas numéricas normalizadas no backend.
- Erros de domínio retornados de forma amigável no frontend.
- Dashboard sem quebra em caso de base vazia.

## Próxima entrega sugerida
Implementar camada de serviços no `code.gs` e adicionar health-check automático acionado ao abrir a aplicação.
