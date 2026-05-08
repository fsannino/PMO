# Plano de Modernização — Collab:Build

> Plano vivo. Cada sessão tem escopo fechado, prompt pronto para colar em uma nova
> conversa do Claude Code, e critérios de aceitação ("DoD") objetivos.

## Levantamento (sumário)

Sistema legado: **102 ASPs** (SMR) + **42 ASPs** (SMR_Light) + **37 forms VB6**
(Manutencao_SMR) + **8 forms VB6** (Manutencao_Light) + DLL VB6 de carga
+ scripts SQL Server 2000/2005 com **45 tabelas, 57 views, 153 procedures, 20+ functions**.

Seis frentes/módulos: **CUT** (CutOver), **GVI** (Governança/Saneamento),
**GRF** (Gráfico), **TCP** (Teste de Carga), **TIN** (Teste Integrado), **PMO** (master).

Carga de dados: 6 templates Excel — **IS** (Issues), **RK** (Risks), **AC** (Actions),
**CR** (Change Requests), **TK** (Tasks/cronograma), **TC** (Traceability) — via DLL
`CargaPlanilhasPMOnline.dll`. Integração MS Project via queries `.dqy` ODBC sobre
`MSP_Tasks`/`MSP_Resources`/`MSP_Links`.

Regras-chave: **Janela de HB** (`CONFIG_FECHAMENTO` por dia da semana) bloqueia
medições fora do horário; **Scorecard** com `Andamento`, `Atrasadas_0-10`,
`Atrasadas_+10`; **Curva S** com BCWP/ACWP; **Matriz de acesso** (`Access`)
controla quem vê qual projeto/módulo.

---

## Stack alvo

| Camada              | Escolha                                              |
| ------------------- | ---------------------------------------------------- |
| Frontend + Backend  | Next.js 15 (App Router) + React 19 + TypeScript      |
| Estilo              | Tailwind CSS + componentes próprios                  |
| ORM / DB            | Prisma + PostgreSQL (SQLite em dev)                  |
| Auth                | NextAuth (Credentials + JWT)                         |
| Validação           | Zod                                                  |
| Excel               | SheetJS (`xlsx`)                                     |
| MS Project          | `fast-xml-parser` (formato .xml exportado)           |
| Gantt / Charts      | SVG próprio (leve) + `recharts` para Curva S         |
| Testes              | Vitest (unit) + Playwright (e2e)                     |
| CI                  | GitHub Actions                                        |
| Empacotamento       | Docker + docker-compose (app + Postgres)             |

---

## Mapa de funcionalidades — legado → moderno

| Legado                                            | Moderno (módulo `pmo-modern`)             | Sessão |
| ------------------------------------------------- | ----------------------------------------- | ------ |
| `Login.asp`, `global.asa`                         | `/login`, NextAuth                        | 2      |
| `MenuPrincipal.asp`                               | `<AppShell>` + `/`                        | 2      |
| `Manutencao_SMR.exe` (cadastros)                  | `/admin/*`                                | 3      |
| `frmCadConfigFechamento`                          | `/admin/config-fechamento`                | 3      |
| `Tasks` + `Log_Tarefas` + `TarefasExcluidas_GVI`  | `Task` + `TaskHistory` (soft-delete)      | 4      |
| `CUT_Incluir_Tarefas`, `CUT_Alterar_Datas`        | `/projects/[id]/tasks` + Gantt edit       | 4      |
| `clsCargaPlanilhasPMOnline.dll` + 6 .xls          | `/projects/[id]/import` (Excel)           | 5      |
| Queries `.dqy` (MSP_Tasks)                        | `/projects/[id]/import` (MS Project XML)  | 5      |
| `Medicao_Filtro` + `Medicao_Detalhe`              | `/projects/[id]/measurement`              | 6      |
| `CONFIG_FECHAMENTO` + `VerificaHB()`              | middleware/serviço `closingWindow.ts`     | 6      |
| `SP_LISTAR_PAINEL_SINOPTICO_*` + views            | `/projects/[id]/synoptic` + `recharts`    | 7      |
| `DadosFlagCurvaS` + `RetornaBCWP`                 | serviço `curvaS.ts`                       | 7      |
| `GVI_Issues`, `Risks`, `Actions`, `Chang_Request` | `/projects/[id]/{issues,risks,actions,crs}` | 8    |
| `PMO_Relatorio_*` (Detalhado/Consolidado/Excel)   | `/projects/[id]/reports`                  | 9      |
| `Scorecard*`, `Issues_KPI`                        | dashboard + relatórios                    | 7, 9   |
| Módulos CUT/GVI/TCP/TIN/GRF                       | `Project.module` + UI específicas         | 10     |

---

## Sessões

### Sessão 0 — Fundação (parcial / em andamento)

**Status:** ⚠️ scaffold inicial commitado em `1d3e957`. A revisar após Sessão 1.

**Já feito:**
- `.gitignore`, `README.md` raiz.
- `pmo-modern/`: `package.json`, `tsconfig`, `next.config.mjs`, Tailwind, PostCSS, `.env.example`.
- `prisma/schema.prisma` preliminar (User/Project/Task/Measurement/Issue/Risk/Comment/ImportLog).
- `src/lib/db.ts`, `auth.ts`, `excel-import.ts` preliminares.

**Pendente para fechar a Sessão 0:**
- Layout raiz (`src/app/layout.tsx`), CSS global, página inicial vazia, healthcheck.
- `npm install` e build local validados.

### Sessão 1 — Modelo de domínio definitivo + seed

**Objetivo:** Prisma schema cobrindo **tudo de P0** (e estrutura para P1) com tipos
1:1 com o legado, mais seed com 1 projeto-exemplo e 5 usuários.

**Entidades a adicionar/refinar** (sobre o schema preliminar):
- `Action`, `ChangeRequest` + `ChangeRequestIncrement` + `ChangeRequestAlteration`.
- `Traceability` (Req↔Test↔Task).
- `Area`, `Frente`, `Governanca`, `Equipe`, `Unidade`, `Comite`.
- `Access` (matriz user↔project↔module).
- `Module` (`CUT|GVI|GRF|TCP|TIN|PMO|LIGHT`).
- `ConfigFechamento` (janelas por dia da semana).
- `Document` (anexos por projeto).
- `TaskHistory` (auditoria de tarefas, substitui `Log_Tarefas` + `TarefasExcluidas_GVI`).
- `MeasurementLock` (bloqueia confirmadas).
- `Scorecard` (snapshot diário; calculado, não persistido em escrita).

**DoD:**
- `npm run db:push` aplica sem erro.
- `npm run db:seed` cria: 5 usuários (admin/manager/member/viewer + light), 6 áreas, 1 projeto, 25 tarefas hierárquicas, 3 issues, 2 risks, 1 CR, janelas HB padrão.
- `npm run db:studio` abre e mostra dados.

**Prompt da sessão:**
```
Você está continuando o projeto pmo-modern em /home/user/PMO/pmo-modern (branch
claude/evaluate-repository-SO97E). Leia PLAN.md (Sessão 1) e o schema atual em
prisma/schema.prisma. Faça:

1. Refine prisma/schema.prisma adicionando: Action, ChangeRequest +
   ChangeRequestIncrement + ChangeRequestAlteration, Traceability, Area, Frente,
   Governanca, Equipe, Unidade, Comite, Access (user×project×module), Module
   (enum CUT/GVI/GRF/TCP/TIN/PMO/LIGHT), ConfigFechamento, Document, TaskHistory,
   MeasurementLock. Mantenha entidades existentes; ajuste relações.
2. Crie scripts/seed.ts com dados realistas (5 usuários, 6 áreas, 1 projeto com
   25 tarefas hierárquicas com predecessores, 3 issues, 2 risks, 1 CR, janelas
   HB padrão seg-sex 8-18h).
3. Rode `npm install`, `npm run db:push`, `npm run db:seed`. Resolva qualquer
   erro até o seed rodar limpo.
4. Commit "feat(db): modelo de domínio definitivo + seed".

DoD: prisma studio abre e mostra todas as entidades populadas.
```

### Sessão 2 — Auth + RBAC + Shell da aplicação

**Objetivo:** login funcional, layout da aplicação, middleware de autorização baseado em
`Role` e na matriz `Access`.

**Entregas:**
- `src/app/layout.tsx` (root) + `src/app/(app)/layout.tsx` (área autenticada com sidebar/topbar).
- `src/app/login/page.tsx` (server action + zod).
- `src/app/api/auth/[...nextauth]/route.ts`.
- `src/middleware.ts` — protege `/(app)/*`.
- `src/lib/access.ts` — `assertAccess(userId, projectId, module)`.
- `src/components/AppShell.tsx`, `Sidebar.tsx`, `UserMenu.tsx`.
- Página `/` (dashboard placeholder com lista de projetos do usuário).

**DoD:** login com seed `admin@pmo.local / admin123` funciona; usuário sem permissão
em projeto X recebe 403 ao acessar `/projects/X/*`.

**Prompt:**
```
Continuando pmo-modern. Sessão 2 (ver PLAN.md): implemente NextAuth credentials,
middleware protegendo /(app)/*, AppShell com sidebar (Projetos/Issues/Risks/CRs/
Relatórios/Admin), página /login (server action + zod) e /api/auth/[...nextauth].
Adicione src/lib/access.ts com assertAccess(userId, projectId, module) usando a
tabela Access. DoD: logar com admin@pmo.local/admin123, ver dashboard com lista
de projetos do usuário; acesso a projeto sem permissão retorna 403. Commit
"feat(auth): NextAuth + RBAC + AppShell".
```

### Sessão 3 — Manutenção / Admin (substitui Manutencao_SMR.exe)

**Objetivo:** CRUDs administrativos completos.

**Telas em `/admin/*`:**
- `users` — CRUD + reset de senha + ativar/desativar.
- `projects` — CRUD + ownership + módulo + datas + baseline.
- `areas`, `frentes`, `governancas`, `equipes`, `unidades`, `comites`.
- `access` — matriz visual (user×project×module).
- `config-fechamento` — janelas HB por dia da semana.
- `modules` — habilitação por projeto.
- `audit-log` — visualização de eventos.

**Padrões:** server actions + zod, formulário compartilhado (`<EntityForm/>`),
tabelas com filtro/ordenação/paginação.

**DoD:** admin consegue cadastrar projeto, criar usuário, conceder acesso,
configurar janela HB — tudo via UI sem tocar no banco.

**Prompt:**
```
Sessão 3 do pmo-modern (PLAN.md). Implemente CRUDs admin em /admin/*: users,
projects, areas, frentes, governancas, equipes, unidades, comites, access
(matriz), config-fechamento (janelas HB), modules. Use server actions + zod.
Crie src/components/ui/{DataTable, EntityForm, ConfirmDialog}. Acesso restrito
a Role=ADMIN via assertRole. DoD: admin cria projeto novo, cadastra 2 usuários,
abre acesso de cada um para o projeto, configura janela HB seg-sex 9-17 — tudo
pela UI. Commit "feat(admin): manutenção de dados mestre".
```

### Sessão 4 — Cronograma online (Tasks + Gantt)

**Objetivo:** CRUD completo de tarefas com hierarquia/dependências, edição inline,
visualização Gantt, histórico (substitui `CUT_Incluir_Tarefas`, `CUT_Alterar_Datas`,
`CUT_Excluir_Tarefas` + `Log_Tarefas`).

**Entregas:**
- `/projects/[id]/tasks` (lista hierárquica, drag-drop reorder, edição inline).
- `/projects/[id]/gantt` (SVG próprio, zoom dia/semana/mês, drag-resize, mostra
  predecessores como setas, milestones como losangos, baseline como barra de fundo).
- API REST: `GET/POST/PATCH/DELETE /api/projects/:id/tasks`,
  `POST /api/tasks/:id/dependencies`, `POST /api/tasks/:id/restore` (undo).
- Validação: dependência cíclica, datas coerentes, predecessor existe.
- `TaskHistory` gravado em todo update/delete.

**DoD:** criar projeto novo, adicionar 10 tarefas hierárquicas com predecessores,
visualizar Gantt, mover uma tarefa, ver entrada em histórico, restaurar tarefa
deletada.

**Prompt:**
```
Sessão 4 (PLAN.md). Implemente cronograma online: /projects/[id]/tasks (lista
hierárquica editável) e /projects/[id]/gantt (SVG próprio com zoom dia/semana/mês,
drag-resize de barras, setas de dependência, baseline em fundo, milestones).
APIs em src/app/api/projects/[id]/tasks/*. Toda mutação grava TaskHistory. Validar
ciclos com BFS. Soft-delete + restore. DoD: criar 10 tarefas hierárquicas com
3 dependências, mover datas no Gantt, deletar e restaurar uma tarefa, ver histórico
no /tasks/[taskId]/history. Commit "feat(schedule): tasks CRUD + Gantt + auditoria".
```

### Sessão 5 — Importação Excel + MS Project XML

**Objetivo:** substituir `clsCargaPlanilhasPMOnline.dll` e queries `.dqy`.

**Entregas:**
- `src/lib/excel-import.ts` — refinar para 6 templates (IS, RK, AC, CR, TK, TC),
  detecção automática de template pelo cabeçalho, mapeamento PT-BR/EN.
- `src/lib/msproject-import.ts` — parser `.xml` do MS Project (formato Project XML
  schema 2003+). Lê `Tasks`, `Predecessors`, `Resources`, `Assignments`.
- `/projects/[id]/import` (UI):
  - Upload (drag-drop) com escolha de modo: **substituir**, **mesclar**, **apenas-novos**.
  - Preview tabular antes de confirmar (primeiras 50 linhas + warnings).
  - Botão "Confirmar importação" → grava + cria `ImportLog`.
- `templates/` — gera/baixa templates .xlsx vazios para cada tipo.
- `/projects/[id]/import/history` — log de importações com download do warnings.txt.

**DoD:**
1. Importa `Manutencao_Light/Arq_pmo/TK.xls` (real do legado) e tarefas aparecem.
2. Importa export XML do MS Project (sample) — tarefas e dependências aparecem.
3. Preview mostra warnings (datas inválidas, predecessor inexistente).
4. Re-importação em modo "mesclar" não duplica.

**Prompt:**
```
Sessão 5 (PLAN.md). Implemente importação Excel (6 templates IS/RK/AC/CR/TK/TC,
detecção por header, PT-BR e EN) e MS Project XML (.xml exportado). UI em
/projects/[id]/import com drag-drop, preview de até 50 linhas + warnings, três
modos: substituir/mesclar/apenas-novos. Cria ImportLog. /projects/[id]/import/history
lista logs com download de warnings. /projects/[id]/import/templates baixa
.xlsx vazios. Teste com Manutencao_Light/Arq_pmo/TK.xls real. Commit
"feat(import): excel multi-template + MS Project XML".
```

### Sessão 6 — Medição + Janela de HB + Scorecard

**Objetivo:** substituir `Medicao_Filtro` + `Medicao_Detalhe` + `Confirmacao_Medicao`
+ `VerificaHB()` + `Scorecard`.

**Entregas:**
- `src/lib/closing-window.ts` — `isMeasurementOpen(now, dayOfWeek): boolean`
  consultando `ConfigFechamento`.
- `/projects/[id]/measurement` — lista de tarefas do usuário (filtros: equipe,
  período, status), entrada inline de % e horas trabalhadas, botão "Confirmar".
- Confirmação grava `Measurement` com `confirmed=true`, `confirmedAt`, e cria
  `MeasurementLock` impedindo nova medição daquela tarefa no período.
- `src/lib/scorecard.ts` — funções `andamento(projectId)`,
  `atrasadas10(projectId)`, `atrasadasMais10(projectId)`, equivalentes
  às `ScoreCardAndamento*`/`ScoreCardAtrasadas*` do legado.
- Card de scorecard no dashboard e em `/projects/[id]`.
- Banner amarelo "Janela de medição fechada" quando fora do horário.

**DoD:** dentro da janela, registrar medição funciona; fora, é bloqueado com
mensagem; após confirmar, valor não é mais editável; scorecard atualiza no
dashboard.

**Prompt:**
```
Sessão 6 (PLAN.md). Implemente medição + janela HB + scorecard.
src/lib/closing-window.ts checa ConfigFechamento. /projects/[id]/measurement
permite entrada de % e horas; confirmação cria Measurement.confirmed=true e
MeasurementLock impedindo edição. Bloqueia entrada fora da janela (banner
amarelo). src/lib/scorecard.ts implementa andamento/atrasadas10/atrasadasMais10.
Card de scorecard no dashboard e na página do projeto. DoD: medir dentro da
janela funciona, fora bloqueia, confirmada não edita, scorecard atualiza.
Commit "feat(measurement): medição + janela HB + scorecard".
```

### Sessão 7 — Painel Sinóptico + Curva S

**Objetivo:** substituir `PMO_Relatorio_Painel_Sinoptico_*` + `DadosFlagCurvaS` +
`RetornaBCWP`/`RetornaACWP`.

**Entregas:**
- `src/lib/curva-s.ts` — calcula séries Previsto / Realizado / BCWP / ACWP por
  semana ou mês a partir das `Task` e `Measurement`.
- `/projects/[id]/synoptic` — três visualizações:
  1. Curva S (linha; recharts).
  2. Previsto vs Realizado por equipe (barras).
  3. Heatmap de atrasos (semana × equipe).
- Filtros: período, equipe, área.
- Export PNG de cada gráfico.

**DoD:** projeto seed com 25 tarefas mostra Curva S coerente; mover medições e
ver curva atualizar.

**Prompt:**
```
Sessão 7 (PLAN.md). Implemente painel sinóptico + Curva S em
/projects/[id]/synoptic. src/lib/curva-s.ts calcula Previsto/Realizado/BCWP/ACWP
por semana e mês. UI com 3 gráficos (recharts): Curva S linha, Previsto vs
Realizado barras por equipe, heatmap de atrasos semana×equipe. Filtros período/
equipe/área. Export PNG. DoD: projeto seed mostra Curva S; medir uma tarefa,
recarregar, ver curva mudar. Commit "feat(reports): painel sinóptico + curva S".
```

### Sessão 8 — Issues, Risks, Actions, Change Requests

**Objetivo:** substituir `GVI_Issues`, `Risks`, `Actions`, `Chang_Request*` e
threads de comentários.

**Entregas:**
- `/projects/[id]/issues`, `/risks`, `/actions`, `/change-requests` —
  cada um com lista (filtros: status, severidade/criticidade, owner) +
  detalhe com workflow + thread de comentários.
- Workflow visual (badge colorido + dropdown de transição).
- ChangeRequest suporta múltiplas `Increment` e `Alteration` (espelha legado).
- Ação rápida no Gantt: "Criar issue para esta tarefa".
- KPI cards (totais, abertos, fechados, atraso médio).

**DoD:** criar issue, comentar, mudar status até CLOSED; criar risk com
prob/impact e ver exposure calculada; criar CR com 2 incrementos.

**Prompt:**
```
Sessão 8 (PLAN.md). CRUD + comentários + workflow para Issues, Risks, Actions,
ChangeRequests em /projects/[id]/{issues,risks,actions,change-requests}. Cada
tem lista filtrável, detalhe, thread de comments, transição de status. CR tem
Increments e Alterations. Atalho no Gantt "criar issue para tarefa". KPI
cards. DoD: criar issue→comentar→fechar; criar risk com prob/impact; criar
CR com 2 increments. Commit "feat(issues): issues/risks/actions/CRs + threads".
```

### Sessão 9 — Relatórios + Exportação Excel/PDF

**Objetivo:** substituir família `PMO_Relatorio_*` (Detalhado, Consolidado,
Criticidade, Comentários, Tarefas, Saneamento, Issues_Detalhado, Issues_KPI).

**Entregas:**
- `/projects/[id]/reports` — galeria de relatórios.
- Cada relatório: filtros → preview → exportação **Excel** (xlsx) e **PDF**
  (`@react-pdf/renderer`).
- Lista mínima:
  1. Detalhado (todas as tarefas).
  2. Consolidado (resumo por equipe/área).
  3. Criticidade (issues + risks ordenados por criticidade).
  4. Comentários (todas threads do projeto).
  5. Issues KPI (séries por mês).
  6. Tarefas atrasadas (com causas, se preenchidas).

**DoD:** abrir cada relatório, exportar XLSX e PDF, abrir no Excel/visualizador.

**Prompt:**
```
Sessão 9 (PLAN.md). 6 relatórios em /projects/[id]/reports: Detalhado,
Consolidado, Criticidade, Comentários, Issues KPI, Tarefas atrasadas. Cada um:
filtros → preview HTML → export XLSX (xlsx) e PDF (@react-pdf/renderer).
Galeria de relatórios na rota raiz de /reports. DoD: gerar e baixar XLSX+PDF
de cada um, conferir formatação no Excel. Commit "feat(reports): 6 relatórios
+ export XLSX/PDF".
```

### Sessão 10 — Módulos especializados (CUT, GVI, TCP, TIN, GRF) — P1

**Objetivo:** habilitar lógica/UI específica de cada módulo. Modela como
`Project.module: Module` + features condicionais.

**Por módulo:**
- **CUT (CutOver):** painel sinóptico CUT com janelas de operação.
- **GVI (Governança):** estimativas + saneamento + equipes próprias.
- **TCP (Teste de Carga):** entrada por casos de teste em vez de %.
- **TIN (Teste Integrado):** percentdt + cases.
- **GRF:** medição simplificada.

**DoD:** projeto-exemplo de cada módulo com sua UI específica.

**Prompt:**
```
Sessão 10 (PLAN.md). Módulos especializados. Adicione UIs específicas
condicionais a Project.module:
- CUT: painel CutOver com janelas de operação.
- GVI: estimativas + saneamento (planilha-like).
- TCP: entrada por casos de teste.
- TIN: percentdt + cases por caso de teste.
- GRF: medição simplificada (apenas %).
Use feature flags por módulo. Crie projetos seed adicionais (1 por módulo).
DoD: navegar em cada projeto e ver o painel correspondente. Commit
"feat(modules): UIs específicas CUT/GVI/TCP/TIN/GRF".
```

### Sessão 11 — Polimento, testes, Docker, CI

**Entregas:**
- Testes Vitest para `lib/*` (closing-window, scorecard, curva-s, parsers).
- Smoke test Playwright: login → criar projeto → importar TK.xls → ver Gantt →
  medir → ver scorecard.
- `Dockerfile` (multi-stage) + `docker-compose.yml` (app + Postgres).
- GitHub Actions: lint + typecheck + test em PRs.
- `OPERATIONS.md` (deploy, backup, migrations).
- Migração Prisma de SQLite → PostgreSQL no compose.

**DoD:** `docker-compose up` sobe app+db; smoke test passa; CI verde.

**Prompt:**
```
Sessão 11 (PLAN.md). Polimento + entrega.
- Vitest para src/lib/* (cobertura mínima 70%).
- Playwright smoke: login→criar projeto→importar TK.xls→Gantt→medir→scorecard.
- Dockerfile multi-stage + docker-compose.yml (app+Postgres+volume).
- GitHub Actions: lint+typecheck+vitest em PRs; build em main.
- OPERATIONS.md: deploy, backup pg_dump, migrations.
- Trocar provider Prisma para postgresql; ajustar seed.
DoD: docker-compose up sobe tudo; smoke test verde; CI verde.
Commit "chore: testes + docker + CI + ops docs".
```

---

## Cronograma sugerido

| Sessão | Esforço |
| ------ | ------- |
| 0      | ½ sessão (em andamento) |
| 1      | 1 sessão |
| 2      | 1 sessão |
| 3      | 1–2 sessões |
| 4      | 2 sessões |
| 5      | 1–2 sessões |
| 6      | 1 sessão |
| 7      | 1 sessão |
| 8      | 1–2 sessões |
| 9      | 1 sessão |
| 10     | 1–2 sessões |
| 11     | 1 sessão |
| **Total** | **~12–16 sessões** |

## Como usar este plano

1. Para iniciar uma sessão, abra uma nova conversa e cole o **prompt** dela.
2. Ao final de cada sessão, marque-a como ✅ concluída neste arquivo (commit).
3. Ajustes de escopo entre sessões: edite as seções correspondentes antes de
   começar a próxima.
4. PR único permanece aberto em `claude/evaluate-repository-SO97E` ou abra
   um PR por sessão (preferível para review menor).
