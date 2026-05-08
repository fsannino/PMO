# PMO — Collab:Build

Repositório que contém **dois sistemas**:

1. **Legado (Classic ASP + VB6, ~2003-2005)** — código original do Collab:Build.
2. **`pmo-modern/`** — reescrita moderna em **Next.js + TypeScript + Prisma**, com as funcionalidades essenciais reimplementadas: cadastro de projetos e tarefas, medições, issues, riscos, importação de cronograma a partir de **Excel** e **MS Project (XML)**, visualização **Gantt**, relatórios e tela de manutenção.

---

## Estrutura

```
.
├── pmo-modern/                # Aplicação moderna (ver README dedicado)
│
├── SMR/                       # [LEGADO] App ASP principal (CUT, GVI, GRF, TCP, TIN, PMO)
├── SMR_Light/                 # [LEGADO] App ASP "Light"
├── Carga/                     # [LEGADO] VB6 — DLL de carga de planilhas
├── Manutencao_Light/          # [LEGADO] VB6 — manutenção do BD Light + planilhas modelo
├── Manutencao_SMR/            # [LEGADO] VB6 — manutenção dos BDs SMR/CUT/EP/SC/TCP/TIN/SMRTRAB
├── Queries/                   # [LEGADO] Consultas Excel (.dqy)
└── Scripts/                   # [LEGADO] Scripts SQL (Script_SMR.txt, Script_SMR_Light.txt)
```

## Começando com a versão moderna

Banco de dados: **PostgreSQL** (via Docker localmente, **Supabase** em produção).

```bash
cd pmo-modern
cp .env.example .env
npm install
npm run docker:db        # sobe Postgres em :5432 (precisa Docker)
npm run db:deploy        # aplica migrations
npm run db:seed          # popula dados de exemplo
npm run dev              # http://localhost:3000
```

Abra <http://localhost:3000>. Login padrão criado pelo seed: `admin@pmo.local` / `admin123`.

Mais detalhes:
- [`pmo-modern/PLAN.md`](pmo-modern/PLAN.md) — plano de modernização em sessões.
- [`pmo-modern/OPERATIONS.md`](pmo-modern/OPERATIONS.md) — runbook de deploy, backup, migrations.

## Status do legado

O código em `SMR/`, `SMR_Light/`, `Carga/`, `Manutencao_*/` e `Queries/` é mantido **somente para referência histórica e migração**. Não é executado em produção a partir deste repositório. Os scripts SQL em `Scripts/` continuam úteis para entender o esquema original e foram a base do modelo Prisma de `pmo-modern/`.
