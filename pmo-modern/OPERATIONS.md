# OPERATIONS — CollabZ

Runbook operacional: dev local, deploy em produção (Vercel + Supabase),
backup, restore, migrations.

---

## 1. Dev local (Postgres via Docker)

### Pré-requisitos
- Node 20+
- Docker Desktop (ou Docker Engine + compose)

### Subir o ambiente
```bash
cd pmo-modern
cp .env.example .env       # já vem com DATABASE_URL apontando para localhost
npm install
npm run docker:db          # sobe Postgres em :5432
npm run db:deploy          # aplica migrations versionadas
npm run db:seed            # popula com dados de exemplo
npm run dev                # http://localhost:3000
```

### Logins do seed
| Usuário | Senha | Papel |
|---|---|---|
| `admin@pmo.local` | `admin123` | ADMIN |
| `gerente@pmo.local` | `gerente123` | MANAGER |
| `alpha@pmo.local` | `membro123` | MEMBER |
| `beta@pmo.local` | `membro123` | MEMBER |
| `viewer@pmo.local` | `viewer123` | VIEWER |

### Abrir Prisma Studio (UI do banco)
```bash
npm run db:studio          # http://localhost:5555
```

### Resetar tudo
```bash
npm run db:reset           # apaga volume + reaplica migrations + seed
```

### Parar Postgres local
```bash
docker compose down        # para os containers
docker compose down -v     # apaga também o volume (perde os dados)
```

---

## 2. Deploy em produção — Vercel + Supabase

### Passo a passo

#### 2.1. Criar Postgres no Supabase
1. Crie conta gratuita em <https://supabase.com>.
2. **New Project** → escolha região mais próxima (ex.: South America São Paulo).
3. Defina senha forte do Postgres e salve num gerenciador.
4. Espere ~2 min para o cluster subir.
5. Em **Settings → Database**:
   - **Connection string → URI** (Pooler / Transaction mode):
     - Copie. Use como `DATABASE_URL` (com `?pgbouncer=true&connection_limit=1`).
   - **Connection string → URI** (Direct connection):
     - Copie. Use como `DIRECT_URL` (necessário para `prisma migrate`).

#### 2.2. Aplicar migrations no Supabase
```bash
# localmente, com DATABASE_URL e DIRECT_URL apontando para o Supabase:
DATABASE_URL='postgresql://postgres.xxxxx:SENHA@aws-0-region.pooler.supabase.com:6543/postgres?pgbouncer=true&connection_limit=1' \
DIRECT_URL='postgresql://postgres.xxxxx:SENHA@aws-0-region.pooler.supabase.com:5432/postgres' \
npx prisma migrate deploy

# (opcional) popular com seed inicial:
DATABASE_URL='...' npx tsx scripts/seed.ts
```

#### 2.3. Deploy na Vercel
1. Crie conta em <https://vercel.com>.
2. **Add New → Project** → conecte sua conta GitHub e selecione o repositório.
3. **Root Directory**: `pmo-modern` (importante).
4. **Environment Variables** (Production + Preview):
   - `DATABASE_URL`: connection string de pooling do Supabase.
   - `DIRECT_URL`: connection string direta do Supabase.
   - `NEXTAUTH_SECRET`: rode `openssl rand -base64 32`.
   - `NEXTAUTH_URL`: `https://seu-projeto.vercel.app` (depois pode trocar por domínio próprio).
5. **Deploy**.
6. Após o primeiro deploy, configure o **domínio próprio** em Project Settings → Domains.

#### 2.4. Aplicar migrations futuras
```bash
# após criar um patch de schema:
npm run db:migrate -- --name nome-da-mudanca   # local
git add prisma/migrations/* && git commit && git push
# CI / Vercel build vão rodar prisma migrate deploy automaticamente
# (configurar postinstall ou script de build se necessário — ver "Build hook")
```

#### Build hook na Vercel (recomendado)
Para que o deploy aplique migrations automaticamente, ajuste o script de build:

```jsonc
// pmo-modern/package.json
{
  "scripts": {
    "build": "prisma migrate deploy && next build"
  }
}
```

> Atenção: `migrate deploy` precisa de `DIRECT_URL` em ambientes com pooler (Supabase).

---

## 3. Deploy alternativo — VPS com Docker Compose

Para hospedar em VPS próprio (Hetzner, DigitalOcean, AWS Lightsail):

```bash
# no servidor:
git clone https://github.com/fsannino/PMO.git
cd PMO/pmo-modern
cp .env.example .env
# editar .env com NEXTAUTH_SECRET forte, NEXTAUTH_URL público
docker compose up -d --build
docker compose exec app npx prisma migrate deploy
docker compose exec app npx tsx scripts/seed.ts   # opcional
```

Depois, **terminar HTTPS e domínio** com Caddy (recomendado, automático)
ou Nginx + certbot:

```caddyfile
# /etc/caddy/Caddyfile
collabz.suaempresa.com {
    reverse_proxy localhost:3000
}
```

---

## 4. Backup e restore

### Backup automático (Supabase)
- Free tier: backup diário automático, retenção 7 dias.
- Pago: PITR (point-in-time recovery).

### Backup manual com `pg_dump`
```bash
# do banco do Supabase (use DIRECT_URL, não pooler):
pg_dump "postgresql://postgres.xxx:SENHA@aws-0-region.pooler.supabase.com:5432/postgres" \
  --no-owner --no-acl --format=custom -f backup-$(date +%F).dump

# do Postgres local:
docker compose exec db pg_dump -U collabz collabz --no-owner --no-acl \
  --format=custom -f /tmp/backup.dump
docker compose cp db:/tmp/backup.dump ./backup-$(date +%F).dump
```

### Backup automatizado (cron)
Crie `/etc/cron.daily/collabz-backup`:

```bash
#!/bin/sh
DATE=$(date +%F)
pg_dump "$DATABASE_URL" --no-owner --no-acl --format=custom \
  -f /var/backups/collabz/backup-$DATE.dump
# manter só os últimos 30 dias
find /var/backups/collabz -name "backup-*.dump" -mtime +30 -delete
```

### Restore
```bash
# em banco vazio:
pg_restore --no-owner --no-acl -d "$DATABASE_URL" backup-2026-05-04.dump

# substituindo banco existente (cuidado!):
pg_restore --no-owner --no-acl --clean --if-exists -d "$DATABASE_URL" backup.dump
```

---

## 5. Migrations — workflow

### Criar nova migration (dev)
```bash
# 1. edite prisma/schema.prisma
# 2. gere e aplique:
npm run db:migrate -- --name nome-descritivo
# (cria prisma/migrations/<timestamp>_nome-descritivo/migration.sql)
# 3. commit e push:
git add prisma/migrations/* prisma/schema.prisma
git commit -m "feat(db): adicionar X"
git push
```

### Aplicar migrations em prod
```bash
DATABASE_URL='...' DIRECT_URL='...' npx prisma migrate deploy
```

### Rollback
Prisma não tem rollback automático. Para desfazer:
1. Crie nova migration que **reverte** as mudanças (ALTER/DROP inverso).
2. Ou restaure do backup mais recente anterior à migration problemática.

---

## 6. Monitoramento e logs

### Vercel
- **Project → Logs**: stdout/stderr em tempo real.
- **Project → Analytics**: requests, latência, erros.

### Supabase
- **Database → Reports**: queries lentas, conexões, tamanho.
- **Logs → Postgres**: log do banco.

### Sentry (opcional, recomendado para produção)
1. Crie conta em <https://sentry.io>.
2. `npm install @sentry/nextjs`.
3. `npx sentry-wizard@latest -i nextjs` — configura tudo.
4. Adicione `SENTRY_DSN` nas env vars da Vercel.

---

## 7. Rotação de secrets

Sempre que for rotacionar `NEXTAUTH_SECRET`:
1. Gere novo: `openssl rand -base64 32`.
2. Atualize na Vercel (Production env var).
3. Faça redeploy.
4. **Atenção**: todos os usuários serão deslogados (JWT antigos invalidados).

---

## 8. Troubleshooting

| Sintoma | Causa provável | Solução |
|---|---|---|
| `Error: P1001 — Can't reach database` | DATABASE_URL errado ou banco fora do ar | Conferir Supabase → Database → Status |
| `prisma migrate deploy` trava em "Applying" | Conexão via pooler não suporta migrate | Use `DIRECT_URL` (sem pgbouncer) |
| Login não funciona em prod | `NEXTAUTH_URL` incorreto | Conferir em env vars; deve ser HTTPS |
| Cookies não persistem | `NEXTAUTH_URL` ≠ URL real | Conferir; cookies são amarrados ao domínio |
| `next build` falha por tipo | Prisma client desatualizado | `npx prisma generate` |
| Vercel build esgota memória | Build pesado | Mude para "Hobby Plus" ou divida processo |

---

## 9. Custos estimados (free tiers, mai/2026)

| Serviço | Free tier | Pago a partir de |
|---|---|---|
| Vercel | 100 GB-h/mês compute, 100 GB bandwidth | Pro $20/mês |
| Supabase | 500 MB DB, 1 GB Storage, 2 GB egress | Pro $25/mês |
| Sentry | 5K erros/mês | Team $26/mês |

Para piloto interno (~10–50 usuários, ~5 projetos), **free tier é suficiente**.

---

## 10. Checklist de produção

Antes de "go live":
- [ ] `NEXTAUTH_SECRET` gerado com `openssl rand -base64 32`.
- [ ] `NEXTAUTH_URL` apontando para domínio HTTPS final.
- [ ] Banco com `prisma migrate deploy` aplicado e seed (opcional) rodado.
- [ ] Senha do admin do seed alterada (`/admin/users` → reset).
- [ ] Backup automatizado configurado.
- [ ] Sentry conectado (opcional mas recomendado).
- [ ] Domínio próprio com HTTPS.
- [ ] Janelas HB (`/admin/config-fechamento`) revisadas para o fuso real.
- [ ] Usuários reais criados (admin, managers, members, viewers).
- [ ] Smoke test: login → criar projeto → importar TK → ver Gantt → exportar XLSX/PDF.
