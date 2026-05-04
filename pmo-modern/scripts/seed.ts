import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";

const prisma = new PrismaClient();

function days(n: number) {
  return n * 24 * 60 * 60 * 1000;
}

async function main() {
  console.log("🌱 Limpando dados antigos…");
  // Ordem importa pelas FKs.
  await prisma.taskHistory.deleteMany();
  await prisma.measurementLock.deleteMany();
  await prisma.measurement.deleteMany();
  await prisma.comment.deleteMany();
  await prisma.taskDependency.deleteMany();
  await prisma.traceability.deleteMany();
  await prisma.changeRequestIncrement.deleteMany();
  await prisma.changeRequestAlteration.deleteMany();
  await prisma.changeRequest.deleteMany();
  await prisma.action.deleteMany();
  await prisma.issue.deleteMany();
  await prisma.risk.deleteMany();
  await prisma.task.deleteMany();
  await prisma.importLog.deleteMany();
  await prisma.document.deleteMany();
  await prisma.configFechamento.deleteMany();
  await prisma.access.deleteMany();
  await prisma.membership.deleteMany();
  await prisma.project.deleteMany();
  await prisma.user.deleteMany();
  await prisma.equipe.deleteMany();
  await prisma.comite.deleteMany();
  await prisma.governanca.deleteMany();
  await prisma.frente.deleteMany();
  await prisma.area.deleteMany();
  await prisma.unidade.deleteMany();

  // ─── Dados mestre ────────────────────────────────────────────────────
  console.log("🏢 Criando unidades, áreas, frentes, governanças, comitês, equipes…");

  const [unREFAP, unRECAP] = await Promise.all([
    prisma.unidade.create({ data: { code: "REFAP", name: "Refinaria Alberto Pasqualini" } }),
    prisma.unidade.create({ data: { code: "RECAP", name: "Refinaria de Capuava" } }),
  ]);

  const areas = await Promise.all([
    prisma.area.create({ data: { code: "TI", name: "Tecnologia da Informação" } }),
    prisma.area.create({ data: { code: "MAN", name: "Manutenção" } }),
    prisma.area.create({ data: { code: "OPE", name: "Operações" } }),
    prisma.area.create({ data: { code: "ENG", name: "Engenharia" } }),
    prisma.area.create({ data: { code: "PCM", name: "Planejamento e Controle" } }),
    prisma.area.create({ data: { code: "QSE", name: "Qualidade, Saúde e Segurança" } }),
  ]);
  const [aTI, aMAN, aOPE, aENG, aPCM, aQSE] = areas;

  const [fIMPL, fTEST, fOPER] = await Promise.all([
    prisma.frente.create({ data: { code: "IMPL", name: "Implantação" } }),
    prisma.frente.create({ data: { code: "TEST", name: "Testes" } }),
    prisma.frente.create({ data: { code: "OPER", name: "Operação Assistida" } }),
  ]);

  const govSMR = await prisma.governanca.create({
    data: { code: "GOV-SMR", name: "Governança SMR" },
  });

  const comGD = await prisma.comite.create({
    data: { code: "COM-GD", name: "Comitê de Gestão e Diretoria" },
  });

  const [eqAlpha, eqBeta, eqGama] = await Promise.all([
    prisma.equipe.create({ data: { code: "ALPHA", name: "Equipe Alpha" } }),
    prisma.equipe.create({ data: { code: "BETA",  name: "Equipe Beta" } }),
    prisma.equipe.create({ data: { code: "GAMA",  name: "Equipe Gama" } }),
  ]);

  // ─── Usuários ────────────────────────────────────────────────────────
  console.log("👥 Criando usuários…");
  const hash = (pwd: string) => bcrypt.hashSync(pwd, 10);

  const [admin, manager, mAlpha, mBeta, viewer] = await Promise.all([
    prisma.user.create({
      data: {
        email: "admin@pmo.local",
        name: "Administrador",
        passwordHash: hash("admin123"),
        role: "ADMIN",
        unidadeId: unREFAP.id,
        areaId: aTI.id,
      },
    }),
    prisma.user.create({
      data: {
        email: "gerente@pmo.local",
        name: "Joana Gerente",
        passwordHash: hash("gerente123"),
        role: "MANAGER",
        unidadeId: unREFAP.id,
        areaId: aPCM.id,
      },
    }),
    prisma.user.create({
      data: {
        email: "alpha@pmo.local",
        name: "Pedro Alpha",
        passwordHash: hash("membro123"),
        role: "MEMBER",
        unidadeId: unREFAP.id,
        areaId: aMAN.id,
      },
    }),
    prisma.user.create({
      data: {
        email: "beta@pmo.local",
        name: "Marina Beta",
        passwordHash: hash("membro123"),
        role: "MEMBER",
        unidadeId: unREFAP.id,
        areaId: aOPE.id,
      },
    }),
    prisma.user.create({
      data: {
        email: "viewer@pmo.local",
        name: "Rafael Viewer",
        passwordHash: hash("viewer123"),
        role: "VIEWER",
        unidadeId: unRECAP.id,
        areaId: aQSE.id,
      },
    }),
  ]);

  // ─── Projeto SMR ────────────────────────────────────────────────────
  console.log("📁 Criando projeto SMR-2026…");
  const start = new Date("2026-02-02"); // segunda-feira
  const project = await prisma.project.create({
    data: {
      code: "SMR-2026",
      name: "CollabZ Manutenção e Refino 2026",
      description: "Projeto SMR de modernização das rotinas de PCM e parada programada.",
      module: "PMO",
      status: "ACTIVE",
      priority: 1,
      unidadeId: unREFAP.id,
      governancaId: govSMR.id,
      startDate: start,
      endDate: new Date(start.getTime() + days(180)),
      baselineDate: start,
      ownerId: manager.id,
    },
  });

  // membros + matriz de acesso
  await prisma.membership.createMany({
    data: [
      { userId: admin.id, projectId: project.id, role: "ADMIN" },
      { userId: manager.id, projectId: project.id, role: "MANAGER" },
      { userId: mAlpha.id, projectId: project.id, role: "MEMBER" },
      { userId: mBeta.id,  projectId: project.id, role: "MEMBER" },
      { userId: viewer.id, projectId: project.id, role: "VIEWER" },
    ],
  });

  await prisma.access.createMany({
    data: [
      { userId: admin.id,   projectId: project.id, module: "PMO", canRead: true, canWrite: true, canAdmin: true },
      { userId: manager.id, projectId: project.id, module: "PMO", canRead: true, canWrite: true, canAdmin: true },
      { userId: mAlpha.id,  projectId: project.id, module: "PMO", canRead: true, canWrite: true },
      { userId: mBeta.id,   projectId: project.id, module: "PMO", canRead: true, canWrite: true },
      { userId: viewer.id,  projectId: project.id, module: "PMO", canRead: true },
    ],
  });

  // ─── Janelas HB padrão (seg-sex 8h-18h, projeto + global) ────────────
  console.log("⏰ Configurando janelas HB…");
  const hbDays = [
    { dayOfWeek: 0, enabled: false }, // dom
    { dayOfWeek: 1, enabled: true  }, // seg
    { dayOfWeek: 2, enabled: true  }, // ter
    { dayOfWeek: 3, enabled: true  }, // qua
    { dayOfWeek: 4, enabled: true  }, // qui
    { dayOfWeek: 5, enabled: true  }, // sex
    { dayOfWeek: 6, enabled: false }, // sáb
  ];
  await prisma.configFechamento.createMany({
    data: hbDays.map((d) => ({
      projectId: project.id,
      dayOfWeek: d.dayOfWeek,
      enabled: d.enabled,
      startTime: "08:00",
      endTime: "18:00",
    })),
  });
  // janela global default
  await prisma.configFechamento.createMany({
    data: hbDays.map((d) => ({
      projectId: null,
      dayOfWeek: d.dayOfWeek,
      enabled: d.enabled,
      startTime: "08:00",
      endTime: "18:00",
    })),
  });

  // ─── Cronograma: 25 tarefas hierárquicas ─────────────────────────────
  console.log("📋 Criando 25 tarefas hierárquicas com predecessores…");

  type SeedTask = {
    key: string;
    wbs: string;
    name: string;
    parent?: string;
    offsetStart: number; // dias desde project.startDate
    duration: number;    // dias
    isMilestone?: boolean;
    isSummary?: boolean;
    assignee?: string;
    equipe?: string;
    area?: string;
    frente?: string;
    predecessors?: string[]; // keys
    percent?: number;
  };

  const seedTasks: SeedTask[] = [
    // 1. Mobilização (resumo)
    { key: "T1",   wbs: "1",     name: "Mobilização e planejamento",        offsetStart: 0,  duration: 20, isSummary: true, area: aPCM.code, frente: fIMPL.code },
    { key: "T1.1", wbs: "1.1",   name: "Reunião de kickoff",                parent: "T1", offsetStart: 0,  duration: 1,  isMilestone: true, assignee: manager.email, equipe: eqAlpha.code, area: aPCM.code, frente: fIMPL.code, percent: 100 },
    { key: "T1.2", wbs: "1.2",   name: "Levantamento de requisitos",        parent: "T1", offsetStart: 1,  duration: 10, predecessors: ["T1.1"], assignee: mAlpha.email, equipe: eqAlpha.code, area: aTI.code,  frente: fIMPL.code, percent: 60 },
    { key: "T1.3", wbs: "1.3",   name: "Definição de escopo",               parent: "T1", offsetStart: 8,  duration: 5,  predecessors: ["T1.2"], assignee: manager.email, equipe: eqAlpha.code, area: aPCM.code, frente: fIMPL.code, percent: 30 },
    { key: "T1.4", wbs: "1.4",   name: "Aprovação de baseline",             parent: "T1", offsetStart: 13, duration: 2,  predecessors: ["T1.3"], assignee: manager.email, equipe: eqAlpha.code, area: aPCM.code, frente: fIMPL.code, isMilestone: true },

    // 2. Engenharia (resumo)
    { key: "T2",   wbs: "2",     name: "Engenharia",                        offsetStart: 15, duration: 60, isSummary: true, area: aENG.code, frente: fIMPL.code },
    { key: "T2.1", wbs: "2.1",   name: "Modelagem de processos AS-IS",      parent: "T2", offsetStart: 15, duration: 15, predecessors: ["T1.4"], assignee: mBeta.email, equipe: eqBeta.code, area: aENG.code, frente: fIMPL.code, percent: 20 },
    { key: "T2.2", wbs: "2.2",   name: "Desenho TO-BE",                     parent: "T2", offsetStart: 30, duration: 20, predecessors: ["T2.1"], assignee: mBeta.email,  equipe: eqBeta.code, area: aENG.code, frente: fIMPL.code },
    { key: "T2.3", wbs: "2.3",   name: "Especificação técnica",             parent: "T2", offsetStart: 50, duration: 15, predecessors: ["T2.2"], assignee: mAlpha.email, equipe: eqBeta.code, area: aENG.code, frente: fIMPL.code },
    { key: "T2.4", wbs: "2.4",   name: "Aprovação técnica",                 parent: "T2", offsetStart: 65, duration: 2,  predecessors: ["T2.3"], assignee: manager.email, equipe: eqBeta.code, area: aENG.code, frente: fIMPL.code, isMilestone: true },

    // 3. Construção (resumo)
    { key: "T3",   wbs: "3",     name: "Construção",                        offsetStart: 67, duration: 60, isSummary: true, area: aTI.code, frente: fIMPL.code },
    { key: "T3.1", wbs: "3.1",   name: "Setup de ambientes",                parent: "T3", offsetStart: 67, duration: 5,  predecessors: ["T2.4"], assignee: mAlpha.email, equipe: eqGama.code, area: aTI.code, frente: fIMPL.code },
    { key: "T3.2", wbs: "3.2",   name: "Desenvolvimento - módulo cronograma", parent: "T3", offsetStart: 72, duration: 25, predecessors: ["T3.1"], assignee: mAlpha.email, equipe: eqGama.code, area: aTI.code, frente: fIMPL.code },
    { key: "T3.3", wbs: "3.3",   name: "Desenvolvimento - módulo medição",   parent: "T3", offsetStart: 80, duration: 20, predecessors: ["T3.1"], assignee: mBeta.email,  equipe: eqGama.code, area: aTI.code, frente: fIMPL.code },
    { key: "T3.4", wbs: "3.4",   name: "Desenvolvimento - relatórios",       parent: "T3", offsetStart: 100, duration: 20, predecessors: ["T3.2", "T3.3"], assignee: mBeta.email, equipe: eqGama.code, area: aTI.code, frente: fIMPL.code },
    { key: "T3.5", wbs: "3.5",   name: "Code review e correções",            parent: "T3", offsetStart: 115, duration: 12, predecessors: ["T3.4"], assignee: mAlpha.email, equipe: eqGama.code, area: aTI.code, frente: fIMPL.code },

    // 4. Testes (resumo)
    { key: "T4",   wbs: "4",     name: "Testes",                            offsetStart: 127, duration: 30, isSummary: true, area: aQSE.code, frente: fTEST.code },
    { key: "T4.1", wbs: "4.1",   name: "Testes unitários",                  parent: "T4", offsetStart: 127, duration: 8,  predecessors: ["T3.5"], assignee: mAlpha.email, equipe: eqAlpha.code, area: aTI.code,  frente: fTEST.code },
    { key: "T4.2", wbs: "4.2",   name: "Testes integrados",                 parent: "T4", offsetStart: 135, duration: 10, predecessors: ["T4.1"], assignee: mBeta.email,  equipe: eqAlpha.code, area: aQSE.code, frente: fTEST.code },
    { key: "T4.3", wbs: "4.3",   name: "Testes de carga",                   parent: "T4", offsetStart: 145, duration: 6,  predecessors: ["T4.2"], assignee: mAlpha.email, equipe: eqAlpha.code, area: aQSE.code, frente: fTEST.code },
    { key: "T4.4", wbs: "4.4",   name: "UAT (homologação)",                 parent: "T4", offsetStart: 151, duration: 6,  predecessors: ["T4.3"], assignee: manager.email, equipe: eqAlpha.code, area: aPCM.code, frente: fTEST.code },

    // 5. Cutover & Operação (resumo)
    { key: "T5",   wbs: "5",     name: "Cutover e operação assistida",      offsetStart: 157, duration: 23, isSummary: true, area: aOPE.code, frente: fOPER.code },
    { key: "T5.1", wbs: "5.1",   name: "Treinamento de usuários",           parent: "T5", offsetStart: 157, duration: 7,  predecessors: ["T4.4"], assignee: mBeta.email,  equipe: eqBeta.code, area: aPCM.code, frente: fOPER.code },
    { key: "T5.2", wbs: "5.2",   name: "Cutover",                           parent: "T5", offsetStart: 164, duration: 3,  predecessors: ["T5.1"], assignee: manager.email, equipe: eqBeta.code, area: aOPE.code, frente: fOPER.code, isMilestone: true },
    { key: "T5.3", wbs: "5.3",   name: "Operação assistida",                parent: "T5", offsetStart: 167, duration: 10, predecessors: ["T5.2"], assignee: mAlpha.email, equipe: eqBeta.code, area: aOPE.code, frente: fOPER.code },
    { key: "T5.4", wbs: "5.4",   name: "Encerramento e lições aprendidas",  parent: "T5", offsetStart: 177, duration: 3,  predecessors: ["T5.3"], assignee: manager.email, equipe: eqBeta.code, area: aPCM.code, frente: fOPER.code, isMilestone: true },
  ];

  const userByEmail = new Map([admin, manager, mAlpha, mBeta, viewer].map((u) => [u.email, u]));
  const equipeByCode = new Map([eqAlpha, eqBeta, eqGama].map((e) => [e.code, e]));
  const areaByCode = new Map(areas.map((a) => [a.code, a]));
  const frenteByCode = new Map([fIMPL, fTEST, fOPER].map((f) => [f.code, f]));

  // duas passadas: criar tarefas, depois ligar parent + dependências.
  const created = new Map<string, { id: string }>();

  for (const t of seedTasks) {
    const startDate = new Date(start.getTime() + days(t.offsetStart));
    const endDate = new Date(start.getTime() + days(t.offsetStart + t.duration));
    const c = await prisma.task.create({
      data: {
        projectId: project.id,
        wbs: t.wbs,
        externalId: t.key,
        name: t.name,
        startDate,
        endDate,
        baselineStart: startDate,
        baselineEnd: endDate,
        durationDays: t.duration,
        percentDone: t.percent ?? 0,
        status: t.isMilestone && (t.percent ?? 0) >= 100
          ? "DONE"
          : (t.percent ?? 0) > 0
            ? "IN_PROGRESS"
            : "NOT_STARTED",
        isMilestone: !!t.isMilestone,
        isSummary: !!t.isSummary,
        assigneeId: t.assignee ? userByEmail.get(t.assignee)?.id : undefined,
        equipeId: t.equipe ? equipeByCode.get(t.equipe)?.id : undefined,
        areaId: t.area ? areaByCode.get(t.area)?.id : undefined,
        frenteId: t.frente ? frenteByCode.get(t.frente)?.id : undefined,
      },
    });
    created.set(t.key, c);
  }

  // segunda passada: parents + dependências
  for (const t of seedTasks) {
    const me = created.get(t.key)!;
    if (t.parent) {
      await prisma.task.update({
        where: { id: me.id },
        data: { parentId: created.get(t.parent)!.id },
      });
    }
    for (const pred of t.predecessors ?? []) {
      const p = created.get(pred);
      if (!p) continue;
      await prisma.taskDependency.create({
        data: {
          predecessorId: p.id,
          successorId: me.id,
          type: "FS",
          lagDays: 0,
        },
      });
    }
    // histórico CREATED para todas
    await prisma.taskHistory.create({
      data: { taskId: me.id, userId: admin.id, action: "CREATED", note: "Importado pelo seed" },
    });
  }

  // ─── Medições já confirmadas para tarefas com percentual ──────────────
  console.log("📏 Criando medições confirmadas iniciais…");
  for (const t of seedTasks) {
    if (!t.percent || t.isSummary) continue;
    const taskId = created.get(t.key)!.id;
    const period = `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, "0")}`;
    const m = await prisma.measurement.create({
      data: {
        taskId,
        userId: t.assignee ? userByEmail.get(t.assignee)!.id : manager.id,
        percentDone: t.percent,
        hoursWorked: t.percent * 0.4,
        period,
        confirmed: true,
        confirmedAt: new Date(),
        comment: "Medição inicial via seed",
      },
    });
    await prisma.measurementLock.create({
      data: { taskId, period, measurementId: m.id },
    });
  }

  // ─── Issues ──────────────────────────────────────────────────────────
  console.log("🐞 Criando 3 issues…");
  await prisma.issue.create({
    data: {
      projectId: project.id,
      ownerId: mAlpha.id,
      areaId: aTI.id,
      taskId: created.get("T1.2")!.id,
      title: "Stakeholder de operações ainda não disponibilizou agenda",
      description: "Pendência para fechamento dos requisitos do módulo de operações.",
      severity: "HIGH",
      priority: "HIGH",
      status: "OPEN",
      workflow: "ABERTO",
    },
  });
  await prisma.issue.create({
    data: {
      projectId: project.id,
      ownerId: mBeta.id,
      areaId: aENG.id,
      title: "Padrão de nomenclatura de tags divergente entre unidades",
      description: "Necessário alinhar com Engenharia e PCM antes do TO-BE.",
      severity: "MEDIUM",
      priority: "MEDIUM",
      status: "IN_PROGRESS",
      workflow: "EM_ANALISE",
    },
  });
  await prisma.issue.create({
    data: {
      projectId: project.id,
      ownerId: manager.id,
      areaId: aPCM.id,
      title: "Inconsistência de horímetros nos equipamentos críticos",
      description: "Detectada diferença entre o que está no SAP e na planilha de campo.",
      severity: "LOW",
      priority: "LOW",
      status: "RESOLVED",
      workflow: "RESOLVIDO",
      closedAt: new Date(),
    },
  });

  // ─── Risks ───────────────────────────────────────────────────────────
  console.log("⚠️  Criando 2 riscos…");
  const r1 = await prisma.risk.create({
    data: {
      projectId: project.id,
      ownerId: manager.id,
      title: "Atraso na liberação de equipamentos para teste de carga",
      description: "Equipe de operação pode não conseguir liberar a janela necessária.",
      probability: 0.4,
      impact: 0.7,
      exposure: 0.4 * 0.7,
      mitigation: "Reservar janela alternativa de fim de semana em julho.",
      status: "MITIGATING",
    },
  });
  const r2 = await prisma.risk.create({
    data: {
      projectId: project.id,
      ownerId: mAlpha.id,
      title: "Curva de aprendizado da nova UI por parte da equipe de campo",
      description: "Pode haver resistência inicial ao novo sistema.",
      probability: 0.6,
      impact: 0.5,
      exposure: 0.6 * 0.5,
      mitigation: "Treinamento obrigatório + gemba walks na primeira semana.",
      status: "ANALYZING",
    },
  });

  // ─── Change Request ──────────────────────────────────────────────────
  console.log("📝 Criando 1 Change Request com 2 incrementos…");
  const cr = await prisma.changeRequest.create({
    data: {
      projectId: project.id,
      ownerId: manager.id,
      comiteId: comGD.id,
      name: "Inclusão de relatório executivo mensal",
      description: "Solicitação do comitê para gerar PDF executivo no fim de cada mês.",
      priority: "MEDIUM",
      status: "UNDER_REVIEW",
      dueDate: new Date(start.getTime() + days(90)),
    },
  });
  await prisma.changeRequestIncrement.createMany({
    data: [
      { changeRequestId: cr.id, sequence: 1, description: "Layout do relatório aprovado pelo comitê." },
      { changeRequestId: cr.id, sequence: 2, description: "Inclusão de KPI de aderência à baseline." },
    ],
  });

  // ─── Comentários de exemplo ──────────────────────────────────────────
  await prisma.comment.create({
    data: { authorId: manager.id, body: "Vamos priorizar isso na próxima sprint.", riskId: r1.id },
  });
  await prisma.comment.create({
    data: { authorId: mBeta.id, body: "Preparei o material de treinamento, posso compartilhar.", riskId: r2.id },
  });
  await prisma.comment.create({
    data: { authorId: admin.id, body: "Aprovo o layout proposto, sigam com o desenvolvimento.", changeRequestId: cr.id },
  });

  // ─── Action de exemplo (extra-DoD, baixo custo) ──────────────────────
  await prisma.action.create({
    data: {
      projectId: project.id,
      ownerId: mAlpha.id,
      title: "Validar templates de carga IS/RK/AC/CR/TK/TC com PCM",
      description: "Confirmar com PCM se os cabeçalhos do legado serão preservados.",
      dueDate: new Date(start.getTime() + days(30)),
      priority: "HIGH",
    },
  });

  // ─── Projetos exemplo para outros módulos (Sessão 10) ───────────────
  console.log("🧩 Criando projetos exemplo de outros módulos…");
  const moduleSeed: Array<{ code: string; name: string; module: string; offset: number; duration: number }> = [
    { code: "CUT-2026",   name: "CutOver — Migração SAP",        module: "CUT",   offset: 100, duration: 30 },
    { code: "GVI-2026",   name: "Governança Integrada (saneamento)", module: "GVI", offset: 0, duration: 365 },
    { code: "TCP-2026",   name: "Teste de Carga — Plataforma",   module: "TCP",   offset: 60,  duration: 45 },
    { code: "TIN-2026",   name: "Teste Integrado — SAP × MES",   module: "TIN",   offset: 80,  duration: 60 },
    { code: "GRF-2026",   name: "Acompanhamento Gráfico — KPIs", module: "GRF",   offset: 0,   duration: 365 },
    { code: "LIGHT-2026", name: "Light — Acompanhamento simples", module: "LIGHT", offset: 0,  duration: 180 },
  ];
  for (const ms of moduleSeed) {
    const ps = new Date(start.getTime() + days(ms.offset));
    const pe = new Date(ps.getTime() + days(ms.duration));
    const proj = await prisma.project.create({
      data: {
        code: ms.code,
        name: ms.name,
        description: `Projeto de demonstração do módulo ${ms.module}.`,
        module: ms.module,
        status: "ACTIVE",
        priority: 5,
        unidadeId: unREFAP.id,
        startDate: ps,
        endDate: pe,
        baselineDate: ps,
        ownerId: manager.id,
      },
    });
    await prisma.access.createMany({
      data: [
        { userId: admin.id,   projectId: proj.id, module: ms.module, canRead: true, canWrite: true, canAdmin: true },
        { userId: manager.id, projectId: proj.id, module: ms.module, canRead: true, canWrite: true, canAdmin: true },
        { userId: mAlpha.id,  projectId: proj.id, module: ms.module, canRead: true, canWrite: true },
      ],
    });
    // 3 tarefas exemplo cada
    for (let i = 1; i <= 3; i++) {
      const ts = new Date(ps.getTime() + days((ms.duration / 3) * (i - 1)));
      const te = new Date(ts.getTime() + days(ms.duration / 3));
      await prisma.task.create({
        data: {
          projectId: proj.id,
          wbs: String(i),
          name: `${ms.module} — atividade ${i}`,
          startDate: ts,
          endDate: te,
          baselineStart: ts,
          baselineEnd: te,
          durationDays: Math.round(ms.duration / 3),
          percentDone: i === 1 ? 100 : i === 2 ? 40 : 0,
          status: i === 1 ? "DONE" : i === 2 ? "IN_PROGRESS" : "NOT_STARTED",
          assigneeId: mAlpha.id,
          equipeId: eqAlpha.id,
        },
      });
    }
  }

  console.log("✅ Seed concluído!");
  console.log("   Login admin: admin@pmo.local / admin123");
  console.log("   Login gerente: gerente@pmo.local / gerente123");
  console.log("   Login membros: alpha@pmo.local | beta@pmo.local / membro123");
  console.log("   Login viewer: viewer@pmo.local / viewer123");
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
