import * as XLSX from "xlsx";

export type ImportedTask = {
  externalId?: string;
  wbs?: string;
  name: string;
  startDate: Date;
  endDate: Date;
  durationDays?: number;
  percentDone?: number;
  isMilestone?: boolean;
  isSummary?: boolean;
  parentExternalId?: string;
  predecessorExternalIds?: string[];
};

export type ImportResult = {
  tasks: ImportedTask[];
  warnings: string[];
};

const HEADER_MAP: Record<string, keyof ImportedTask> = {
  id: "externalId",
  "id externo": "externalId",
  externalid: "externalId",
  wbs: "wbs",
  edt: "wbs",
  nome: "name",
  tarefa: "name",
  name: "name",
  task: "name",
  inicio: "startDate",
  início: "startDate",
  start: "startDate",
  startdate: "startDate",
  fim: "endDate",
  termino: "endDate",
  término: "endDate",
  finish: "endDate",
  enddate: "endDate",
  duracao: "durationDays",
  duração: "durationDays",
  duration: "durationDays",
  perc: "percentDone",
  percentual: "percentDone",
  "% concluido": "percentDone",
  "%concluido": "percentDone",
  percentdone: "percentDone",
  marco: "isMilestone",
  milestone: "isMilestone",
  resumo: "isSummary",
  summary: "isSummary",
  pai: "parentExternalId",
  parent: "parentExternalId",
  predecessores: "predecessorExternalIds",
  predecessoras: "predecessorExternalIds",
  predecessors: "predecessorExternalIds",
};

function normalize(s: string) {
  return s
    .toString()
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

function parseDate(value: unknown): Date | null {
  if (value == null || value === "") return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    // Excel serial date
    const utcDays = Math.floor(value - 25569);
    const ms = utcDays * 86400 * 1000;
    return new Date(ms);
  }
  const s = value.toString().trim();
  // dd/mm/yyyy ou dd-mm-yyyy
  const br = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})/);
  if (br) {
    const [, d, m, y] = br;
    const year = y.length === 2 ? 2000 + Number(y) : Number(y);
    return new Date(year, Number(m) - 1, Number(d));
  }
  const iso = new Date(s);
  return isNaN(iso.getTime()) ? null : iso;
}

function parseBool(v: unknown): boolean {
  if (typeof v === "boolean") return v;
  if (v == null) return false;
  const s = v.toString().trim().toLowerCase();
  return s === "1" || s === "sim" || s === "true" || s === "yes" || s === "x";
}

function parseNumber(v: unknown): number | undefined {
  if (v == null || v === "") return undefined;
  const n = typeof v === "number" ? v : Number(v.toString().replace(",", "."));
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Importa um cronograma a partir de um arquivo Excel (.xlsx/.xls).
 * Aceita cabeçalhos em PT-BR ou EN (case/acento-insensitive).
 * Colunas mínimas: nome, início, fim.
 */
export function parseExcelSchedule(buffer: Buffer): ImportResult {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return { tasks: [], warnings: ["Planilha vazia."] };

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    raw: true,
    defval: "",
  });
  if (rows.length === 0) return { tasks: [], warnings: ["Nenhuma linha."] };

  const sample = rows[0];
  const headerToField: Record<string, keyof ImportedTask> = {};
  for (const key of Object.keys(sample)) {
    const norm = normalize(key);
    if (HEADER_MAP[norm]) headerToField[key] = HEADER_MAP[norm];
  }

  const warnings: string[] = [];
  const tasks: ImportedTask[] = [];

  rows.forEach((row, idx) => {
    const t: Partial<ImportedTask> = {};
    for (const [src, field] of Object.entries(headerToField)) {
      const value = row[src];
      switch (field) {
        case "startDate":
        case "endDate": {
          const d = parseDate(value);
          if (d) t[field] = d;
          break;
        }
        case "durationDays":
        case "percentDone":
          t[field] = parseNumber(value);
          break;
        case "isMilestone":
        case "isSummary":
          t[field] = parseBool(value);
          break;
        case "predecessorExternalIds":
          t[field] = value
            ? value
                .toString()
                .split(/[,;]/)
                .map((x) => x.trim())
                .filter(Boolean)
            : [];
          break;
        default:
          t[field] = value ? value.toString().trim() : undefined;
      }
    }

    if (!t.name) {
      warnings.push(`Linha ${idx + 2}: ignorada (sem nome).`);
      return;
    }
    if (!t.startDate || !t.endDate) {
      warnings.push(`Linha ${idx + 2} ("${t.name}"): datas ausentes/ inválidas.`);
      return;
    }
    if (t.endDate < t.startDate) {
      warnings.push(`Linha ${idx + 2} ("${t.name}"): fim < início.`);
    }
    tasks.push(t as ImportedTask);
  });

  return { tasks, warnings };
}
