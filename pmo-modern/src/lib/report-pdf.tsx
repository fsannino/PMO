// Export PDF para os 6 relatórios — usa @react-pdf/renderer.

import { Document, Page, Text, View, StyleSheet, pdf } from "@react-pdf/renderer";
import * as React from "react";
import type {
  ReportDetalhadoRow,
  ReportConsolidadoRow,
  ReportCriticidadeRow,
  ReportCommentRow,
  ReportIssuesKPIRow,
  ReportAtrasadasRow,
  ReportHeader,
  ReportId,
} from "./reports";
import { REPORTS } from "./reports";

const styles = StyleSheet.create({
  page: { padding: 28, fontSize: 9, fontFamily: "Helvetica" },
  h1: { fontSize: 16, fontWeight: 700, marginBottom: 4 },
  meta: { fontSize: 9, color: "#475569", marginBottom: 12 },
  kpiBox: {
    flexDirection: "row",
    backgroundColor: "#f1f5f9",
    padding: 8,
    borderRadius: 4,
    marginBottom: 12,
  },
  kpiCell: { flex: 1 },
  kpiLabel: { color: "#475569", fontSize: 8 },
  kpiValue: { fontSize: 12, fontWeight: 700, marginTop: 2 },
  table: { width: "100%", border: "1px solid #cbd5e1", borderRadius: 2 },
  row: { flexDirection: "row", borderBottom: "1px solid #e2e8f0" },
  rowLast: { flexDirection: "row" },
  thead: { backgroundColor: "#f1f5f9", fontWeight: 700 },
  cell: { padding: 4, borderRight: "1px solid #e2e8f0" },
  cellLast: { padding: 4 },
  footer: { position: "absolute", bottom: 12, left: 28, right: 28, textAlign: "center", fontSize: 8, color: "#94a3b8" },
});

function Header({ title, header }: { title: string; header: ReportHeader }) {
  return (
    <>
      <Text style={styles.h1}>{title}</Text>
      <Text style={styles.meta}>
        {header.projectCode} — {header.projectName} · gerado em {header.generatedAt} por {header.generatedBy}
      </Text>
      <View style={styles.kpiBox}>
        <View style={styles.kpiCell}>
          <Text style={styles.kpiLabel}>Andamento</Text>
          <Text style={styles.kpiValue}>{header.scorecard.andamentoPct}%</Text>
        </View>
        <View style={styles.kpiCell}>
          <Text style={styles.kpiLabel}>Tarefas</Text>
          <Text style={styles.kpiValue}>{header.scorecard.totalTasks}</Text>
        </View>
        <View style={styles.kpiCell}>
          <Text style={styles.kpiLabel}>Atrasadas</Text>
          <Text style={styles.kpiValue}>{header.scorecard.late}</Text>
        </View>
      </View>
    </>
  );
}

function Table({ headers, rows, widths }: { headers: string[]; rows: string[][]; widths: number[] }) {
  return (
    <View style={styles.table}>
      <View style={[styles.row, styles.thead]}>
        {headers.map((h, i) => (
          <View key={i} style={[i === headers.length - 1 ? styles.cellLast : styles.cell, { width: `${widths[i]}%` }]}>
            <Text>{h}</Text>
          </View>
        ))}
      </View>
      {rows.map((r, ri) => (
        <View key={ri} style={ri === rows.length - 1 ? styles.rowLast : styles.row}>
          {r.map((c, i) => (
            <View key={i} style={[i === r.length - 1 ? styles.cellLast : styles.cell, { width: `${widths[i]}%` }]}>
              <Text>{c}</Text>
            </View>
          ))}
        </View>
      ))}
    </View>
  );
}

function ReportDoc({
  title,
  header,
  headers,
  rows,
  widths,
}: {
  title: string;
  header: ReportHeader;
  headers: string[];
  rows: string[][];
  widths: number[];
}) {
  return (
    <Document>
      <Page size="A4" orientation="landscape" style={styles.page}>
        <Header title={title} header={header} />
        <Table headers={headers} rows={rows} widths={widths} />
        <Text style={styles.footer} render={({ pageNumber, totalPages }) => `${pageNumber} / ${totalPages}`} fixed />
      </Page>
    </Document>
  );
}

export async function buildReportPdf(reportId: ReportId, header: ReportHeader, data: unknown): Promise<Buffer> {
  const title = REPORTS[reportId].label;
  let h: string[] = [];
  let widths: number[] = [];
  let rows: string[][] = [];

  switch (reportId) {
    case "detalhado": {
      h = ["WBS", "Nome", "Início", "Fim", "Dur.", "%", "Status", "Resp.", "Equipe", "Atraso"];
      widths = [6, 28, 9, 9, 5, 5, 11, 12, 8, 7];
      rows = (data as ReportDetalhadoRow[]).map((r) => [
        r.wbs, r.name, r.start, r.end, String(r.duration), `${r.percent}%`, r.status, r.assignee, r.equipe, r.daysLate ? `${r.daysLate}d` : "—",
      ]);
      break;
    }
    case "consolidado": {
      h = ["Equipe", "Total", "Concl.", "Em curso", "Não iniciadas", "Atrasadas", "Trabalho (d)", "% Médio"];
      widths = [16, 10, 12, 14, 16, 12, 12, 8];
      rows = (data as ReportConsolidadoRow[]).map((r) => [
        r.equipe, String(r.total), String(r.done), String(r.inProgress), String(r.notStarted), String(r.late), String(r.workDays), `${r.avgPercent}%`,
      ]);
      break;
    }
    case "criticidade": {
      h = ["Tipo", "Título", "Nível", "Status", "Owner"];
      widths = [10, 50, 10, 15, 15];
      rows = (data as ReportCriticidadeRow[]).map((r) => [r.type, r.title, r.level, r.status, r.owner]);
      break;
    }
    case "comentarios": {
      h = ["Data", "Autor", "Tipo", "Item", "Comentário"];
      widths = [14, 14, 8, 22, 42];
      rows = (data as ReportCommentRow[]).map((r) => [r.date, r.author, r.type, r.target, r.body]);
      break;
    }
    case "issues-kpi": {
      h = ["Mês", "Abertas", "Fechadas", "Saldo"];
      widths = [25, 25, 25, 25];
      rows = (data as ReportIssuesKPIRow[]).map((r) => [r.month, String(r.opened), String(r.closed), String(r.open)]);
      break;
    }
    case "atrasadas": {
      h = ["WBS", "Nome", "Fim", "Dias atraso", "%", "Resp.", "Equipe"];
      widths = [7, 35, 10, 10, 6, 18, 14];
      rows = (data as ReportAtrasadasRow[]).map((r) => [r.wbs, r.name, r.end, String(r.daysLate), `${r.percent}%`, r.assignee, r.equipe]);
      break;
    }
  }

  const doc = React.createElement(ReportDoc, { title, header, headers: h, rows, widths });
  const stream = await pdf(doc).toBlob();
  const arrayBuffer = await stream.arrayBuffer();
  return Buffer.from(arrayBuffer);
}
