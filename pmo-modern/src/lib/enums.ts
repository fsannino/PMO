// Enums "lógicos" do domínio. Como SQLite não suporta enums no Prisma, mantemos
// strings no banco e validamos no app via Zod / union types.
//
// Em produção (PostgreSQL), reintroduzir como enums Prisma nativos.

import { z } from "zod";

export const Roles = ["ADMIN", "MANAGER", "MEMBER", "VIEWER"] as const;
export type Role = (typeof Roles)[number];
export const RoleSchema = z.enum(Roles);

export const Modules = ["CUT", "GVI", "GRF", "TCP", "TIN", "PMO", "LIGHT"] as const;
export type Module = (typeof Modules)[number];
export const ModuleSchema = z.enum(Modules);

export const ProjectStatuses = ["ACTIVE", "PAUSED", "CLOSED", "ARCHIVED"] as const;
export type ProjectStatus = (typeof ProjectStatuses)[number];
export const ProjectStatusSchema = z.enum(ProjectStatuses);

export const TaskStatuses = ["NOT_STARTED", "IN_PROGRESS", "DELAYED", "DONE", "CANCELLED"] as const;
export type TaskStatus = (typeof TaskStatuses)[number];
export const TaskStatusSchema = z.enum(TaskStatuses);

export const DependencyTypes = ["FS", "SS", "FF", "SF"] as const;
export type DependencyType = (typeof DependencyTypes)[number];
export const DependencyTypeSchema = z.enum(DependencyTypes);

export const HistoryActions = ["CREATED", "UPDATED", "DELETED", "RESTORED", "IMPORTED", "MEASURED"] as const;
export type HistoryAction = (typeof HistoryActions)[number];
export const HistoryActionSchema = z.enum(HistoryActions);

export const Severities = ["LOW", "MEDIUM", "HIGH", "CRITICAL"] as const;
export type Severity = (typeof Severities)[number];
export const SeveritySchema = z.enum(Severities);

export const Priorities = ["LOW", "MEDIUM", "HIGH", "CRITICAL"] as const;
export type Priority = (typeof Priorities)[number];
export const PrioritySchema = z.enum(Priorities);

export const IssueStatuses = ["OPEN", "IN_PROGRESS", "RESOLVED", "CLOSED"] as const;
export type IssueStatus = (typeof IssueStatuses)[number];
export const IssueStatusSchema = z.enum(IssueStatuses);

export const RiskStatuses = ["IDENTIFIED", "ANALYZING", "MITIGATING", "ACCEPTED", "CLOSED"] as const;
export type RiskStatus = (typeof RiskStatuses)[number];
export const RiskStatusSchema = z.enum(RiskStatuses);

export const ActionStatuses = ["OPEN", "IN_PROGRESS", "DONE", "CANCELLED"] as const;
export type ActionStatus = (typeof ActionStatuses)[number];
export const ActionStatusSchema = z.enum(ActionStatuses);

export const ChangeRequestStatuses = [
  "OPEN",
  "UNDER_REVIEW",
  "APPROVED",
  "IN_IMPLEMENTATION",
  "IMPLEMENTED",
  "REJECTED",
  "CANCELLED",
] as const;
export type ChangeRequestStatus = (typeof ChangeRequestStatuses)[number];
export const ChangeRequestStatusSchema = z.enum(ChangeRequestStatuses);

export const ImportSources = ["EXCEL", "MSPROJECT_XML", "MANUAL"] as const;
export type ImportSource = (typeof ImportSources)[number];
export const ImportSourceSchema = z.enum(ImportSources);

export const ImportTemplates = ["IS", "RK", "AC", "CR", "TK", "TC"] as const;
export type ImportTemplate = (typeof ImportTemplates)[number];
export const ImportTemplateSchema = z.enum(ImportTemplates);

export const ImportModes = ["REPLACE", "MERGE", "ONLY_NEW"] as const;
export type ImportMode = (typeof ImportModes)[number];
export const ImportModeSchema = z.enum(ImportModes);
