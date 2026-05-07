// Re-exporta os enums nativos do Prisma + Zod schemas para validação no app.
// (Em PostgreSQL temos enums nativos. As constantes []const ainda servem para
// iterações em UI — selects, dropdowns.)

import { z } from "zod";
import {
  Role as PrismaRole,
  Module as PrismaModule,
  ProjectStatus as PrismaProjectStatus,
  TaskStatus as PrismaTaskStatus,
  DependencyType as PrismaDependencyType,
  HistoryAction as PrismaHistoryAction,
  Severity as PrismaSeverity,
  Priority as PrismaPriority,
  IssueStatus as PrismaIssueStatus,
  RiskStatus as PrismaRiskStatus,
  ActionStatus as PrismaActionStatus,
  ChangeRequestStatus as PrismaChangeRequestStatus,
  ImportSource as PrismaImportSource,
  ImportTemplate as PrismaImportTemplate,
  ImportMode as PrismaImportMode,
} from "@prisma/client";

export type Role = PrismaRole;
export const Roles = Object.values(PrismaRole) as readonly Role[];
export const RoleSchema = z.nativeEnum(PrismaRole);

export type Module = PrismaModule;
export const Modules = Object.values(PrismaModule) as readonly Module[];
export const ModuleSchema = z.nativeEnum(PrismaModule);

export type ProjectStatus = PrismaProjectStatus;
export const ProjectStatuses = Object.values(PrismaProjectStatus) as readonly ProjectStatus[];
export const ProjectStatusSchema = z.nativeEnum(PrismaProjectStatus);

export type TaskStatus = PrismaTaskStatus;
export const TaskStatuses = Object.values(PrismaTaskStatus) as readonly TaskStatus[];
export const TaskStatusSchema = z.nativeEnum(PrismaTaskStatus);

export type DependencyType = PrismaDependencyType;
export const DependencyTypes = Object.values(PrismaDependencyType) as readonly DependencyType[];
export const DependencyTypeSchema = z.nativeEnum(PrismaDependencyType);

export type HistoryAction = PrismaHistoryAction;
export const HistoryActions = Object.values(PrismaHistoryAction) as readonly HistoryAction[];
export const HistoryActionSchema = z.nativeEnum(PrismaHistoryAction);

export type Severity = PrismaSeverity;
export const Severities = Object.values(PrismaSeverity) as readonly Severity[];
export const SeveritySchema = z.nativeEnum(PrismaSeverity);

export type Priority = PrismaPriority;
export const Priorities = Object.values(PrismaPriority) as readonly Priority[];
export const PrioritySchema = z.nativeEnum(PrismaPriority);

export type IssueStatus = PrismaIssueStatus;
export const IssueStatuses = Object.values(PrismaIssueStatus) as readonly IssueStatus[];
export const IssueStatusSchema = z.nativeEnum(PrismaIssueStatus);

export type RiskStatus = PrismaRiskStatus;
export const RiskStatuses = Object.values(PrismaRiskStatus) as readonly RiskStatus[];
export const RiskStatusSchema = z.nativeEnum(PrismaRiskStatus);

export type ActionStatus = PrismaActionStatus;
export const ActionStatuses = Object.values(PrismaActionStatus) as readonly ActionStatus[];
export const ActionStatusSchema = z.nativeEnum(PrismaActionStatus);

export type ChangeRequestStatus = PrismaChangeRequestStatus;
export const ChangeRequestStatuses = Object.values(PrismaChangeRequestStatus) as readonly ChangeRequestStatus[];
export const ChangeRequestStatusSchema = z.nativeEnum(PrismaChangeRequestStatus);

export type ImportSource = PrismaImportSource;
export const ImportSources = Object.values(PrismaImportSource) as readonly ImportSource[];
export const ImportSourceSchema = z.nativeEnum(PrismaImportSource);

export type ImportTemplate = PrismaImportTemplate;
export const ImportTemplates = Object.values(PrismaImportTemplate) as readonly ImportTemplate[];
export const ImportTemplateSchema = z.nativeEnum(PrismaImportTemplate);

export type ImportMode = PrismaImportMode;
export const ImportModes = Object.values(PrismaImportMode) as readonly ImportMode[];
export const ImportModeSchema = z.nativeEnum(PrismaImportMode);
