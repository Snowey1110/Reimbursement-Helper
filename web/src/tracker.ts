import type { FormVersion } from "./types";
import type { AppLanguage } from "./i18n";

export interface ReportTrackerConfig {
  enabled: boolean;
  endpoint: string;
}

export interface ReportTrackerStats {
  total: number;
  byForm: Record<FormVersion, number>;
  byLanguage: Partial<Record<AppLanguage, number>>;
}

function normalizeEndpoint(value: unknown): string {
  if (typeof value !== "string") return "";
  return value.replace(/\/+$/, "");
}

function statsFromPayload(payload: any): ReportTrackerStats | null {
  if (!payload || typeof payload.total !== "number") return null;
  return {
    total: payload.total,
    byForm: {
      USA: Number(payload.byForm?.USA ?? 0),
      Korea: Number(payload.byForm?.Korea ?? 0)
    },
    byLanguage: {
      en: Number(payload.byLanguage?.en ?? 0),
      zh: Number(payload.byLanguage?.zh ?? 0),
      ko: Number(payload.byLanguage?.ko ?? 0)
    }
  };
}

export async function loadTrackerConfig(): Promise<ReportTrackerConfig | null> {
  if (typeof fetch !== "function") return null;
  try {
    const response = await fetch("./tracker-config.json", { cache: "no-store" });
    if (!response.ok) return null;
    const parsed = await response.json();
    const endpoint = normalizeEndpoint(parsed.endpoint);
    return parsed.enabled === true && endpoint ? { enabled: true, endpoint } : null;
  } catch {
    return null;
  }
}

export async function loadTrackerStats(config: ReportTrackerConfig): Promise<ReportTrackerStats | null> {
  try {
    const response = await fetch(`${config.endpoint}/stats`, { cache: "no-store" });
    if (!response.ok) return null;
    return statsFromPayload(await response.json());
  } catch {
    return null;
  }
}

export async function trackReportGenerated(config: ReportTrackerConfig | null, formVersion: FormVersion, language: AppLanguage): Promise<ReportTrackerStats | null> {
  if (!config) return null;
  try {
    const response = await fetch(`${config.endpoint}/track`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        event: "report_generated",
        formVersion,
        language
      }),
      keepalive: true
    });
    if (!response.ok) return null;
    return statsFromPayload(await response.json());
  } catch {
    return null;
  }
}
