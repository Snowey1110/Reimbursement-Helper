export interface Env {
  REPORT_DB: D1Database;
  ALLOWED_ORIGINS?: string;
}

type FormVersion = "USA" | "Korea";
type AppLanguage = "en" | "zh" | "ko";

const VALID_FORMS = new Set<FormVersion>(["USA", "Korea"]);
const VALID_LANGUAGES = new Set<AppLanguage>(["en", "zh", "ko"]);

function allowedOrigins(env: Env): string[] {
  return (env.ALLOWED_ORIGINS || "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean);
}

function corsHeaders(request: Request, env: Env): HeadersInit {
  const origin = request.headers.get("Origin") || "";
  const allowed = allowedOrigins(env);
  const allowOrigin = allowed.includes(origin) ? origin : allowed[0] || "*";
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store"
  };
}

function json(request: Request, env: Env, payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...corsHeaders(request, env),
      "Content-Type": "application/json"
    }
  });
}

async function stats(env: Env) {
  const totalRow = await env.REPORT_DB.prepare("SELECT COALESCE(SUM(count), 0) AS total FROM report_counts").first<{ total: number }>();
  const byFormRows = await env.REPORT_DB.prepare("SELECT form_version, COALESCE(SUM(count), 0) AS count FROM report_counts GROUP BY form_version").all<{
    form_version: FormVersion;
    count: number;
  }>();
  const byLanguageRows = await env.REPORT_DB.prepare("SELECT language, COALESCE(SUM(count), 0) AS count FROM report_counts GROUP BY language").all<{
    language: AppLanguage;
    count: number;
  }>();
  const recentRows = await env.REPORT_DB.prepare(
    "SELECT day, COALESCE(SUM(count), 0) AS count FROM report_counts GROUP BY day ORDER BY day DESC LIMIT 30"
  ).all<{ day: string; count: number }>();

  return {
    total: Number(totalRow?.total ?? 0),
    byForm: Object.fromEntries((byFormRows.results || []).map((row) => [row.form_version, Number(row.count)])),
    byLanguage: Object.fromEntries((byLanguageRows.results || []).map((row) => [row.language, Number(row.count)])),
    recentDays: recentRows.results || []
  };
}

async function track(request: Request, env: Env): Promise<Response> {
  const body = (await request.json().catch(() => null)) as null | {
    event?: string;
    formVersion?: string;
    language?: string;
  };
  if (body?.event !== "report_generated") return json(request, env, { error: "Unsupported event." }, 400);
  const formVersion = body.formVersion === "Korea" ? "Korea" : body.formVersion === "USA" ? "USA" : undefined;
  const language = body.language === "zh" || body.language === "ko" ? body.language : body.language === "en" ? "en" : undefined;
  if (!formVersion || !language || !VALID_FORMS.has(formVersion) || !VALID_LANGUAGES.has(language)) {
    return json(request, env, { error: "Invalid tracker payload." }, 400);
  }

  const now = new Date();
  const day = now.toISOString().slice(0, 10);
  const updatedAt = now.toISOString();
  await env.REPORT_DB.prepare(
    `INSERT INTO report_counts (day, form_version, language, count, updated_at)
     VALUES (?, ?, ?, 1, ?)
     ON CONFLICT(day, form_version, language)
     DO UPDATE SET count = count + 1, updated_at = excluded.updated_at`
  )
    .bind(day, formVersion, language, updatedAt)
    .run();

  return json(request, env, await stats(env));
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (request.method === "OPTIONS") return new Response(null, { headers: corsHeaders(request, env) });
    const url = new URL(request.url);
    if (url.pathname === "/health") return json(request, env, { ok: true });
    if (url.pathname === "/stats" && request.method === "GET") return json(request, env, await stats(env));
    if (url.pathname === "/track" && request.method === "POST") return track(request, env);
    return json(request, env, { error: "Not found." }, 404);
  }
};
