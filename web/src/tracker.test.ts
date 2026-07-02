import { afterEach, describe, expect, it, vi } from "vitest";
import { loadTrackerConfig, loadTrackerStats, trackReportGenerated } from "./tracker";

describe("report tracker", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("stays disabled when the public config has no endpoint", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response(JSON.stringify({ enabled: false, endpoint: "" }), { status: 200 }))
    );

    await expect(loadTrackerConfig()).resolves.toBeNull();
  });

  it("loads stats and tracks generated reports", async () => {
    const fetchMock = vi.fn(async (url: string, options?: RequestInit) => {
      if (url.endsWith("tracker-config.json")) {
        return new Response(JSON.stringify({ enabled: true, endpoint: "https://tracker.example.com/" }), { status: 200 });
      }
      if (url === "https://tracker.example.com/stats") {
        return new Response(JSON.stringify({ total: 2, byForm: { USA: 1, Korea: 1 }, byLanguage: { zh: 2 } }), { status: 200 });
      }
      if (url === "https://tracker.example.com/track" && options?.method === "POST") {
        return new Response(JSON.stringify({ total: 3, byForm: { USA: 2, Korea: 1 }, byLanguage: { zh: 3 } }), { status: 200 });
      }
      return new Response("not found", { status: 404 });
    });
    vi.stubGlobal("fetch", fetchMock);

    const config = await loadTrackerConfig();
    expect(config).toEqual({ enabled: true, endpoint: "https://tracker.example.com" });
    await expect(loadTrackerStats(config!)).resolves.toMatchObject({ total: 2, byForm: { USA: 1, Korea: 1 } });
    await expect(trackReportGenerated(config, "USA", "zh")).resolves.toMatchObject({ total: 3, byForm: { USA: 2, Korea: 1 } });

    const trackBody = JSON.parse(String(fetchMock.mock.calls.find(([url]) => url === "https://tracker.example.com/track")?.[1]?.body));
    expect(trackBody).toMatchObject({ event: "report_generated", formVersion: "USA", language: "zh" });
  });
});
