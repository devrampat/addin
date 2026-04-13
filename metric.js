interface RequestMetric {
  requestNumber: number;
  status: number;
  durationMs: number;
  retryAfter?: number;
  timestamp: string;
  error?: string;
}

const metrics: RequestMetric[] = [];
let requestCounter = 0;

function recordMetric(metric: RequestMetric): void {
  metrics.push(metric);
  console.log(
    `[METRIC] #${metric.requestNumber} | ` +
    `${metric.status} | ` +
    `${metric.durationMs}ms` +
    (metric.retryAfter ? ` | Retry-After: ${metric.retryAfter}s` : "") +
    (metric.error ? ` | Error: ${metric.error}` : "")
  );
}


async function graphFetch<T>(
  url: string,
  token: string,
  retries = 3
): Promise<T> {
  for (let i = 0; i < retries; i++) {
    const start = Date.now();
    requestCounter++;
    const currentRequest = requestCounter;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    const durationMs = Date.now() - start;

    // ✅ 429 — record and retry
    if (response.status === 429) {
      const retryAfter = parseInt(response.headers.get("Retry-After") ?? "5", 10);
      recordMetric({
        requestNumber: currentRequest,
        status: 429,
        durationMs,
        retryAfter,
        timestamp: new Date().toISOString(),
      });
      await delay(retryAfter * 1000);
      continue;
    }

    // ✅ 401 CAE challenge
    if (response.status === 401) {
      const wwwAuth = response.headers.get("WWW-Authenticate");
      const claimsMatch = wwwAuth?.match(/claims="([^"]+)"/);
      recordMetric({
        requestNumber: currentRequest,
        status: 401,
        durationMs,
        timestamp: new Date().toISOString(),
        error: claimsMatch ? "CAE claims challenge" : "Auth failed",
      });
      if (claimsMatch) {
        token = await getTokenWithClaims(claimsMatch[1]);
        continue;
      }
      throw new Error("Authentication failed.");
    }

    if (!response.ok) {
      const err = await response.json();
      recordMetric({
        requestNumber: currentRequest,
        status: response.status,
        durationMs,
        timestamp: new Date().toISOString(),
        error: err?.error?.message,
      });
      throw new Error(err?.error?.message ?? "Graph API call failed");
    }

    // ✅ Success
    recordMetric({
      requestNumber: currentRequest,
      status: 200,
      durationMs,
      timestamp: new Date().toISOString(),
    });

    return response.json() as Promise<T>;
  }

  throw new Error("Throttled — max retries reached.");
}

Office.onReady(() => {
  document.addEventListener("keydown", (e) => {
    if (e.shiftKey && e.key === "R") {
      generateAndDownloadReport();
    }
  });
});