import fetch from "node-fetch";

// ─── Config ──────────────────────────────────────────────────────────────────
const TOKEN = "PASTE_YOUR_TOKEN_HERE";
const CONVERSATION_ID = "PASTE_CONVERSATION_ID_HERE"; // from addin logs
const TOTAL_REQUESTS = 50;       // simulate 50 rapid clicks
const CONCURRENT_REQUESTS = 5;   // 5 at a time
const DELAY_BETWEEN_BATCHES = 500; // ms between batches

interface RequestResult {
  requestNumber: number;
  status: number;
  durationMs: number;
  retryAfter?: number;
  error?: string;
}

// ─── Single Graph Call ────────────────────────────────────────────────────────
async function makeGraphCall(requestNumber: number): Promise<RequestResult> {
  const start = Date.now();

  try {
    const response = await fetch(
      `https://graph.microsoft.com/beta/me/mailFolders/inbox/messages` +
      `?$filter=conversationId eq '${CONVERSATION_ID}'` +
      `&$select=id,subject,from,receivedDateTime` +
      `&$top=10`,
      {
        headers: {
          Authorization: `Bearer ${TOKEN}`,
          "Content-Type": "application/json",
        },
      }
    );

    const durationMs = Date.now() - start;
    const retryAfter = response.headers.get("Retry-After");

    return {
      requestNumber,
      status: response.status,
      durationMs,
      retryAfter: retryAfter ? parseInt(retryAfter) : undefined,
    };

  } catch (err) {
    return {
      requestNumber,
      status: 0,
      durationMs: Date.now() - start,
      error: err instanceof Error ? err.message : "Unknown error",
    };
  }
}

// ─── Run Test ─────────────────────────────────────────────────────────────────
async function runLoadTest(): Promise<void> {
  console.log(`\n🚀 Starting load test — ${TOTAL_REQUESTS} requests, ${CONCURRENT_REQUESTS} concurrent\n`);

  const results: RequestResult[] = [];
  const batches = Math.ceil(TOTAL_REQUESTS / CONCURRENT_REQUESTS);

  for (let batch = 0; batch < batches; batch++) {
    const batchStart = batch * CONCURRENT_REQUESTS;
    const batchEnd = Math.min(batchStart + CONCURRENT_REQUESTS, TOTAL_REQUESTS);
    const batchRequests = Array.from(
      { length: batchEnd - batchStart },
      (_, i) => makeGraphCall(batchStart + i + 1)
    );

    console.log(`Batch ${batch + 1}/${batches} — requests ${batchStart + 1} to ${batchEnd}`);
    const batchResults = await Promise.all(batchRequests);
    results.push(...batchResults);

    // Log each result immediately
    batchResults.forEach(r => {
      const icon = r.status === 200 ? "✅" : r.status === 429 ? "⚠️" : "❌";
      console.log(
        `  ${icon} Request ${r.requestNumber}: ` +
        `${r.status} | ${r.durationMs}ms` +
        (r.retryAfter ? ` | Retry-After: ${r.retryAfter}s` : "") +
        (r.error ? ` | Error: ${r.error}` : "")
      );
    });

    // Delay between batches
    if (batch < batches - 1) {
      await new Promise(res => setTimeout(res, DELAY_BETWEEN_BATCHES));
    }
  }

  generateReport(results);
}

// ─── Report ───────────────────────────────────────────────────────────────────
function generateReport(results: RequestResult[]): void {
  const successful = results.filter(r => r.status === 200);
  const throttled = results.filter(r => r.status === 429);
  const failed = results.filter(r => r.status !== 200 && r.status !== 429);

  const durations = successful.map(r => r.durationMs);
  const avgDuration = durations.reduce((a, b) => a + b, 0) / durations.length;
  const minDuration = Math.min(...durations);
  const maxDuration = Math.max(...durations);
  const p95Duration = durations.sort((a, b) => a - b)[Math.floor(durations.length * 0.95)];

  console.log(`
╔════════════════════════════════════════╗
║         GRAPH API LOAD TEST REPORT     ║
╚════════════════════════════════════════╝

📊 SUMMARY
──────────────────────────────────────
Total Requests      : ${results.length}
Successful (200)    : ${successful.length} (${((successful.length / results.length) * 100).toFixed(1)}%)
Throttled (429)     : ${throttled.length} (${((throttled.length / results.length) * 100).toFixed(1)}%)
Failed              : ${failed.length} (${((failed.length / results.length) * 100).toFixed(1)}%)

⏱️  RESPONSE TIMES (successful only)
──────────────────────────────────────
Average             : ${avgDuration.toFixed(0)}ms
Min                 : ${minDuration}ms
Max                 : ${maxDuration}ms
P95                 : ${p95Duration}ms

⚠️  THROTTLING
──────────────────────────────────────
429s encountered    : ${throttled.length}
Max Retry-After     : ${throttled.length > 0 ? Math.max(...throttled.map(r => r.retryAfter ?? 0)) : 0}s
First 429 at req    : ${throttled.length > 0 ? throttled[0].requestNumber : "N/A"}

🔧 TEST CONFIG
──────────────────────────────────────
Total Requests      : ${results.length}
Concurrent          : ${CONCURRENT_REQUESTS}
Delay between batch : ${DELAY_BETWEEN_BATCHES}ms
  `);

  // Save to JSON for further analysis
  const report = {
    summary: {
      total: results.length,
      successful: successful.length,
      throttled: throttled.length,
      failed: failed.length,
      successRate: `${((successful.length / results.length) * 100).toFixed(1)}%`,
    },
    responseTimes: {
      avg: `${avgDuration.toFixed(0)}ms`,
      min: `${minDuration}ms`,
      max: `${maxDuration}ms`,
      p95: `${p95Duration}ms`,
    },
    throttling: {
      count: throttled.length,
      firstOccurredAtRequest: throttled.length > 0 ? throttled[0].requestNumber : null,
      maxRetryAfterSeconds: throttled.length > 0 ? Math.max(...throttled.map(r => r.retryAfter ?? 0)) : 0,
    },
    rawResults: results,
  };

  require("fs").writeFileSync(
    "load-test-report.json",
    JSON.stringify(report, null, 2)
  );

  console.log("📁 Full report saved to load-test-report.json");
}

// ─── Run ──────────────────────────────────────────────────────────────────────
runLoadTest().catch(console.error);