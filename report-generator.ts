function generateAndDownloadReport(): void {
  const successful = metrics.filter(m => m.status === 200);
  const throttled = metrics.filter(m => m.status === 429);
  const failed = metrics.filter(m => m.status !== 200 && m.status !== 429);
  const durations = successful.map(m => m.durationMs).sort((a, b) => a - b);
  const avg = durations.reduce((a, b) => a + b, 0) / durations.length;
  const p95 = durations[Math.floor(durations.length * 0.95)];

  const report = {
    generatedAt: new Date().toISOString(),
    summary: {
      totalRequests: metrics.length,
      successful: successful.length,
      throttled: throttled.length,
      failed: failed.length,
      successRate: `${((successful.length / metrics.length) * 100).toFixed(1)}%`,
    },
    responseTimes: {
      avgMs: Math.round(avg),
      minMs: durations[0],
      maxMs: durations[durations.length - 1],
      p95Ms: p95,
    },
    throttling: {
      count: throttled.length,
      firstOccurredAtRequest: throttled[0]?.requestNumber ?? null,
      maxRetryAfterSeconds: throttled.length > 0
        ? Math.max(...throttled.map(m => m.retryAfter ?? 0))
        : 0,
    },
    rawMetrics: metrics,
  };

  // ✅ Download as JSON file from browser
  const blob = new Blob(
    [JSON.stringify(report, null, 2)],
    { type: "application/json" }
  );
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `graph-load-test-${new Date().toISOString()}.json`;
  a.click();
  URL.revokeObjectURL(url);

  console.log("📊 Report:", JSON.stringify(report.summary, null, 2));
}