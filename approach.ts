let LOAD_TEST_MODE = false;
let LOAD_TEST_REQUESTS = 15;

Office.onReady(() => {
  // ✅ Expose controls to F12 console
  (window as any).generateReport = generateAndDownloadReport;
  (window as any).clearMetrics = () => localStorage.removeItem("addin_metrics");

  // ✅ Expose load test flags
  (window as any).setLoadTestMode = (enabled: boolean) => {
    LOAD_TEST_MODE = enabled;
    console.log(`Load test mode: ${enabled ? "ON 🟢" : "OFF 🔴"}`);
  };
  (window as any).setLoadTestRequests = (count: number) => {
    LOAD_TEST_REQUESTS = count;
    console.log(`Load test requests set to: ${count}`);
  };
  (window as any).loadTestStatus = () => {
    console.log({
      loadTestMode: LOAD_TEST_MODE,
      loadTestRequests: LOAD_TEST_REQUESTS,
      metricsRecorded: JSON.parse(localStorage.getItem("addin_metrics") ?? "[]").length
    });
  };

  console.log(`
🔧 Addin Debug Controls Available:
  setLoadTestMode(true/false)   — toggle load test
  setLoadTestRequests(number)   — set request count
  loadTestStatus()              — check current config
  generateReport()              — download report
  clearMetrics()                — reset metrics
  `);
});



const LOAD_TEST_MODE = true;  // ✅ toggle off for prod
const LOAD_TEST_REQUESTS = 15;

Office.onReady(async () => {
  if (isProcessing) return;

  try {
    isProcessing = true;
    const token = await getToken();

    if (LOAD_TEST_MODE) {
      // ✅ Fire multiple Graph calls in sequence from single click
      showNotification("InformationalMessage", `⏳ Load testing — ${LOAD_TEST_REQUESTS} requests...`);

      for (let i = 0; i < LOAD_TEST_REQUESTS; i++) {
        await graphFetch(
          `https://graph.microsoft.com/beta/me/mailFolders/inbox/messages` +
          `?$filter=conversationId eq '${Office.context.mailbox.item?.conversationId}'` +
          `&$select=id,subject,from,receivedDateTime` +
          `&$top=10`,
          token
        );
        console.log(`Request ${i + 1}/${LOAD_TEST_REQUESTS} done`);
      }

      showNotification("InformationalMessage", "✅ Load test complete — run generateReport() in console");

    } else {
      // Normal flow
      const thread = await getMailThread(token);
      sendToAWS(thread, token).catch(console.error);
      showNotification("InformationalMessage", "✅ Mail processed successfully!");
    }

  } catch (err) {
    showNotification("ErrorMessage", "❌ Error: " + (err instanceof Error ? err.message : "Unknown"));
  } finally {
    isProcessing = false;
  }
});




