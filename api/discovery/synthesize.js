const { saveDiscoveryRun } = require('../_lib/db');
const { handleOptions, readJson, requireMethod, sendJson } = require('../_lib/http');
const { startBackgroundSynthesis } = require('../_lib/openai');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  if (!requireMethod(request, response, 'POST')) {
    return;
  }

  try {
    const payload = await readJson(request);
    const { sourceKind, sourceName, extractedText } = payload;

    if (!sourceKind || !sourceName || !extractedText) {
      sendJson(response, 400, {
        error: 'sourceKind, sourceName, and extractedText are required.'
      });
      return;
    }

    const synthesis = await startBackgroundSynthesis(payload);
    if (synthesis.pending) {
      sendJson(response, 202, {
        ok: true,
        queued: true,
        responseId: synthesis.responseId,
        responseStatus: synthesis.responseStatus,
        engine: 'The Distillery',
        message: 'The Distillery run started. Poll /api/discovery/status for completion.'
      });
      return;
    }

    const persistence = await saveDiscoveryRun(payload, synthesis);

    sendJson(response, 200, {
      ok: true,
      runId: persistence.runId,
      stored: persistence.stored,
      persistenceError: persistence.persistenceError || null,
      counts: persistence.counts || null,
      engine: 'The Distillery',
      fallbackReason: synthesis.fallbackReason || null,
      outputText: synthesis.outputText,
      canonicalDelta: synthesis.canonicalDelta
    });
  } catch (error) {
    sendJson(response, error.statusCode || 500, {
      ok: false,
      error: error?.statusCode && error instanceof Error ? error.message : 'The Distillery could not finish this run.'
    });
  }
};
