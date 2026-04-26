const { saveDiscoveryRun } = require('../_lib/db');
const { handleOptions, readJson, requireMethod, sendJson } = require('../_lib/http');
const { retrieveBackgroundSynthesis } = require('../_lib/openai');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  if (!requireMethod(request, response, 'POST')) {
    return;
  }

  try {
    const payload = await readJson(request);
    const { responseId, sourceKind, sourceName, extractedText } = payload;

    if (!responseId) {
      sendJson(response, 400, {
        ok: false,
        error: 'responseId is required.'
      });
      return;
    }

    const synthesis = await retrieveBackgroundSynthesis(responseId);
    if (synthesis.pending) {
      sendJson(response, 202, {
        ok: true,
        queued: true,
        responseId: synthesis.responseId,
        responseStatus: synthesis.responseStatus,
        engine: 'The Distillery',
        message: `The Distillery run is ${synthesis.responseStatus}.`
      });
      return;
    }

    if (!sourceKind || !sourceName || !extractedText) {
      sendJson(response, 200, {
        ok: true,
        queued: false,
        needsPayload: true,
        responseId: synthesis.responseId,
        responseStatus: synthesis.responseStatus,
        engine: 'The Distillery',
        message: 'The Distillery run is complete. Send the source payload once to save the run.',
        outputText: synthesis.outputText,
        canonicalDelta: synthesis.canonicalDelta
      });
      return;
    }

    const persistence = await saveDiscoveryRun(payload, synthesis);
    sendJson(response, 200, {
      ok: true,
      queued: false,
      responseId: synthesis.responseId,
      responseStatus: synthesis.responseStatus,
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
      error: error?.statusCode && error instanceof Error ? error.message : 'Could not check Distillery run status.'
    });
  }
};
