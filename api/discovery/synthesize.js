const { saveDiscoveryRun } = require('../_lib/db');
const { handleOptions, readJson, requireMethod, sendJson } = require('../_lib/http');
const { isMetadataOnlyDiscovery, localSynthesis } = require('../_lib/local-discovery');
const { synthesizeWithOpenAI } = require('../_lib/openai');

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

    let synthesis;
    if (isMetadataOnlyDiscovery(payload)) {
      synthesis = localSynthesis(payload, 'metadata-only source; native export required for deep lineage');
    } else {
      try {
        synthesis = await synthesizeWithOpenAI(payload);
      } catch (error) {
        synthesis = localSynthesis(
          payload,
          error instanceof Error ? `LLM fallback: ${error.message}` : 'LLM fallback: synthesis failed'
        );
      }
    }
    const persistence = await saveDiscoveryRun(payload, synthesis);

    sendJson(response, 200, {
      ok: true,
      runId: persistence.runId,
      stored: persistence.stored,
      persistenceError: persistence.persistenceError || null,
      counts: persistence.counts || null,
      model: synthesis.model,
      fallbackReason: synthesis.fallbackReason || null,
      outputText: synthesis.outputText,
      canonicalDelta: synthesis.canonicalDelta
    });
  } catch (error) {
    sendJson(response, error.statusCode || 500, {
      ok: false,
      error: error instanceof Error ? error.message : 'Discovery synthesis failed.',
      detail: error.detail || null
    });
  }
};
