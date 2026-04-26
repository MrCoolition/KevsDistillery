const { hasDatabase, listRuns } = require('../_lib/db');
const { handleOptions, requireAuth, requireMethod, sendJson } = require('../_lib/http');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  if (!requireMethod(request, response, 'GET') || !requireAuth(request, response)) {
    return;
  }

  if (!hasDatabase()) {
    sendJson(response, 200, {
      ok: true,
      runs: [],
      databaseConfigured: false
    });
    return;
  }

  try {
    const url = new URL(request.url, 'http://127.0.0.1');
    const limit = Math.min(Number(url.searchParams.get('limit') || 20), 50);
    const runs = await listRuns(limit);
    sendJson(response, 200, {
      ok: true,
      databaseConfigured: true,
      runs
    });
  } catch (error) {
    sendJson(response, 500, {
      ok: false,
      error: error instanceof Error ? error.message : 'Could not load discovery runs.'
    });
  }
};
