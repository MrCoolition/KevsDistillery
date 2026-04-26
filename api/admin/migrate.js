const { ensureSchema } = require('../_lib/db');
const { handleOptions, requireMethod, sendJson } = require('../_lib/http');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  if (!requireMethod(request, response, 'POST')) {
    return;
  }

  try {
    const ready = await ensureSchema();
    sendJson(response, ready ? 200 : 400, {
      ok: ready,
      databaseReady: ready,
      error: ready ? null : 'DATABASE_URL, POSTGRES_URL, or NEON_DATABASE_URL is not configured.'
    });
  } catch (error) {
    sendJson(response, 500, {
      ok: false,
      error: error instanceof Error ? error.message : 'Migration failed.'
    });
  }
};
