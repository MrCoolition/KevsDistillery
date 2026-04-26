const { databaseStatus, ensureSchema } = require('../_lib/db');
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
    const status = ready ? await databaseStatus() : null;
    sendJson(response, ready ? 200 : 400, {
      ok: ready,
      workspaceReady: ready,
      workspace: status ? {
        configured: Boolean(status.configured),
        ready: Boolean(status.ready),
        tableCount: status.tableCount,
        requiredTableCount: status.requiredTableCount,
        missingTables: status.missingTables,
        error: status.error || null
      } : null,
      error: ready ? null : 'Workspace storage is not configured.'
    });
  } catch (error) {
    sendJson(response, 500, {
      ok: false,
      error: 'Workspace verification failed.'
    });
  }
};
