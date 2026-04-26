const { getRun, hasDatabase } = require('../_lib/db');
const { handleOptions, requireMethod, sendJson } = require('../_lib/http');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  if (!requireMethod(request, response, 'GET')) {
    return;
  }

  if (!hasDatabase()) {
    sendJson(response, 404, {
      ok: false,
      error: 'Database is not configured.'
    });
    return;
  }

  try {
    const url = new URL(request.url, 'http://127.0.0.1');
    const id = url.searchParams.get('id');
    if (!id) {
      sendJson(response, 400, {
        ok: false,
        error: 'id query parameter is required.'
      });
      return;
    }

    const run = await getRun(id);
    sendJson(response, run ? 200 : 404, {
      ok: Boolean(run),
      run
    });
  } catch (error) {
    sendJson(response, 500, {
      ok: false,
      error: error instanceof Error ? error.message : 'Could not load discovery run.'
    });
  }
};
