const { databaseStatus } = require('./_lib/db');
const { handleOptions, sendJson } = require('./_lib/http');

module.exports = async function handler(request, response) {
  if (handleOptions(request, response)) {
    return;
  }

  let database;

  try {
    database = await databaseStatus();
  } catch (error) {
    database = {
      configured: true,
      ready: false,
      error: 'Workspace check failed.'
    };
  }

  sendJson(response, 200, {
    ok: true,
    distillery: {
      ready: Boolean(process.env.OPENAI_API_KEY),
      label: 'The Distillery'
    },
    workspace: {
      configured: Boolean(database.configured),
      ready: Boolean(database.ready),
      tableCount: database.tableCount,
      requiredTableCount: database.requiredTableCount,
      missingTables: database.missingTables,
      error: database.error || null
    }
  });
};
