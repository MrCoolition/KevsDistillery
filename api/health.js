const { databaseStatus } = require('./_lib/db');
const { handleOptions, sendJson } = require('./_lib/http');
const { model } = require('./_lib/openai');

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
      error: error instanceof Error ? error.message : 'Database check failed.'
    };
  }

  sendJson(response, 200, {
    ok: true,
    model,
    openAIConfigured: Boolean(process.env.OPENAI_API_KEY),
    database
  });
};
