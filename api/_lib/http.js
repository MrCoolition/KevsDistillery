function setHeaders(response, statusCode) {
  response.statusCode = statusCode;
  response.setHeader('Content-Type', 'application/json');
  response.setHeader('Cache-Control', 'no-store');
  response.setHeader('Access-Control-Allow-Origin', process.env.DISTILLERY_ALLOWED_ORIGIN || '*');
  response.setHeader('Access-Control-Allow-Headers', 'content-type');
  response.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
}

function sendJson(response, statusCode, payload) {
  setHeaders(response, statusCode);
  response.end(JSON.stringify(payload));
}

function handleOptions(request, response) {
  if (request.method !== 'OPTIONS') {
    return false;
  }

  setHeaders(response, 204);
  response.end();
  return true;
}

async function readJson(request, maxBytes = 900000) {
  const chunks = [];
  let size = 0;

  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBytes) {
      const error = new Error('Request body is too large.');
      error.statusCode = 413;
      throw error;
    }
    chunks.push(chunk);
  }

  const raw = Buffer.concat(chunks).toString('utf8');
  return raw ? JSON.parse(raw) : {};
}

function requireMethod(request, response, method) {
  if (request.method === method) {
    return true;
  }

  sendJson(response, 405, {
    error: `Method ${request.method} is not allowed. Use ${method}.`
  });
  return false;
}

module.exports = {
  handleOptions,
  readJson,
  requireMethod,
  sendJson
};
