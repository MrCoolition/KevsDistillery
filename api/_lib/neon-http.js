const FIRST_HOST_LABEL = /^[^.]+\./;

const TYPE_PARSERS = new Map([
  [16, parseBoolean],
  [20, parseNumber],
  [21, parseNumber],
  [23, parseNumber],
  [26, parseNumber],
  [700, parseNumber],
  [701, parseNumber],
  [1700, parseNumber],
  [114, parseJson],
  [3802, parseJson]
]);

function endpointFor(hostname, jwtAuth = false) {
  if (hostname.startsWith('api.') || hostname.startsWith('apiauth.')) {
    return `https://${hostname}/sql`;
  }

  const prefix = jwtAuth ? 'apiauth.' : 'api.';
  return `https://${hostname.replace(FIRST_HOST_LABEL, prefix)}/sql`;
}

function parseBoolean(value) {
  if (typeof value === 'boolean') {
    return value;
  }

  return value === 't' || value === 'true';
}

function parseNumber(value) {
  if (typeof value === 'number') {
    return value;
  }

  const number = Number(value);
  return Number.isFinite(number) ? number : value;
}

function parseJson(value) {
  if (typeof value !== 'string') {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function normalizeParam(value) {
  if (value === undefined) {
    return null;
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return value;
}

function toParameterizedQuery(strings, values) {
  let query = '';
  const params = [];

  strings.forEach((part, index) => {
    query += part;
    if (index < values.length) {
      params.push(normalizeParam(values[index]));
      query += `$${params.length}`;
    }
  });

  return {
    query,
    params
  };
}

function processResult(result) {
  const fields = Array.isArray(result?.fields) ? result.fields : [];
  const rows = Array.isArray(result?.rows) ? result.rows : [];

  if (!fields.length) {
    return [];
  }

  return rows.map((row) => {
    if (!Array.isArray(row)) {
      return row;
    }

    return Object.fromEntries(fields.map((field, index) => {
      const parser = TYPE_PARSERS.get(field.dataTypeID);
      const value = row[index] === null ? null : (parser ? parser(row[index]) : row[index]);
      return [field.name, value];
    }));
  });
}

function createNeonHttpSql(connectionString) {
  const databaseUrl = new URL(connectionString);
  const endpoint = endpointFor(databaseUrl.hostname);

  async function execute(query, params = []) {
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Neon-Connection-String': connectionString,
        'Neon-Raw-Text-Output': 'true',
        'Neon-Array-Mode': 'true'
      },
      body: JSON.stringify({
        query,
        params: params.map(normalizeParam)
      })
    });

    const contentType = response.headers.get('content-type') || '';
    const body = contentType.includes('application/json') ? await response.json() : await response.text();

    if (!response.ok) {
      const message = typeof body === 'object' && body?.message
        ? body.message
        : `Neon HTTP query failed with status ${response.status}.`;
      const error = new Error(message);
      error.status = response.status;
      error.detail = body;
      throw error;
    }

    return processResult(body);
  }

  function sql(strings, ...values) {
    const prepared = toParameterizedQuery(strings, values);
    return execute(prepared.query, prepared.params);
  }

  sql.query = (query, params = []) => execute(query, params);
  sql.driver = 'neon-http';

  return sql;
}

module.exports = {
  createNeonHttpSql
};
