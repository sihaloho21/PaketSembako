exports.handler = async function handler(event) {
  const gasBaseUrl = String(process.env.GAS_WEB_APP_URL || '').trim();
  const gasApiKey = String(process.env.GAS_API_KEY || '').trim();

  if (!gasBaseUrl) {
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
      },
      body: JSON.stringify({
        ok: false,
        error: 'GAS_WEB_APP_URL belum diset di Netlify Environment Variables.',
      }),
    };
  }

  const method = String(event.httpMethod || 'GET').toUpperCase();
  if (method === 'OPTIONS') {
    return {
      statusCode: 204,
      headers: {
        'Cache-Control': 'no-store',
      },
      body: '',
    };
  }

  let targetUrl;
  try {
    targetUrl = new URL(gasBaseUrl);
  } catch (err) {
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
      },
      body: JSON.stringify({
        ok: false,
        error: 'GAS_WEB_APP_URL tidak valid.',
      }),
    };
  }

  const query = event.queryStringParameters || {};
  Object.keys(query).forEach((key) => {
    if (query[key] !== undefined && query[key] !== null) {
      targetUrl.searchParams.set(key, String(query[key]));
    }
  });

  if (gasApiKey && !targetUrl.searchParams.get('apiKey')) {
    targetUrl.searchParams.set('apiKey', gasApiKey);
  }

  const requestHeaders = {
    Accept: 'application/json, text/plain;q=0.9, */*;q=0.8',
  };

  let requestBody;
  if (method !== 'GET' && method !== 'HEAD') {
    let payload = {};

    if (event.body) {
      const rawBody = event.isBase64Encoded
        ? Buffer.from(event.body, 'base64').toString('utf8')
        : event.body;
      try {
        payload = JSON.parse(rawBody);
      } catch (err) {
        payload = {};
      }
    }

    if (payload && typeof payload === 'object' && !Array.isArray(payload) && gasApiKey && !payload.apiKey) {
      payload.apiKey = gasApiKey;
    }

    requestBody = JSON.stringify(payload || {});
    requestHeaders['Content-Type'] = 'text/plain;charset=utf-8';
  }

  try {
    const upstream = await fetch(targetUrl.toString(), {
      method,
      headers: requestHeaders,
      body: requestBody,
      redirect: 'follow',
    });

    const text = await upstream.text();

    return {
      statusCode: upstream.status,
      headers: {
        'Content-Type': upstream.headers.get('content-type') || 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
      },
      body: text,
    };
  } catch (err) {
    return {
      statusCode: 502,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
      },
      body: JSON.stringify({
        ok: false,
        error: 'Proxy gagal menghubungi Apps Script: ' + (err && err.message ? err.message : String(err)),
      }),
    };
  }
};