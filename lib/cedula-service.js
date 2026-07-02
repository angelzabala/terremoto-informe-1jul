const INDEX_URL = 'https://www.sistemaspnp.com/cedula/index.php';
const RESULT_URL = 'https://www.sistemaspnp.com/cedula/resultado.php';
const REQUEST_DELAY_MS = 2500;
const MAX_RETRIES = 5;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function solveCaptcha(question) {
  const expr = question
    .replace(/×/g, '*')
    .replace(/÷/g, '/')
    .replace(/,/g, '.')
    .trim();

  const match = expr.match(/^(\d+(?:\.\d+)?)\s*([+\-*/])\s*(\d+(?:\.\d+)?)$/);
  if (!match) {
    throw new Error(`Captcha no reconocido: ${question}`);
  }

  const left = Number(match[1]);
  const operator = match[2];
  const right = Number(match[3]);

  switch (operator) {
    case '+':
      return left + right;
    case '-':
      return left - right;
    case '*':
      return left * right;
    case '/':
      return right === 0 ? NaN : left / right;
    default:
      throw new Error(`Operador no soportado: ${operator}`);
  }
}

function stripHtml(text) {
  return text.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function parseFieldMap(html) {
  const fields = {};
  const paragraphs = html.match(/<p[^>]*>[\s\S]*?<\/p>/gi) ?? [];

  for (const paragraph of paragraphs) {
    const text = stripHtml(paragraph);
    if (!text || text === 'Nueva búsqueda') continue;

    const separator = text.indexOf(':');
    if (separator === -1) continue;

    const label = text.slice(0, separator).trim();
    let value = text.slice(separator + 1).trim();

    if (label === 'Cédula' || label === 'Cedula') {
      const digits = value.match(/\d+/);
      value = digits ? digits[0] : value;
    }

    fields[label] = value;
  }

  return fields;
}

function emptyResult(status) {
  return {
    status,
    rif: '',
    primerApellido: '',
    segundoApellido: '',
    nombres: '',
    retryable: false,
  };
}

function parseConsultResult(html) {
  if (/RECORD_NOT_FOUND/i.test(html)) {
    return emptyResult('NO ENCONTRADO');
  }

  const alertMatch = html.match(/class=['"]alert alert-danger['"][^>]*>([\s\S]*?)<\/div>/i);
  if (alertMatch) {
    const errorText = stripHtml(alertMatch[1]);
    if (/RECORD_NOT_FOUND/i.test(errorText)) {
      return emptyResult('NO ENCONTRADO');
    }

    return {
      status: 'ERROR',
      error: errorText,
      rif: '',
      primerApellido: '',
      segundoApellido: '',
      nombres: '',
      retryable: /CAPTCHA|sesi[oó]n/i.test(errorText),
    };
  }

  const fields = parseFieldMap(html);
  const hasPersonalData = Boolean(
    fields['Primer Apellido'] ||
      fields.Nombres ||
      fields.RIF,
  );

  if (!hasPersonalData) {
    return {
      status: 'ERROR',
      error: 'Respuesta sin datos personales',
      rif: '',
      primerApellido: '',
      segundoApellido: '',
      nombres: '',
      retryable: true,
    };
  }

  return {
    status: 'ENCONTRADO',
    rif: fields.RIF ?? '',
    primerApellido: fields['Primer Apellido'] ?? '',
    segundoApellido: fields['Segundo Apellido'] ?? '',
    nombres: fields.Nombres ?? '',
    retryable: false,
  };
}

async function fetchFormPage() {
  const response = await fetch(INDEX_URL, {
    headers: {
      'User-Agent': 'consultar-cedulas/1.0',
    },
  });

  if (!response.ok) {
    throw new Error(`No se pudo cargar el formulario (${response.status})`);
  }

  const html = await response.text();
  const cookies = (response.headers.getSetCookie?.() ?? [])
    .map((cookie) => cookie.split(';')[0])
    .join('; ');

  const captchaMatch = html.match(/CAPTCHA:\s*¿Cuánto es\s*(.+?)\?/i);
  if (!captchaMatch) {
    throw new Error('No se encontró el captcha en la página');
  }

  return {
    cookies,
    captchaQuestion: captchaMatch[1].trim(),
  };
}

async function consultCedula(cedula) {
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt += 1) {
    try {
      const form = await fetchFormPage();
      const captchaAnswer = solveCaptcha(form.captchaQuestion);

      const response = await fetch(RESULT_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Referer: INDEX_URL,
          Cookie: form.cookies,
          'User-Agent': 'consultar-cedulas/1.0',
        },
        body: new URLSearchParams({
          cedula,
          captcha: String(Math.round(captchaAnswer)),
          jeje: '',
        }),
      });

      if (!response.ok) {
        throw new Error(`Consulta fallida (${response.status})`);
      }

      const html = await response.text();
      const parsed = parseConsultResult(html);

      if (parsed.retryable && attempt < MAX_RETRIES) {
        throw new Error(parsed.error ?? 'Respuesta inválida, reintentando');
      }

      return parsed;
    } catch (error) {
      lastError = error;
      if (attempt < MAX_RETRIES) {
        await sleep(REQUEST_DELAY_MS);
      }
    }
  }

  return {
    status: 'ERROR',
    error: lastError?.message ?? 'Error desconocido',
    rif: '',
    primerApellido: '',
    segundoApellido: '',
    nombres: '',
    retryable: false,
  };
}

module.exports = {
  consultCedula,
  REQUEST_DELAY_MS,
};
