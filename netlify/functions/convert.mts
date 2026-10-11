/**
 * Creates a CloudConvert job for one file.
 *
 * The browser sends only the two formats. This function builds the job with
 * the key from Netlify's CLOUDCONVERT_API_KEY variable (mark it "Contains
 * secret values") and returns the signed upload form, so the file itself goes
 * straight from the browser to CloudConvert and never through Netlify.
 */

// The formats offered in the app's format menus (index.html and app.js).
const OUTPUT_FORMATS = new Set([
  'pdf', 'docx', 'doc', 'pptx', 'xlsx', 'rtf', 'txt',
  'odt', 'odp', 'ods', 'jpg', 'png', 'webp', 'gif', 'epub', 'mobi',
]);
const INPUT_FORMAT = /^[a-z0-9]{1,10}$/;

const json = (status: number, data: unknown) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });

// Browsers send an Origin header with every POST; it must be this site.
const isSameSite = (origin: string | null, host: string) => {
  if (!origin) return true;
  try {
    return new URL(origin).host === host;
  } catch {
    return false;
  }
};

export default async (req: Request) => {
  if (req.method !== 'POST') return json(405, { message: 'Use POST.' });

  // Block other websites from spending the credits through their visitors' browsers.
  if (!isSameSite(req.headers.get('origin'), new URL(req.url).host)) {
    return json(403, { message: 'Cross-origin requests are not allowed.' });
  }

  const key = process.env.CLOUDCONVERT_API_KEY;
  if (!key) return json(501, { message: 'Conversion service not configured.' });

  const body = await req.json().catch(() => null);
  const inputFormat = String(body?.inputFormat ?? '').toLowerCase();
  const outputFormat = String(body?.outputFormat ?? '').toLowerCase();
  if (!INPUT_FORMAT.test(inputFormat) || !OUTPUT_FORMATS.has(outputFormat)) {
    return json(400, { message: 'Unsupported file format.' });
  }

  try {
    const res = await fetch('https://api.cloudconvert.com/v2/jobs', {
      method: 'POST',
      headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tasks: {
          'upload-task': { operation: 'import/upload' },
          'convert-task': {
            operation: 'convert',
            input: 'upload-task',
            input_format: inputFormat,
            output_format: outputFormat,
          },
          'export-task': { operation: 'export/url', input: 'convert-task' },
        },
      }),
    });
    const data = await res.json().catch(() => ({}));

    // 422 means CloudConvert cannot do this conversion; its message says why.
    if (res.status === 422) return json(422, { message: data.message || 'This conversion is not supported.' });
    if (res.status === 429) return json(429, { message: 'The conversion service is busy. Try again in a minute.' });
    if (!res.ok) return json(502, { message: 'The conversion service could not start the job.' });

    const upload = data.data?.tasks?.find((t: { name: string }) => t.name === 'upload-task')?.result?.form;
    if (!upload) return json(502, { message: 'The conversion service did not return an upload form.' });
    return json(200, { id: data.data.id, upload: { url: upload.url, parameters: upload.parameters } });
  } catch {
    return json(502, { message: 'The conversion service could not be reached.' });
  }
};

export const config = {
  path: '/api/convert',
  // Each visitor can start 10 conversions per minute; extra requests get HTTP 429.
  rateLimit: { windowLimit: 10, windowSize: 60, aggregateBy: ['ip', 'domain'] },
};
