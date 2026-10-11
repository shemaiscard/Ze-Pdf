/**
 * Reports the state of a CloudConvert job created by convert.mts.
 *
 * Returns only what the app needs: the status, the download link once the job
 * has finished, or the error message if it failed.
 */

const JOB_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const json = (status: number, data: unknown) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });

const isSameSite = (origin: string | null, host: string) => {
  if (!origin) return true;
  try {
    return new URL(origin).host === host;
  } catch {
    return false;
  }
};

type Task = { name: string; status: string; message?: string; result?: { files?: { url: string }[] } };

export default async (req: Request) => {
  if (req.method !== 'GET') return json(405, { message: 'Use GET.' });

  const url = new URL(req.url);
  if (!isSameSite(req.headers.get('origin'), url.host)) {
    return json(403, { message: 'Cross-origin requests are not allowed.' });
  }

  const key = process.env.CLOUDCONVERT_API_KEY;
  if (!key) return json(501, { message: 'Conversion service not configured.' });

  const id = url.searchParams.get('id') ?? '';
  if (!JOB_ID.test(id)) return json(400, { message: 'Invalid job id.' });

  try {
    const res = await fetch(`https://api.cloudconvert.com/v2/jobs/${id}`, {
      headers: { Authorization: `Bearer ${key}` },
    });
    if (res.status === 404) return json(404, { message: 'Job not found.' });
    if (!res.ok) return json(502, { message: 'The conversion service could not be reached.' });

    const job = (await res.json()).data as { status: string; tasks: Task[] };
    if (job.status === 'finished') {
      const file = job.tasks.find(t => t.name === 'export-task')?.result?.files?.[0];
      if (!file) return json(502, { message: 'The converted file is missing.' });
      return json(200, { status: 'finished', url: file.url });
    }
    if (job.status === 'error') {
      const failed = job.tasks.find(t => t.status === 'error');
      return json(200, { status: 'error', message: failed?.message || 'Conversion error on server.' });
    }
    return json(200, { status: job.status });
  } catch {
    return json(502, { message: 'The conversion service could not be reached.' });
  }
};

export const config = {
  path: '/api/convert-status',
  // The app polls every 2 seconds while a file converts, so this limit is higher.
  rateLimit: { windowLimit: 60, windowSize: 60, aggregateBy: ['ip', 'domain'] },
};
