// functions/api/bestand/[id].js
//
// Cloudflare Pages Function voor bijlagen bij agendapunten. De bestanden staan in een R2-bucket
// (binding: BESTANDEN); de naam en omschrijving staan in de vergadering zelf (D1).
//   PUT    /api/bestand/<id>   -> bestand opslaan (ruwe bestandsinhoud als body, naam in header X-Bestandsnaam)
//   GET    /api/bestand/<id>   -> bestand ophalen
//   DELETE /api/bestand/<id>   -> bestand verwijderen

const MAX_BYTES = 25 * 1024 * 1024; // 25 MB
const TOEGESTAAN = ['pdf', 'ppt', 'pptx', 'doc', 'docx', 'xls', 'xlsx', 'jpg', 'jpeg', 'png', 'txt', 'csv'];

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}
function checkAuth(request, env) {
  if (!env.API_SECRET) return true;
  return request.headers.get('X-Api-Key') === env.API_SECRET;
}
function geldigId(id) {
  return typeof id === 'string' && /^[A-Za-z0-9-]{8,64}$/.test(id);
}
function veiligeNaam(naam) {
  return String(naam || 'bestand').replace(/[\r\n"\\\/]/g, '_').slice(0, 150);
}

export async function onRequestPut({ params, env, request }) {
  if (!checkAuth(request, env)) return json({ fout: 'niet geautoriseerd' }, 401);
  if (!env.BESTANDEN) return json({ fout: 'R2-bucket (BESTANDEN) is niet gekoppeld aan het project' }, 500);
  if (!geldigId(params.id)) return json({ fout: 'ongeldig id' }, 400);
  const naam = veiligeNaam(decodeURIComponent(request.headers.get('X-Bestandsnaam') || 'bestand'));
  const ext = (naam.split('.').pop() || '').toLowerCase();
  if (!TOEGESTAAN.includes(ext)) return json({ fout: 'bestandstype niet toegestaan' }, 415);
  const lengte = parseInt(request.headers.get('Content-Length') || '0', 10);
  if (lengte > MAX_BYTES) return json({ fout: 'bestand is groter dan 25 MB' }, 413);
  const inhoud = await request.arrayBuffer();
  if (inhoud.byteLength > MAX_BYTES) return json({ fout: 'bestand is groter dan 25 MB' }, 413);
  if (inhoud.byteLength === 0) return json({ fout: 'leeg bestand' }, 400);
  await env.BESTANDEN.put(params.id, inhoud, {
    httpMetadata: { contentType: request.headers.get('Content-Type') || 'application/octet-stream' },
    customMetadata: { naam },
  });
  return json({ ok: true, grootte: inhoud.byteLength });
}

export async function onRequestGet({ params, env, request }) {
  if (!checkAuth(request, env)) return json({ fout: 'niet geautoriseerd' }, 401);
  if (!env.BESTANDEN) return json({ fout: 'R2-bucket (BESTANDEN) is niet gekoppeld aan het project' }, 500);
  if (!geldigId(params.id)) return json({ fout: 'ongeldig id' }, 400);
  const obj = await env.BESTANDEN.get(params.id);
  if (!obj) return json({ fout: 'niet gevonden' }, 404);
  const naam = veiligeNaam(obj.customMetadata && obj.customMetadata.naam);
  return new Response(obj.body, {
    headers: {
      'Content-Type': (obj.httpMetadata && obj.httpMetadata.contentType) || 'application/octet-stream',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(naam)}`,
      'X-Content-Type-Options': 'nosniff',
      'Cache-Control': 'private, no-store',
    },
  });
}

export async function onRequestDelete({ params, env, request }) {
  if (!checkAuth(request, env)) return json({ fout: 'niet geautoriseerd' }, 401);
  if (!env.BESTANDEN) return json({ fout: 'R2-bucket (BESTANDEN) is niet gekoppeld aan het project' }, 500);
  if (!geldigId(params.id)) return json({ fout: 'ongeldig id' }, 400);
  await env.BESTANDEN.delete(params.id);
  return json({ ok: true });
}
