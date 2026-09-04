// functions/api/opslag/[key].js
//
// Cloudflare Pages Function — de enige plek die rechtstreeks met de D1-database praat.
// Verwerkt: GET  /api/opslag/<sleutel>   -> huidige waarde ophalen
//           PUT  /api/opslag/<sleutel>   -> waarde wegschrijven (maakt de rij aan als die nog niet bestaat)
//
// Beveiliging: een simpele gedeelde sleutel (API_SECRET) die zowel hier als in de frontend-
// build moet overeenkomen. Dit is een basale drempel, geen sterke beveiliging — de sleutel
// zit immers ook in de gebouwde frontend-code en is met wat moeite te achterhalen. Voor een
// stevigere afscherming: overweeg later Cloudflare Access (Zero Trust) vóór deze Function te
// zetten, dat is buiten de app-code om te regelen via het Cloudflare-dashboard.

function checkAuth(request, env) {
  if (!env.API_SECRET) return true; // geen sleutel ingesteld -> geen controle (alleen handig tijdens eerste testen, niet aan te raden voor productie)
  const meegestuurd = request.headers.get('X-Api-Key');
  return meegestuurd === env.API_SECRET;
}

export async function onRequestGet(context) {
  const { params, env, request } = context;
  if (!checkAuth(request, env)) {
    return new Response(JSON.stringify({ fout: 'niet geautoriseerd' }), { status: 401, headers: { 'Content-Type': 'application/json' } });
  }
  const sleutel = params.key;
  try {
    const rij = await env.DB.prepare('SELECT waarde FROM opslag WHERE sleutel = ?').bind(sleutel).first();
    if (!rij) {
      return new Response(JSON.stringify({ fout: 'niet gevonden' }), { status: 404, headers: { 'Content-Type': 'application/json' } });
    }
    return new Response(JSON.stringify({ sleutel, waarde: rij.waarde }), { headers: { 'Content-Type': 'application/json' } });
  } catch (e) {
    return new Response(JSON.stringify({ fout: 'databasefout', detail: String(e) }), { status: 500, headers: { 'Content-Type': 'application/json' } });
  }
}

export async function onRequestPut(context) {
  const { params, env, request } = context;
  if (!checkAuth(request, env)) {
    return new Response(JSON.stringify({ fout: 'niet geautoriseerd' }), { status: 401, headers: { 'Content-Type': 'application/json' } });
  }
  const sleutel = params.key;
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ fout: 'ongeldige JSON-inhoud' }), { status: 400, headers: { 'Content-Type': 'application/json' } });
  }
  if (typeof body.waarde !== 'string') {
    return new Response(JSON.stringify({ fout: 'veld "waarde" moet tekst zijn' }), { status: 400, headers: { 'Content-Type': 'application/json' } });
  }
  try {
    await env.DB.prepare(
      `INSERT INTO opslag (sleutel, waarde, bijgewerkt_op) VALUES (?, ?, datetime('now'))
       ON CONFLICT(sleutel) DO UPDATE SET waarde = excluded.waarde, bijgewerkt_op = excluded.bijgewerkt_op`
    ).bind(sleutel, body.waarde).run();
    return new Response(JSON.stringify({ ok: true }), { headers: { 'Content-Type': 'application/json' } });
  } catch (e) {
    return new Response(JSON.stringify({ fout: 'databasefout', detail: String(e) }), { status: 500, headers: { 'Content-Type': 'application/json' } });
  }
}
