// Entrega la última versión publicada con api/publicar:
//   GET /api/datos                 -> { id, total, publicadoEl } (o {} si aún no hay nada publicado)
//   GET /api/datos?id=…&parte=i    -> la parte i (gzip) de esa versión
import { list, get } from '@vercel/blob';

// Credenciales del almacenamiento Blob: token clásico (BLOB_READ_WRITE_TOKEN) o, en almacenes nuevos,
// OIDC + id del almacén (Vercel lo crea como <prefijo>_STORE_ID al conectarlo al proyecto).
const env = Object.entries(process.env);
const TOKEN = process.env.BLOB_READ_WRITE_TOKEN || (env.find(([k]) => k.endsWith('_READ_WRITE_TOKEN')) || [])[1];
const STORE_ID = process.env.BLOB_STORE_ID || (env.find(([k]) => k.endsWith('_STORE_ID')) || [])[1];
const AUTH = TOKEN ? { token: TOKEN } : STORE_ID ? { storeId: STORE_ID } : null;

const PREFIJO = 'compras/';
const CORTO = { 'Cache-Control': 'public, max-age=0, s-maxage=20, stale-while-revalidate=60' };

async function leer(pathname) {
  for (const access of ['private', 'public']) {
    try { const r = await get(pathname, { ...AUTH, access }); if (r && r.statusCode === 200) return r; } catch (e) { /* prueba el otro tipo */ }
  }
  return null;
}

export async function GET(request) {
  try { return await datos(request); }
  catch (e) { return Response.json({ error: 'Almacenamiento Blob: ' + e.message }, { status: 500, headers: { 'Cache-Control': 'no-store' } }); }
}

async function datos(request) {
  if (!AUTH) return Response.json({}, { headers: CORTO });
  const u = new URL(request.url);
  const id = u.searchParams.get('id');
  const parte = u.searchParams.get('parte');

  if (id != null || parte != null) {
    if (!/^\d{13}$/.test(id || '') || !/^\d{1,2}$/.test(parte || '')) return new Response('Parámetros inválidos', { status: 400 });
    const r = await leer(`${PREFIJO}${id}/parte-${parte.padStart(3, '0')}.bin`);
    if (!r) return new Response('No encontrado', { status: 404, headers: { 'Cache-Control': 'no-store' } });
    // cada versión es inmutable (un id nuevo por publicación): se puede guardar en caché
    return new Response(r.stream, { headers: { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'public, max-age=31536000, immutable' } });
  }

  const r = await list({ ...AUTH, prefix: PREFIJO, limit: 1000 });
  const listos = r.blobs.filter(b => b.pathname.endsWith('/listo.json'))
    .sort((a, b) => (a.pathname < b.pathname ? -1 : 1));
  const ultimo = listos[listos.length - 1];
  if (!ultimo) return Response.json({}, { headers: CORTO });
  const id2 = ultimo.pathname.split('/')[1];
  const total = r.blobs.filter(b => b.pathname.startsWith(`${PREFIJO}${id2}/parte-`)).length;
  return Response.json({ id: Number(id2), total, publicadoEl: new Date(ultimo.uploadedAt).toISOString() }, { headers: CORTO });
}
