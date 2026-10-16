// Recibe los datos que calculó el navegador de quien actualiza el tablero (en partes de <= 3,5 MB,
// por el límite de 4,5 MB por petición de Vercel) y los guarda en Vercel Blob para que
// api/datos se los entregue a cualquiera que abra el link. Protegido con la clave PUBLICAR_CLAVE.
import { put, list, del } from '@vercel/blob';
import { timingSafeEqual } from 'node:crypto';

// Credenciales del almacenamiento Blob: token clásico (BLOB_READ_WRITE_TOKEN) o, en almacenes nuevos,
// OIDC + id del almacén (Vercel lo crea como <prefijo>_STORE_ID al conectarlo al proyecto).
const env = Object.entries(process.env);
const TOKEN = process.env.BLOB_READ_WRITE_TOKEN || (env.find(([k]) => k.endsWith('_READ_WRITE_TOKEN')) || [])[1];
const STORE_ID = process.env.BLOB_STORE_ID || (env.find(([k]) => k.endsWith('_STORE_ID')) || [])[1];
const AUTH = TOKEN ? { token: TOKEN } : STORE_ID ? { storeId: STORE_ID } : null;

const PREFIJO = 'compras/';
let acceso = null; // 'private' o 'public', según cómo se haya creado el almacenamiento Blob

async function subir(pathname, body, contentType) {
  const opts = access => ({ ...AUTH, access, addRandomSuffix: false, allowOverwrite: true, contentType });
  if (acceso) return put(pathname, body, opts(acceso));
  try { const r = await put(pathname, body, opts('private')); acceso = 'private'; return r; }
  catch (e) {
    try { const r = await put(pathname, body, opts('public')); acceso = 'public'; return r; }
    catch (e2) { throw e; } // el error del intento privado es el que explica el problema
  }
}

async function listarTodo() {
  const out = [];
  let cursor;
  do {
    const r = await list({ ...AUTH, prefix: PREFIJO, cursor, limit: 1000 });
    out.push(...r.blobs);
    cursor = r.hasMore ? r.cursor : undefined;
  } while (cursor);
  return out;
}

function claveOk(c) {
  const k = Buffer.from(process.env.PUBLICAR_CLAVE || '');
  const a = Buffer.from(String(c || ''));
  return k.length > 0 && a.length === k.length && timingSafeEqual(a, k);
}

const json = (o, status = 200) => Response.json(o, { status, headers: { 'Cache-Control': 'no-store' } });

export async function POST(request) {
  try { return await publicar(request); }
  catch (e) { return json({ error: 'Almacenamiento Blob: ' + e.message }, 500); }
}

async function publicar(request) {
  if (!process.env.PUBLICAR_CLAVE) return json({ error: 'Falta configurar PUBLICAR_CLAVE en Vercel' }, 500);
  if (!claveOk(request.headers.get('x-clave'))) return json({ error: 'Clave incorrecta' }, 401);
  if (!AUTH) return json({ error: 'Falta conectar el almacenamiento Blob en Vercel' }, 500);

  const u = new URL(request.url);
  const id = u.searchParams.get('id') || '';
  const parte = Number(u.searchParams.get('parte'));
  const total = Number(u.searchParams.get('total'));
  if (!/^\d{13}$/.test(id) || !Number.isInteger(parte) || !Number.isInteger(total) || parte < 0 || parte >= total || total > 50) {
    return json({ error: 'Parámetros inválidos' }, 400);
  }

  const body = Buffer.from(await request.arrayBuffer());
  await subir(`${PREFIJO}${id}/parte-${String(parte).padStart(3, '0')}.bin`, body, 'application/octet-stream');
  if (parte < total - 1) return json({ ok: true });

  // Última parte: confirma que llegaron todas, marca la versión como lista y borra las anteriores.
  const todos = await listarTodo();
  const partes = todos.filter(b => b.pathname.startsWith(`${PREFIJO}${id}/parte-`));
  if (partes.length !== total) return json({ error: `Llegaron ${partes.length} de ${total} partes` }, 500);
  await subir(`${PREFIJO}${id}/listo.json`, JSON.stringify({ id: Number(id), total, publicadoEl: new Date().toISOString() }), 'application/json');
  const viejos = todos.filter(b => !b.pathname.startsWith(`${PREFIJO}${id}/`)).map(b => b.url);
  if (viejos.length) await del(viejos, AUTH);
  return json({ ok: true, id: Number(id) });
}
