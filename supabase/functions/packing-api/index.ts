import { createClient } from 'npm:@supabase/supabase-js@2';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Content-Type': 'application/json',
};
const BUCKET = 'packing-sources';
const MODELS = ['gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-3.7-flash', 'gemini-3.6-flash'];
const PROMPT = `Read this photographed handwritten packing-list page carefully. Return every box and every ACTIVE item as structured data.
Rules:
1. A heading such as Box-13 starts a new box. Keep the exact box number. Never invent a missing box.
2. Ignore lines fully crossed out. If corrected, use final uncrossed text and flag uncertainty.
3. Carry product code forward on a row with only a size. PC-404 M 3 then S 1 means PC-404 M x3 and PC-404 S x1.
4. Multiple sizes with a shared quantity mean one item per size with that quantity.
5. Carefully distinguish PC, FC, PF, AA, W and similar digits. Known vocabulary is only a spelling hint.
6. Preserve ONLY BASE and ONLY CUP as notes, not items or quantities.
7. Never guess unreadable code or quantity; use empty/null and needs_review=true. Describe uncertainty in warnings.
8. Active writing in the PACKED BY footer can belong to the final box.
9. Return boxes in image order. Ignore blank printed rows.
10. Do not add rows just to reach five. Export adds blank rows.
11. A named product such as OLYMPIC MEDAL 2.5\" (GOLDEN) goes entirely in code; size stays empty.
12. Source text is a short visible transcription, useful for human review.
Return JSON matching the schema exactly.`;

const ITEM_SCHEMA = {
  type: 'OBJECT', properties: {
    code: { type: 'STRING' }, size: { type: 'STRING' }, quantity: { type: 'INTEGER', nullable: true },
    note: { type: 'STRING' }, source_text: { type: 'STRING' }, needs_review: { type: 'BOOLEAN' },
  }, required: ['code', 'size', 'quantity', 'note', 'source_text', 'needs_review'],
};
const RESPONSE_SCHEMA = {
  type: 'OBJECT', properties: {
    customer: { type: 'STRING' }, private_mark: { type: 'STRING' },
    boxes: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
      number: { type: 'INTEGER' }, items: { type: 'ARRAY', items: ITEM_SCHEMA }, needs_review: { type: 'BOOLEAN' },
    }, required: ['number', 'items', 'needs_review'] } },
    warnings: { type: 'ARRAY', items: { type: 'STRING' } },
  }, required: ['customer', 'private_mark', 'boxes', 'warnings'],
};

type Actor = { id: string; email: string; role: 'admin' | 'sales'; key_mode: 'shared' | 'personal' };
type Item = { code: string; size: string; quantity: number | null; note: string; source_page?: number; source_text?: string; needs_review?: boolean };
type Box = { number: number; items: Item[]; needs_review?: boolean };
type PackingList = { customer: string; packing_date: string; private_mark: string; transport: string; boxes: Box[]; warnings?: string[] };

function reply(data: unknown, status = 200) { return new Response(JSON.stringify(data), { status, headers: CORS }); }
function fail(message: string, status = 400) { return reply({ error: message }, status); }
function bytesToB64(bytes: Uint8Array) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}
function b64ToBytes(value: string) { return Uint8Array.from(atob(value), c => c.charCodeAt(0)); }
function encryptionKey() {
  const encoded = Deno.env.get('CREDENTIAL_ENCRYPTION_KEY');
  if (!encoded) throw new Error('Credential encryption secret is not configured');
  const raw = b64ToBytes(encoded);
  if (raw.length !== 32) throw new Error('Credential encryption secret must be 32 bytes');
  return crypto.subtle.importKey('raw', raw, 'AES-GCM', false, ['encrypt', 'decrypt']);
}
async function encryptKey(value: string) {
  const nonce = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce }, await encryptionKey(), new TextEncoder().encode(value)));
  return `${bytesToB64(nonce)}.${bytesToB64(encrypted)}`;
}
async function decryptKey(value: string) {
  const [nonce, encrypted] = value.split('.');
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64ToBytes(nonce) }, await encryptionKey(), b64ToBytes(encrypted));
  return new TextDecoder().decode(plain);
}
function secretKey() {
  const current = Deno.env.get('SUPABASE_SECRET_KEYS');
  if (current) return JSON.parse(current).default as string;
  return Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') || '';
}
const admin = createClient(Deno.env.get('SUPABASE_URL')!, secretKey(), { auth: { persistSession: false } });

async function actorFor(request: Request): Promise<Actor> {
  const token = request.headers.get('authorization')?.replace(/^Bearer\s+/i, '');
  if (!token) throw new Error('Sign in first');
  const { data: auth, error: authError } = await admin.auth.getUser(token);
  if (authError || !auth.user) throw new Error('Session expired; sign in again');
  const { data: profile, error } = await admin.from('profiles').select('id,email,role,key_mode').eq('id', auth.user.id).single();
  if (error || !profile) throw new Error('Account profile is not ready');
  return profile as Actor;
}
async function credential(actor: Actor) {
  const ownerKey = actor.key_mode === 'personal' ? actor.id : 'shared';
  const { data } = await admin.from('gemini_credentials').select('encrypted_key').eq('owner_key', ownerKey).maybeSingle();
  if (data?.encrypted_key) return decryptKey(data.encrypted_key);
  if (ownerKey === 'shared') return Deno.env.get('GEMINI_API_KEY') || '';
  return '';
}
async function hasCredential(ownerKey: string) {
  const { data } = await admin.from('gemini_credentials').select('owner_key').eq('owner_key', ownerKey).maybeSingle();
  return Boolean(data) || (ownerKey === 'shared' && Boolean(Deno.env.get('GEMINI_API_KEY')));
}
async function purgeExpiredUploads() {
  const cutoff = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const { data } = await admin.from('temp_uploads').select('path').lt('created_at', cutoff).limit(50);
  if (!data?.length) return;
  const paths = data.map(row => row.path);
  const { error } = await admin.storage.from(BUCKET).remove(paths);
  if (!error) await admin.from('temp_uploads').delete().in('path', paths);
}
async function guidance() {
  const { data } = await admin.from('packing_lists').select('data').eq('status', 'training_approved').order('approved_at', { ascending: false }).limit(200);
  const codes = new Map<string, number>();
  const labels = new Map<string, number>();
  for (const record of data || []) for (const box of (record.data?.boxes || [])) for (const item of (box.items || [])) {
    const code = String(item.code || '').trim().toUpperCase();
    const label = [code, String(item.size || '').trim().toUpperCase()].filter(Boolean).join(' ');
    if (code) codes.set(code, (codes.get(code) || 0) + 1);
    if (label) labels.set(label, (labels.get(label) || 0) + 1);
  }
  const rank = (values: Map<string, number>, count: number) => [...values].sort((a, b) => b[1] - a[1]).slice(0, count).map(v => v[0]).join(', ');
  return `\nVerified company examples are spelling hints only. Current image is authoritative.\nKnown codes: ${rank(codes, 85).slice(0, 1300)}\nLabels: ${rank(labels, 35).slice(0, 1200)}\n`;
}
async function recognize(actor: Actor, body: Record<string, unknown>) {
  const path = String(body.path || '');
  if (!path.startsWith(`${actor.id}/`) || path.includes('..')) return fail('Invalid source path', 403);
  const { data: upload } = await admin.from('temp_uploads').select('path').eq('path', path).eq('owner_id', actor.id).maybeSingle();
  if (!upload) return fail('Source photo was not registered', 404);
  const key = await credential(actor);
  if (!key) return fail(actor.key_mode === 'personal' ? 'Add your personal Gemini key in Settings' : 'Shared Gemini key is not configured', 409);
  const { data: file, error: fileError } = await admin.storage.from(BUCKET).download(path);
  if (fileError || !file) return fail('Could not read source photo', 404);
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.length > 6 * 1024 * 1024) return fail('Prepared photo exceeds 6 MB', 413);
  const prompt = PROMPT + await guidance() + (body.continuation ? `\nThis page may continue Box ${Number(body.continuation)} above the first new heading.` : '');
  const payload = {
    contents: [{ role: 'user', parts: [{ inline_data: { mime_type: 'image/jpeg', data: bytesToB64(bytes) } }, { text: prompt }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: RESPONSE_SCHEMA },
  };
  const errors: string[] = [];
  for (const model of MODELS) {
    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'POST', headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key }, body: JSON.stringify(payload),
        signal: AbortSignal.timeout(110000),
      });
      const result = await response.json();
      if (!response.ok) {
        const message = String(result?.error?.message || `HTTP ${response.status}`).slice(0, 250);
        errors.push(`${model}: ${message}`);
        if (![404, 429, 500, 502, 503, 504].includes(response.status)) break;
        continue;
      }
      const text = result?.candidates?.[0]?.content?.parts?.find((part: { text?: string }) => part.text)?.text;
      if (!text) throw new Error('Gemini returned no structured text');
      const page = JSON.parse(text);
      if (!Array.isArray(page.boxes)) throw new Error('Gemini returned invalid box data');
      if (model !== MODELS[0]) page.warnings = [...(page.warnings || []), `Backup model ${model} was used`];
      return reply({ page, model });
    } catch (error) { errors.push(`${model}: ${String(error).slice(0, 200)}`); }
  }
  return fail(`Recognition failed: ${errors.join(' | ')}`, 502);
}
function validateList(data: PackingList) {
  if (!data || !Array.isArray(data.boxes) || !data.boxes.length) throw new Error('Add at least one box');
  const numbers = new Set<number>();
  for (const box of data.boxes) {
    if (!Number.isInteger(box.number) || box.number < 1 || numbers.has(box.number)) throw new Error('Box numbers must be unique positive numbers');
    numbers.add(box.number);
    if (!Array.isArray(box.items)) throw new Error(`Box ${box.number} has invalid items`);
    for (const [index, item] of box.items.entries()) {
      if (!String(item.code || '').trim()) throw new Error(`Box ${box.number}, row ${index + 1}: item code is empty`);
      if (!Number.isInteger(item.quantity) || (item.quantity as number) < 1) throw new Error(`Box ${box.number}, row ${index + 1}: quantity must be at least 1`);
    }
  }
}
async function finalize(actor: Actor, body: Record<string, unknown>) {
  const data = body.data as PackingList;
  validateList(data);
  const paths = Array.isArray(body.paths) ? body.paths.map(String) : [];
  if (paths.length > 50 || paths.some(path => !path.startsWith(`${actor.id}/`) || path.includes('..'))) return fail('Invalid source photos', 403);
  const existingId = body.id ? String(body.id) : null;
  let existing: { id: string; source_paths: string[]; status: string } | null = null;
  if (existingId) {
    const found = await admin.from('packing_lists').select('id,source_paths,status').eq('id', existingId).eq('owner_id', actor.id).maybeSingle();
    existing = found.data;
    if (!existing) return fail('Final list not found', 404);
  } else if (paths.length === 0) return fail('Source photos are required', 400);
  if (paths.length) {
    const { data: uploads, error } = await admin.from('temp_uploads').select('path').eq('owner_id', actor.id).in('path', paths);
    if (error || uploads?.length !== paths.length) return fail('One or more source photos are unavailable', 400);
  }
  const keep = (Boolean(body.keep_for_training) && paths.length > 0) || (existing?.source_paths?.length || 0) > 0;
  const storedPaths = keep ? (paths.length ? paths : existing?.source_paths || []) : [];
  const record = {
    owner_id: actor.id, data, source_paths: storedPaths,
    source_hashes: Array.isArray(body.hashes) ? body.hashes.map(String).slice(0, 50) : [],
    origin: body.origin === 'manual_import' ? 'manual_import' : 'recognition',
    status: keep ? 'training_candidate' : 'final', updated_at: new Date().toISOString(),
  };
  const query = existing ? admin.from('packing_lists').update(record).eq('id', existing.id).select('id').single()
                         : admin.from('packing_lists').insert(record).select('id').single();
  const { data: saved, error } = await query;
  if (error || !saved) return fail(error?.message || 'Could not save final list', 500);
  if (!keep && paths.length) {
    const { error: removeError } = await admin.storage.from(BUCKET).remove(paths);
    if (!removeError) await admin.from('temp_uploads').delete().in('path', paths).eq('owner_id', actor.id);
  } else if (paths.length) await admin.from('temp_uploads').delete().in('path', paths).eq('owner_id', actor.id);
  return reply({ id: saved.id, training_status: record.status });
}

Deno.serve(async request => {
  if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });
  if (request.method !== 'POST') return fail('POST required', 405);
  try {
    const actor = await actorFor(request);
    const body = await request.json() as Record<string, unknown>;
    const action = String(body.action || '');
    if (action === 'status') {
      await purgeExpiredUploads();
      const { count } = await admin.from('packing_lists').select('id', { count: 'exact', head: true }).eq('status', 'training_approved');
      return reply({
        role: actor.role, key_mode: actor.key_mode, approved_examples: count || 0,
        has_shared_key: await hasCredential('shared'), has_personal_key: await hasCredential(actor.id),
      });
    }
    if (action === 'set-key') {
      const scope = String(body.scope || 'personal');
      if (scope === 'shared' && actor.role !== 'admin') return fail('Only an admin may change the shared key', 403);
      if (!['shared', 'personal'].includes(scope)) return fail('Invalid key scope');
      const key = String(body.key || '').trim();
      if (key.length < 20 || key.length > 300) return fail('Enter a valid Gemini API key');
      const ownerKey = scope === 'shared' ? 'shared' : actor.id;
      const { error } = await admin.from('gemini_credentials').upsert({ owner_key: ownerKey, encrypted_key: await encryptKey(key), updated_by: actor.id, updated_at: new Date().toISOString() });
      if (error) return fail('Could not save key', 500);
      return reply({ saved: true, scope });
    }
    if (action === 'set-key-mode') {
      if (!['shared', 'personal'].includes(String(body.mode))) return fail('Invalid key mode');
      const { error } = await admin.from('profiles').update({ key_mode: body.mode }).eq('id', actor.id);
      if (error) return fail('Could not change key mode', 500);
      return reply({ key_mode: body.mode });
    }
    if (action === 'recognize') return recognize(actor, body);
    if (action === 'finalize') return finalize(actor, body);
    if (action === 'training-queue') {
      if (actor.role !== 'admin') return fail('Admin access required', 403);
      const { data, error } = await admin.from('packing_lists').select('id,owner_id,data,source_paths,origin,created_at').eq('status', 'training_candidate').order('created_at', { ascending: false }).limit(100);
      if (error) return fail(error.message, 500);
      return reply({ candidates: data });
    }
    if (action === 'review-training') {
      if (actor.role !== 'admin') return fail('Admin access required', 403);
      const id = String(body.id || '');
      const approve = Boolean(body.approve);
      const { data: example } = await admin.from('packing_lists').select('id,source_paths').eq('id', id).eq('status', 'training_candidate').maybeSingle();
      if (!example) return fail('Training candidate not found', 404);
      const { error } = await admin.from('packing_lists').update({ status: approve ? 'training_approved' : 'training_rejected', approved_by: actor.id, approved_at: new Date().toISOString() }).eq('id', id);
      if (error) return fail(error.message, 500);
      if (!approve && example.source_paths.length) {
        const { error: removeError } = await admin.storage.from(BUCKET).remove(example.source_paths);
        if (removeError) return fail('Could not delete source photos; review was saved but cleanup is needed', 500);
      }
      if (!approve) await admin.from('packing_lists').update({ source_paths: [] }).eq('id', id);
      return reply({ reviewed: true, approved: approve });
    }
    if (action === 'invite') {
      if (actor.role !== 'admin') return fail('Admin access required', 403);
      const email = String(body.email || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return fail('Enter a valid email');
      const { error } = await admin.auth.admin.inviteUserByEmail(email);
      if (error) return fail(error.message, 400);
      return reply({ invited: true });
    }
    if (action === 'list-sales-members') {
      if (actor.role !== 'admin') return fail('Admin access required', 403);
      const { data, error } = await admin.from('profiles').select('id,email,created_at').eq('role', 'sales').order('created_at', { ascending: false }).limit(500);
      if (error) return fail('Could not load sales members', 500);
      return reply({ members: data || [] });
    }
    if (action === 'create-sales-login' || action === 'reset-sales-password') {
      if (actor.role !== 'admin') return fail('Admin access required', 403);
      const email = String(body.email || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return fail('Enter a valid sales email');
      const random = new Uint8Array(24);
      crypto.getRandomValues(random);
      const password = btoa(String.fromCharCode(...random)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
      if (action === 'create-sales-login') {
        const { data: created, error } = await admin.auth.admin.createUser({ email, password, email_confirm: true });
        if (error || !created.user) return fail(error?.message || 'Could not create sales login', 400);
      } else {
        const { data: profile, error: profileError } = await admin.from('profiles').select('id,role').eq('email', email).maybeSingle();
        if (profileError || !profile) return fail('Sales account not found', 404);
        if (profile.role !== 'sales') return fail('Only sales passwords can be reset here', 403);
        const { error } = await admin.auth.admin.updateUserById(profile.id, { password });
        if (error) return fail(error.message, 400);
      }
      return reply({ email, password });
    }
    return fail('Unknown action', 404);
  } catch (error) { return fail(String(error instanceof Error ? error.message : error), 400); }
});
