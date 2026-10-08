// One-time import of the locally verified reference cases.
// Run with SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY and ADMIN_EMAIL in the environment.
import { readFile, readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createClient } from '@supabase/supabase-js';

const url = process.env.SUPABASE_URL;
const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
const adminEmail = process.env.ADMIN_EMAIL?.toLowerCase();
if (!url || !serviceKey || !adminEmail) throw new Error('SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY and ADMIN_EMAIL are required');
const supabase = createClient(url, serviceKey, { auth: { persistSession: false } });
const { data: users, error: usersError } = await supabase.auth.admin.listUsers();
if (usersError) throw usersError;
const owner = users.users.find(user => user.email?.toLowerCase() === adminEmail);
if (!owner) throw new Error('Admin user does not exist in Supabase Auth');
const root = fileURLToPath(new URL('../../.local/training/', import.meta.url));
const folders = await readdir(root, { withFileTypes: true });
for (const entry of folders.filter(x => x.isDirectory())) {
  const source = JSON.parse(await readFile(join(root, entry.name, 'case.json'), 'utf8'));
  const hashes = source.photo_hashes || [];
  const { data: existing } = await supabase.from('packing_lists').select('id,source_hashes').eq('status', 'training_approved');
  if (existing?.some(row => JSON.stringify(row.source_hashes) === JSON.stringify(hashes))) { process.stdout.write(`Already imported ${entry.name}\n`); continue; }
  const paths = [];
  for (const [index, filename] of source.photos.entries()) {
    const path = `${owner.id}/reference-${entry.name}/page-${index + 1}.jpg`;
    const bytes = await readFile(join(root, entry.name, filename));
    const { error } = await supabase.storage.from('packing-sources').upload(path, bytes, { contentType: 'image/jpeg', upsert: true });
    if (error) throw error;
    paths.push(path);
  }
  const { error } = await supabase.from('packing_lists').insert({
    owner_id: owner.id, data: source.list, source_paths: paths, source_hashes: hashes,
    origin: source.origin === 'manual_import' ? 'manual_import' : 'recognition',
    status: 'training_approved', approved_by: owner.id, approved_at: new Date().toISOString(),
  });
  if (error) throw error;
  process.stdout.write(`Imported ${entry.name}: ${source.list.boxes.length} boxes, ${paths.length} photos\n`);
}
