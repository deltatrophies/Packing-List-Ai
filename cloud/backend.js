import { createClient } from '@supabase/supabase-js';
import * as pdfjs from 'pdfjs-dist';

const url = import.meta.env.VITE_SUPABASE_URL;
const publishableKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY;
export const configured = Boolean(url && publishableKey && !url.includes('YOUR_PROJECT'));
export const supabase = configured ? createClient(url, publishableKey, {
  auth: { autoRefreshToken: true, persistSession: true, detectSessionInUrl: true },
}) : null;
pdfjs.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString();

export async function api(action, payload = {}) {
  if (!supabase) throw new Error('Supabase project is not configured');
  const { data: { session } } = await supabase.auth.getSession();
  if (!session) throw new Error('Sign in first');
  const response = await fetch(`${url}/functions/v1/packing-api`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'apikey': publishableKey, 'Authorization': `Bearer ${session.access_token}` },
    body: JSON.stringify({ action, ...payload }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

async function toJpeg(blob) {
  const bitmap = await createImageBitmap(blob, { imageOrientation: 'from-image' });
  const scale = Math.min(1, 2800 / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const context = canvas.getContext('2d');
  context.fillStyle = 'white'; context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  return new Promise((resolve, reject) => canvas.toBlob(result => result ? resolve(result) : reject(new Error('Could not prepare photo')), 'image/jpeg', .93));
}
export async function preparePages(files) {
  const pages = [];
  for (const file of files) {
    if (file.size > 18 * 1024 * 1024) throw new Error(`${file.name} is larger than 18 MB`);
    if (file.name.toLowerCase().endsWith('.pdf')) {
      const document = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
      if (pages.length + document.numPages > 50) throw new Error('Maximum 50 pages per list');
      for (let i = 1; i <= document.numPages; i++) {
        const page = await document.getPage(i);
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement('canvas'); canvas.width = viewport.width; canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        const blob = await new Promise((resolve, reject) => canvas.toBlob(v => v ? resolve(v) : reject(new Error('PDF rendering failed')), 'image/jpeg', .93));
        pages.push({ blob, name: `${file.name}-page-${i}.jpg` });
      }
      await document.destroy();
    } else if (/\.(jpe?g|png|webp)$/i.test(file.name)) {
      pages.push({ blob: await toJpeg(file), name: file.name });
    } else throw new Error(`${file.name}: unsupported file type`);
  }
  if (!pages.length || pages.length > 50) throw new Error('Choose 1 to 50 pages');
  return pages;
}
export async function uploadPages(pages) {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) throw new Error('Sign in first');
  const batchId = crypto.randomUUID();
  const results = [];
  const uploadedPaths = [];
  try {
    for (const [index, page] of pages.entries()) {
      const path = `${user.id}/${batchId}/page-${index + 1}.jpg`;
      const { error } = await supabase.storage.from('packing-sources').upload(path, page.blob, { contentType: 'image/jpeg', upsert: false });
      if (error) throw error;
      uploadedPaths.push(path);
      const { error: recordError } = await supabase.from('temp_uploads').insert({ path, owner_id: user.id });
      if (recordError) throw recordError;
      const digest = await crypto.subtle.digest('SHA-256', await page.blob.arrayBuffer());
      const hash = [...new Uint8Array(digest)].map(x => x.toString(16).padStart(2, '0')).join('');
      results.push({ path, hash, url: URL.createObjectURL(page.blob) });
    }
    return results;
  } catch (error) {
    if (uploadedPaths.length) {
      await supabase.storage.from('packing-sources').remove(uploadedPaths);
      await supabase.from('temp_uploads').delete().in('path', uploadedPaths);
    }
    throw error;
  }
}

export function combinePages(pages) {
  const data = { customer: '', packing_date: '', private_mark: '', transport: '', boxes: [], warnings: [] };
  const boxes = new Map();
  for (const [index, page] of pages.entries()) {
    if (!data.customer && page.customer?.trim()) data.customer = page.customer.trim();
    if (!data.private_mark && page.private_mark?.trim()) data.private_mark = page.private_mark.trim();
    data.warnings.push(...(page.warnings || []).map(w => `Page ${index + 1}: ${w}`));
    for (const raw of page.boxes || []) {
      if (!Number.isInteger(raw.number) || raw.number < 1) continue;
      let box = boxes.get(raw.number);
      if (!box) { box = { number: raw.number, items: [], needs_review: Boolean(raw.needs_review) }; boxes.set(raw.number, box); }
      else { box.needs_review = true; data.warnings.push(`Box ${raw.number} appears on multiple pages; check merged entries`); }
      for (const item of raw.items || []) box.items.push({
        code: String(item.code || '').trim(), size: String(item.size || '').trim(), quantity: item.quantity ?? null,
        note: String(item.note || '').trim(), source_page: index + 1, source_text: String(item.source_text || ''), needs_review: Boolean(item.needs_review),
      });
    }
  }
  data.boxes = [...boxes.values()].sort((a, b) => a.number - b.number);
  if (data.boxes.length) {
    const missing = [];
    for (let n = data.boxes[0].number; n <= data.boxes.at(-1).number; n++) if (!boxes.has(n)) missing.push(n);
    if (missing.length) data.warnings.push(`Missing box numbers: ${missing.join(', ')}`);
  }
  return data;
}
