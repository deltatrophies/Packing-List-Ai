# Packing List Studio — Delta Trophies

**Live app:** https://delta-packing-list-ai.onrender.com

**Admin login:** `deltatrophies88@gmail.com` (set a password in Settings while signed in; email links remain available when the mail quota allows)

The sales team can open the live link on any PC. The frontend is hosted on Render. Supabase provides sign-in, private source-photo storage, saved lists, admin-managed verified examples, and an Edge Function that calls Gemini. Gemini keys stay server-side; the browser never receives them.

## Use it

1. Ask the admin to create a sales login. Enter the email and temporary password they give you. Change the password in **Settings** after your first sign-in.
2. Upload JPG, PNG, WebP, or scanned PDF pages in order. Use the arrows to fix page order.
3. Click **Read handwritten sheets**. Compare every box with the source photo. Edit quantities, item codes, sizes, notes, and box numbers where needed. **Preview PDF** is available before saving.
4. Click **Final** after checking. Then download Excel or PDF. The three most recent saved lists appear first; use **Show more** to load five older lists at a time. Reopen a saved list to correct it; click **Final** again after edits.

The print layout gives each box at least five item rows, with blank rows when needed. A box stays on one physical print page in both exports.

## Admin work

- Use **Admin → Create sales login** to make an account. The admin can type and confirm a password (at least six characters, with no required character types) or leave both password fields blank to generate a temporary password. Give it to the sales person privately. This does not send email. The **Sales team** list shows created members and their emails. Passwords cannot be viewed later; **Reset password** can set a chosen password or generate a new one and invalidate the old one.
- Use **Settings** to set your own sign-in password or Gemini key. Every new sales account starts on the company shared key. Saving a personal key selects it for that account and the choice persists across sign-ins; the user can switch back to shared mode. An admin can replace the company shared key, which immediately applies to everyone using shared mode. Personal keys remain private to their owners.
- Only the admin can use **Admin → Verified examples**. Upload matching handwritten photos and the human-verified `.xlsx`, compare every imported row with the photos, then click **Save verified example**. Sales users do not see example uploads or a training checkbox; their finalized lists are never added automatically.
- Verified examples supply company item-code and size spelling hints to future OCR requests. This does **not** train or fine-tune Gemini model weights. Photos attached to admin examples remain available for later review. Ordinary list source photos are deleted when Final succeeds. Abandoned temporary uploads are cleaned up when the app checks status after 48 hours.
- The user should still verify every output against its photos. Handwriting and crossed-out corrections can remain ambiguous.

Supabase's built-in email sender has a very low quota, so email invites and magic links can fail with a rate-limit message. Password logins do not send email. For reliable email invitations later, configure a custom SMTP provider in Supabase Auth; do not raise the built-in limit or depend on it for daily sales logins.

## Initial reference data

The initial admin examples are the Sharda Sports Borsad 5-box workbook with two photos and the two Vinayak workbooks covering boxes 1–20 and 21–34 with their matching photos. Only human-Excel imports count as verified examples; an older AI-reviewed list is excluded from recognition hints.

## Technical setup

- Frontend: `cloud/` (Vite, vanilla JavaScript, ExcelJS, jsPDF). Render Static Site builds with `npm ci && npm run build` and publishes `cloud/dist`.
- Backend: `supabase/functions/packing-api/index.ts`. The function authenticates every request, calls Gemini, and manages key settings and admin approvals.
- Database and storage: `supabase/migrations/202610080001_cloud_app.sql`. Row-level security limits a sales user to their own lists and photos; admins can review examples.
- Repository: https://github.com/deltatrophies/Packing-List-Ai
- Supabase project: `ptrxgxhhdujdxlfucynz` (Mumbai region). Render auto-deploys the GitHub `main` branch.

No API keys or database passwords are committed. Local secret files live in ignored `.local/` paths. The public browser configuration has only the Supabase project URL and publishable key.

## Developer checks

```powershell
cd cloud
npm ci
npm test
npm run build
```

The older Windows-only app remains in `packing_app/`; the live team app is in `cloud/`.
