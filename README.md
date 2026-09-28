# TUTO Desk

Team to-do app: **Team** tasks are shared, **Personal** tasks are private. Google sign-in (with an invite link),
read-only Google Calendar, push notifications. Works in the browser and as a home-screen app.

| Folder | What |
|---|---|
| `app/` | **The app.** One Cloudflare Worker: `public/` (the pages), `src/worker.js` (API + reminders), `src/push.js` (push notifications), `schema.sql` (database). |
| `mobile/` | Optional Android wrapper (Capacitor). |
| `old-version/` | The earlier Apps Script + Netlify version. Not used any more. |

Deploy: `cd app && npx wrangler deploy` (Claude Code does this for you).

Live app: https://tuto-desk.tuto-desk.workers.dev

## Set up on a new laptop

1. Install [Node.js](https://nodejs.org) (LTS) and [Git](https://git-scm.com).
2. Download the project and install its tools:
   ```
   git clone https://github.com/veronicainproduct/Tuto-Internal-DESK.git
   cd Tuto-Internal-DESK/app
   npm install
   npx wrangler login
   ```
3. To publish changes: `npx wrangler deploy` (from the `app` folder).

Secrets are **not** in this repo; they live on Cloudflare (`GOOGLE_CLIENT_SECRET`, `VAPID_PRIVATE_JWK`, `INVITE_CODE`).
Publishing only needs the login above. To run the app locally (`npm run dev`), create `app/.dev.vars` with those values.
