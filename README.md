# Field Work Report System (Python / Flask) – deploy on Vercel

No login, as requested. Data is kept in a PostgreSQL database (Vercel has no permanent disk).

## 1. Deploy
1. Put this folder in a GitHub repository (or use the Vercel CLI: `vercel` then `vercel --prod`).
2. On https://vercel.com/new import the repository. Vercel detects Flask automatically – no build settings needed.
3. In your Vercel project open **Storage → Create / Connect Database → Neon (Postgres)** and connect it to the project
   (this adds `DATABASE_URL` / `POSTGRES_URL` automatically).
4. In **Settings → Environment Variables** add `SECRET_KEY` = any long random text (keeps the selected date/session safe).
5. **Redeploy** (new environment variables only apply to new deployments).
6. Open the site. The tables are created automatically on the first visit.

## 2. Run on your own computer (optional)
```
pip install -r requirements.txt
python app.py            # http://127.0.0.1:5000  (uses a local SQLite file local_dev.db)
```
To test against your Neon database instead, run `vercel env pull .env.local` and set `DATABASE_URL` in your shell.

## 3. Notes
* Excel files are generated on the fly: single downloads, search downloads and the combined daily file
  (one sheet for tax payments + one for invitation letters, rebuilt from the saved records each time).
* TIN / Ghana Card number is unique within each report. Everything is saved in CAPITALS.
* Because there is no login, anyone who has the web address can view and change the data.
  Keep the address private and consider turning on Vercel's **Deployment Protection** (Project → Settings)
  so only authorised people can open the site.
* Data from the old PHP/MySQL version is not moved automatically.
