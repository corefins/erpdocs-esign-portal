# ERPDocs Client E-Sign Portal

Streamlit app that sends a client an intake form and an ERPDocs signing session
behind a single link. Client records come from a Google Sheet; signing runs
through the ERPDocs REST API using embedded signing, so the signer never leaves
this app.

Split out of the Tally automation project for hosted testing. The Tally
integration is deliberately **not** included — it needs a private-LAN
connection to Tally Prime that a hosted container cannot make.

---

## Repository policy

**No credential ever enters this repository.** `.gitignore` denies
credential-shaped filenames by default and re-allows only the two templates.
Before your first push, confirm:

```bash
git status --short          # no secrets.json / config.json / *.toml
git ls-files | grep -Ei 'secret|credential|service.?account|\.env'
```

The second command should return only `secrets.example.json` and
`.streamlit/secrets.toml.example`.

---

## Running locally

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

cp secrets.example.json secrets.json     # fill in real values
# create config.json with google_service_account_path pointing at your key file

streamlit run streamlit_app.py
```

Credentials resolve local-file-first, so local development is unchanged.

---

## Deploying to Streamlit Community Cloud

### 1. Push to a private repository

Community Cloud deploys from GitHub. Use a **private** repo — this app reads a
client list, sends mail as you, and holds an ERPDocs API key.

```bash
git init && git add . && git commit -m "Initial commit"
git remote add origin git@github.com:<you>/erpdocs-esign-portal.git
git push -u origin main
```

### 2. Create the app

Point it at this repo, branch `main`, main file `streamlit_app.py`.

### 3. Add secrets

Copy `.streamlit/secrets.toml.example`, fill it in, and paste the result into
**Settings -> Secrets**. There is no secrets file in the repo, so the app cannot
reach Google Sheets, SMTP or ERPDocs until this is done.

Set `form_base_url` to the app's own `https://` URL. Client links are built from
it, so a leftover `localhost` value produces dead links.

### 4. Restrict who can open it

**Settings -> Sharing** — limit viewers to your own accounts. A public URL here
would expose client records and let anyone send mail through your SMTP
credentials.

### 5. Register the app as an ERPDocs embed origin

Embedded signing only renders inside an origin registered on the API key.
In ERPDocs: **Admin -> API & Webhooks**, open the key, add under **Embed
Origins**:

```
https://YOUR-APP.streamlit.app
```

Exact origin — no trailing slash, no path, no wildcard. Without this, minting a
token returns `403 origin_not_allowed`. In the app's Settings -> Send & Embed,
set **Embed Origin** to the same value.

See the [Embedded Signing guide](https://www.erpdocs.com/kb/developers-and-api/embedded-signing).

---

## Known constraints when hosted

**Settings changes do not survive a restart.** The Settings page writes
`signature_settings.json` to local disk. Community Cloud containers have an
ephemeral filesystem and are recycled on redeploy, on config change, and after
inactivity — so anything set through that page reverts to defaults. Treat the
Settings page as a scratchpad and keep the real values in
**Settings -> Secrets**, which does persist.

Making settings durable means moving them out of a file (into `st.secrets` for
static values plus `st.session_state` for per-session ones). Not done here.

**Cold starts.** Idle apps sleep and take a few seconds to wake. An embed token
has a 30–900s TTL and is single-use, so a token minted just before a sleep may
expire; the app mints a fresh one on the next page load.

**No Tally.** Excluded by design, as above.

---

## Environment check

| Setting | Value |
|---|---|
| ERPDocs base URL | `https://www.erpdocs.com` (production) |
| Envelopes | Real. Emails reach real signers. |
| Required API scopes | `template.read`, `signpack.read`, `signpack.write`, `signpack.embed`, `brand.read`, `brand.write` |

This is a production key, so anything the test app sends is a live envelope
against live data. Use the app's **Test Connection** button to confirm the key,
org and granted scopes before sending anything. To test without touching real
data, switch `erpdocs_base_url` to `https://staging.erpdocs.com` and use a
staging key.
