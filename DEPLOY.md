# Deploying the Beta Calculation Audit as a web app

The web app lives in **`app.py`**. It reuses the calculation logic in
**`weighted_beta_v2.py`** (unchanged behaviour; now importable). Your PM opens a
URL, enters a password, uploads a holdings `.xlsx`, and downloads the generated
`Beta_Calculation_Audit.xlsx`.

> **Why hosting helps:** the *server* fetches from Yahoo Finance (yfinance), not
> your PM's machine. Host it outside mainland China and the Yahoo block on his
> side no longer matters — he only needs to reach the website.

---

## 1. Test it locally first

```bash
pip install -r requirements.txt
mkdir -p .streamlit
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then edit the password
streamlit run app.py
```

A browser tab opens at http://localhost:8501. Upload `Holdings.xlsx`, confirm it
produces the same workbook you get from the command-line script.

## 2. Put the code in a private Git repo

Push these files: `app.py`, `weighted_beta_v2.py`, `requirements.txt`.
**Do not commit** `.streamlit/secrets.toml` (add it to `.gitignore`) — the
password goes in the host's Secrets box instead.

## 3. Deploy — pick one

### Option A — Streamlit Community Cloud (free, easiest)
1. Go to share.streamlit.io, sign in with GitHub, connect the private repo.
2. Set the main file to `app.py`.
3. In **Settings → Secrets**, paste:
   ```
   APP_PASSWORD = "your-long-random-password"
   ```
4. Deploy. Share the URL + password with your PM.
5. **Verify Yahoo access:** run one real holdings file right after deploying —
   if fundamentals/prices come back empty, the runner can't reach Yahoo and you
   should use Option B.

### Option B — Small VPS in Hong Kong / Singapore (~$5/mo, most reliable)
More dependable network path to Yahoo, and the data stays on a box you control.
```bash
ssh into the server
sudo apt update && sudo apt install -y python3-pip
git clone <your-repo> && cd <repo>
pip install -r requirements.txt
mkdir -p .streamlit && nano .streamlit/secrets.toml   # add APP_PASSWORD
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```
Then reach it at `http://<server-ip>:8501`. For a clean URL + HTTPS, put nginx
in front and run Streamlit under `systemd` so it restarts automatically.

## 4. Access control
A single shared `APP_PASSWORD` gates the app (see `app.py` → `check_password`).
That's enough for one PM. To harden: restrict by IP at the host/firewall, or add
per-user logins later.

## Security note
Portfolio holdings are sensitive. A password-gated public URL (Option A) still
routes the uploaded file through a third-party server. If that's a compliance
concern, prefer Option B on a company-controlled server, or just run
`streamlit run app.py` on your own machine and share it over the office network.
