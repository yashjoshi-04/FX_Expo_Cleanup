# FX Deal ID Cleanup

Upload the exposure workbook, get back the same file with the Sheet3 uploader template filled in.

## Files
- `app.py` – the Streamlit web page
- `cleanup_logic.py` – the cleanup logic (same rules as the original notebook)
- `requirements.txt` – libraries the app needs
- `.gitignore` – stops Excel files from ever being uploaded to GitHub

## Run it on your own machine first (optional)
    pip install -r requirements.txt
    streamlit run app.py

---

## Putting it online: step by step (no command line needed)

### Part A – GitHub (stores the code)
1. Sign in at https://github.com (create a free account if needed).
2. Click the **+** in the top-right → **New repository**.
3. Name it `fx-deal-id-cleanup`.
4. Choose **Private** (recommended – see the note on confidentiality below).
5. Leave "Add a README file" unticked and click **Create repository**.
6. On the empty repo page click **uploading an existing file**.
7. Drag in `app.py`, `cleanup_logic.py`, `requirements.txt`, `.gitignore` and this README.
8. Click **Commit changes**. Done – the code is now on GitHub.

### Part B – Streamlit Community Cloud (runs the app and gives you the link)
1. Go to https://share.streamlit.io and choose **Continue with GitHub**; authorise it
   (when asked, allow access to private repositories).
2. Click **Create app** → **Deploy a public app from GitHub** (or "from existing repo").
3. Repository: `<your-username>/fx-deal-id-cleanup`  |  Branch: `main`  |  Main file path: `app.py`.
4. Optional: pick a custom URL such as `fx-cleanup`.
5. Click **Deploy**. The first build takes a couple of minutes.
6. When it's up, click **Share** (top-right of the app) to see/send the link.

### Part C – Updating the app later
Open the file on GitHub → pencil icon → edit → **Commit changes**. The app redeploys by itself.
To change the default Deal In ID permanently, edit `DEFAULT_DEAL_IN_ID` in `cleanup_logic.py`.

---

## Confidentiality – please read before sharing the link
The app processes uploaded files on Streamlit's cloud servers, not inside your company network.
- A **public** app means anyone on the internet with the link can open it and upload files.
- Deal IDs and quantities are trading data. Check your firm's IT / compliance policy before using an
  external service for them.
- Safer options: keep the repo **private** and the app **private**, then add colleagues by email under
  **Share → Invite viewers**; or ask IT to host the app on an internal server.
- Never upload real workbooks to GitHub (the `.gitignore` helps with this).
