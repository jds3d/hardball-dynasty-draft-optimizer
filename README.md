# Hardball Dynasty Amateur Draft Tool

Sync your draft Excel file with [WhatIfSports Hardball Dynasty](https://www.whatifsports.com/hbd/Pages/GM/AmateurDraftPlayerPool.aspx): pull prospect data from the site into Excel, then apply your preferred draft order back to the "Rank Players" list on the site.

## Setup

1. **Python 3.10+** and pip.

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. **Chrome** (for Selenium). The script uses `webdriver-manager` to fetch the matching ChromeDriver.

4. Put **`Season x amateur draft-template.xlsx`** in the project root. The script uses it by default.

5. **Auto-login (optional):** To skip logging in each time, add your HBD credentials. Either:
   - Create **`credentials.env`** in the project root with:
     ```
     USERNAME=your_email@example.com
     PASSWORD=your_password
     ```
     (Use `credentials.env.example` as a template.)  
   - Or set env vars **`HBD_USERNAME`** and **`HBD_PASSWORD`**.  
   The file is in `.gitignore`; never commit it.

## Excel file format

Your workbook should have two sheets:

- **Hitters** – header row 6 with columns including `Rnk`, `Player`, `Pos`, `B`, `T`, `Age`, and any rating columns you use.
- **Pitchers** – header row 5 with the same kind of columns.

The tool reads the **draft order** by **interleaving** Hitters and Pitchers and **ranking by score**: Hitters use the **total** column, Pitchers use **Overall Projection**. The combined list is sorted by that score (highest first) and applied to the site. So your formulas in the template drive the order.

## Commands

### 1. Populate Excel from the web (fetch)

Fetches the draft pool from the Amateur Draft Player Pool page, reads the **season number** from the site (e.g. from “Strawberry-Gooden (30) - Scottsdale”), and saves the populated file to **`./outputs/Season N amateur draft YYYY-MM-DD_HH-MM-SS.xlsx`** (e.g. `outputs/Season 30 amateur draft 2025-03-13_16-30-45.xlsx`) so each run gets a unique file and won’t conflict with an open file. Your template is not modified. You must be logged in to WhatIfSports when the browser opens (or log in when prompted).

From the project folder (template in root):

```bash
python main.py fetch
```

Or pass a template path: `python main.py "path/to/template.xlsx" fetch`

Options:

- `--top 500` – number of prospects to load (default 500).
- `--output-dir PATH` – folder for the saved file (default: `outputs`).
- **Chrome**: The script opens a **separate** Chrome window (no profile), so you can leave your normal Chrome open. Log in to HBD in that window when it opens. To reuse a saved login without logging in each time, use a dedicated automation profile: create a second Chrome profile (e.g. “HBD”), log in to HBD there once, then run with `--chrome-profile "%LOCALAPPDATA%\Google\Chrome\User Data\Profile 2"` (close that profile’s Chrome before running).
- `--headless` – run Chrome without a window (not recommended for first run, since you can’t log in interactively).

### 2. Apply Excel order to the web (Rank Players)

Reads the draft order from your Excel (Hitters and Pitchers interleaved, ranked by score), opens the draft pool page, clicks **Rank Players**, reorders the list in the popup to match that order, then clicks **Save**.

Use the populated file in `outputs/` (after you’ve edited/sorted it), or the template:

```bash
python main.py apply-order
python main.py "outputs/Season 30 amateur draft.xlsx" apply-order
```

Use `--chrome-profile` and `--headless` as above if you want.

## Workflow

1. **Fetch** – Run `fetch` once (with browser and login) to fill your Excel with the current draft pool.
2. In Excel, sort or edit the **Hitters** and **Pitchers** sheets so the rows are in the order you want to draft (e.g. by your own ratings or formulas).
3. **Apply** – Run `apply-order` so the site’s "Rank Players" list matches that order.

## Does this work when I run it?

**Logic:** Yes. Fetch reads the draft pool and season from the site, writes data into your template layout, and saves to `outputs/Season N amateur draft.xlsx`. Apply-order reads Hitters + Pitchers from that file (or the template), ranks everyone by score (total / Overall Projection), and reorders the Rank Players list on the site to match.

**Before you run:** Be logged in to WhatIfSports (run without `--headless` once and log in when Chrome opens). Run from the project folder so the default template path is correct.

**If something fails:** The site’s HTML may differ from what the script expects. If you get “Could not find draft prospects table” or “Could not find Rank Players button”, open `web_draft.py` and adjust the selectors (e.g. table id, button text) to match the live page. Template validation runs first; if your workbook is missing the Hitters/Pitchers sheets or the `total` / `Overall Projection` columns, you’ll get a clear error.

## Notes

- **Login**: The Amateur Draft page requires a logged-in session. Run without `--headless` at least once and log in when the browser opens, or use `--chrome-profile` with a profile that’s already logged in (Chrome must be closed).
- **Selectors**: If the site’s HTML changes, table or popup selectors in `web_draft.py` may need to be updated (e.g. `table#dgPlayers`, or the Rank Players / Save button XPaths).
- **Order**: Only the **order** of player names is applied to the Rank Players list; the tool does not create or delete prospects on the site.
