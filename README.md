# Hardball Dynasty Draft Optimizer

Automate your [WhatIfSports Hardball Dynasty](https://www.whatifsports.com/hbd/) amateur draft workflow: scrape the full prospect pool into a scored Excel workbook, then push your preferred draft order back to the site's "Rank Players" list with one command.

## Features

- Scrapes hitting, fielding, pitching, and background-info views into a single Excel file.
- Generates a **Master List** that ranks every prospect by an adjusted score factoring in your template's projection formula, scouting-budget trust, and signability risk.
- Applies that order to the site's Rank Players popup via Selenium (instant JavaScript reorder).
- All penalty weights and formula constants are configurable in `credentials.env`.

## Setup

1. **Python 3.10+** and pip.

2. Create and activate a virtual environment, then install dependencies:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate        # Windows
   # source .venv/bin/activate   # macOS / Linux
   pip install -r requirements.txt
   ```

3. **Google Chrome** must be installed. The script uses `webdriver-manager` to fetch a matching ChromeDriver automatically.

4. Place your **Excel template** (e.g. `Season x amateur draft-template.xlsx`) in the project root. The template must contain:
   - A **Hitters** sheet (header row 6) with columns: `Rnk`, `Player`, `Pos`, `B`, `T`, `Age`, plus rating columns and an **Overall Projection** formula in column A.
   - A **Pitchers** sheet (header row 5) with the same structure.

5. **Credentials and config:** Copy `credentials.env.example` to `credentials.env` and fill in your values. This file is in `.gitignore` — never commit it.

   ```bash
   cp credentials.env.example credentials.env
   ```

## Commands

### `fetch` — Scrape the draft pool into Excel

Pulls all prospect data from the Amateur Draft Player Pool page across four views (Hitting, Fielding/General, Pitching, Background Info), merges them, and writes the result to a timestamped file in `outputs/`.

```bash
python main.py fetch
python main.py "path/to/template.xlsx" fetch
```

| Option | Default | Description |
|--------|---------|-------------|
| `--top N` | 500 | Number of prospects to load |
| `--output-dir PATH` | `outputs` | Folder for the saved file |
| `--headless` | off | Run Chrome without a visible window |
| `--chrome-profile PATH` | none | Chrome user-data dir for saved login |

The template is never modified. Output is saved as `outputs/Season N amateur draft YYYY-MM-DD_HH-MM-SS.xlsx`.

### `apply-order` — Push your draft order to the site

Reads the ranked player list from your Excel's **Master List** sheet (or falls back to Hitters + Pitchers sorted by Overall Projection), opens the Rank Players popup on the site, reorders it to match, and saves.

```bash
python main.py apply-order                          # uses latest file in outputs/
python main.py "outputs/Season 30 amateur draft 2026-03-14_20-52-50.xlsx" apply-order
```

| Option | Default | Description |
|--------|---------|-------------|
| `--headless` | off | Run Chrome without a visible window |
| `--chrome-profile PATH` | none | Chrome user-data dir for saved login |

If no file is specified, the most recently modified `.xlsx` in `outputs/` is used automatically.

## Workflow

1. **Fetch** — Run `fetch` to populate your template with the current draft pool.
2. **Review** — Open the output in Excel. The Master List is pre-sorted by Adjusted Score. Tweak your template formulas or re-run fetch as needed.
3. **Apply** — Run `apply-order` to push the Master List order to the site's Rank Players list.

## Excel Sheet Layout

### Hitters (header row 6)

| Region | Columns | Contents |
|--------|---------|----------|
| Projection | A | Overall Projection (template formula, wrapped with `IFERROR`) |
| Hitting | B–P | Rnk, Player, Pos, B, T, Age, Contact, Power, vs L, vs R, Batting Eye, Baserunning, Arm, Bunt, Overall |
| Fielding | Q onward | Rank, Player, Pos, B, T, Age, Range, Glove, Arm Strength, Arm Accuracy, Pitch Calling, Durability, Health, Speed, Patience, Temper, Makeup, Overall |

### Pitchers (header row 5)

| Region | Columns | Contents |
|--------|---------|----------|
| Projection | A | Overall Projection (template formula, wrapped with `IFERROR`) |
| Ratings | B–S | Rank, Player, Position, B, T, Age, Durability, Stamina, Control, vsL, vsR, Velocity, GB/FB Tendency, Pitch 1–5, Overall |

### Master List (auto-generated)

| Column | Contents |
|--------|----------|
| A | **Adjusted Score** — Excel formula: `= B × D × E` |
| B | Overall Projection (formula referencing source sheet col A) |
| C | Raw Overall (HBD's raw rating) |
| D | Scouting Trust (multiplier from budget config) |
| E | Signability Factor (multiplier from signability text) |
| F | Player |
| G | Pos |
| H | Type (Hitter / Pitcher) |
| I | Category (college / high_school) |
| J | Signability (raw text) |

The Master List is sorted by Adjusted Score descending. Players with a zero score (unscouted, formula errors) are excluded.

### Background Info (auto-generated)

Rnk, Player, Pos, B, T, Age, Hometown, School, Class, Signability.

## Configuration (`credentials.env`)

All settings have sensible defaults. Omit any line to use the default.

### Login

| Key | Description |
|-----|-------------|
| `USERNAME` | WhatIfSports email |
| `PASSWORD` | WhatIfSports password |
| `HEADLESS` | `true` / `false` — run browser without a window (default: `false`) |

### Scouting Budgets

| Key | Default | Description |
|-----|---------|-------------|
| `SCOUTING_COLLEGE` | `0` | Your college scouting budget ($M, 0–20) |
| `SCOUTING_HIGH_SCHOOL` | `0` | Your high school scouting budget ($M, 0–20) |

Whichever category has the higher budget gets a trust factor of 1.0 (no penalty). The other category is penalized relative to it. Players are classified by age: 18 = high school, 19+ = college (including junior college). The `Class` field from Background Info is also used when available.

### Scouting Trust Formula

The trust multiplier for the lower-budget category is computed as:

```
trust = MIN_TRUST + (1 - MIN_TRUST) × (budget / 20) ^ CURVE
```

Then normalized so the higher category always equals 1.0.

| Key | Default | Description |
|-----|---------|-------------|
| `SCOUTING_MIN_TRUST` | `0.10` | Floor multiplier at $0 scouting (0.10 = 90% discount) |
| `SCOUTING_CURVE` | `0.17` | Exponent — lower values mean trust ramps up faster at low budgets |

The max scouting budget in HBD is always $20M.

### Signability Penalties

Each value is a multiplier (0.0–1.0) applied to the player's Adjusted Score on the Master List. 1.0 = no penalty, 0.0 = effectively excluded.

The "first round" and "first five rounds" penalties are conditional: if the player's raw overall rating meets the threshold, there is no penalty (they're good enough to justify the pick).

| Key | Default | Description |
|-----|---------|-------------|
| `SIGN_WILL_SIGN` | `1.0` | "will sign for slot" / "looking to sign" |
| `SIGN_FIRST_ROUND` | `0.90` | "drafted in the first round" (penalty if below threshold) |
| `SIGN_FIRST_ROUND_THRESHOLD` | `70` | No penalty if raw overall >= this |
| `SIGN_FIRST_FIVE` | `0.80` | "drafted in the first five rounds" (penalty if below threshold) |
| `SIGN_FIRST_FIVE_THRESHOLD` | `60` | No penalty if raw overall >= this |
| `SIGN_MAY_SIGN` | `0.60` | "may sign if the deal is right" |
| `SIGN_UNDECIDED` | `0.40` | "undecided" |
| `SIGN_PROBABLY_WONT` | `0.05` | "probably won't sign" |
| `SIGN_UNKNOWN` | `0.0` | "unknown" / "wasn't scouted" |
| `SIGN_FALLBACK` | `0.50` | Any other unrecognized signability text |

## Project Structure

```
hardball-dynasty-draft-optimizer/
├── main.py                 # CLI entry point (fetch / apply-order)
├── web_draft.py            # Selenium scraping, login, and Rank Players automation
├── excel_draft.py          # Excel reading/writing, Master List generation, COM sorting
├── credentials.py          # Loads config from credentials.env / env vars
├── requirements.txt        # Python dependencies
├── credentials.env.example # Template for credentials.env
├── credentials.env         # Your config (gitignored)
├── *.xlsx                  # Your Excel template (gitignored)
└── outputs/                # Generated output files (gitignored)
```

## Troubleshooting

- **Login:** The draft page requires a logged-in session. Run without `--headless` at least once and log in when the browser opens, or use `--chrome-profile` with a profile that's already logged in (close that profile's Chrome window first).
- **"Could not find draft prospects table":** The site's HTML may have changed. Check `web_draft.py` selectors (e.g. `table#dgPlayers`, button XPaths).
- **`#VALUE!` errors in Excel:** Column A formulas are automatically wrapped with `IFERROR(..., 0)` so errors display as 0 rather than `#VALUE!`.
- **Master List not sorted:** On Windows, the script uses Excel COM automation (`pywin32`) to recalculate formulas and sort. If `pywin32` is unavailable, open the file in Excel and sort the Master List by column A descending manually.
- **Slow apply-order:** The script uses JavaScript to reorder the Rank Players list instantly. If JS reorder fails, it falls back to button-clicking (slower but still works).
