# Hardball Dynasty Draft Optimizer

Automate your [WhatIfSports Hardball Dynasty](https://www.whatifsports.com/hbd/) amateur draft workflow: scrape the full prospect pool into a scored Excel workbook, then push your preferred draft order back to the site's "Rank Players" list with one command.

## Features

- Scrapes hitting, fielding, pitching, and background-info views into a single Excel file.
- Generates a **Master List** that ranks every prospect by an adjusted score factoring in your template's projection formula, scouting-budget trust, and signability risk.
- Applies that order to the site's Rank Players popup via Selenium (instant JavaScript reorder).
- All penalty weights and formula constants are configurable in `config.json` and `algorithm.json`.

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

5. **Credentials:** Copy `credentials.env.example` to `credentials.env` and fill in your login. This file is in `.gitignore` — never commit it.

   ```bash
   cp credentials.env.example credentials.env
   ```

6. **Config:** Copy `config.json.example` to `config.json` and adjust scouting budgets and signability penalties for your team.

   ```bash
   cp config.json.example config.json
   ```

## Commands

### `fetch` — Get data only (no formulas)

Pulls all prospect data from the Amateur Draft Player Pool page across four views (Hitting, Fielding/General, Pitching, Background Info), merges them, and writes **only the data** (Hitters, Pitchers, Background Info) to a timestamped file in `outputs/`. No algorithm formulas, no Master List — just the raw data to paste in. You apply the algorithm when you run **Sort master list** or **apply-order**.

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

### `apply-order` — Apply algorithm, sort Master List, then optionally push to the site

1. **Applies the algorithm** to the workbook: writes formulas from `algorithm.json` (including Control override and durability/stamina penalty) into the Hitters and Pitchers sheets, builds the **Master List**, and sorts it by Adjusted Score (via Excel COM on Windows).
2. Prompts: **Push this order to Hardball Dynasty?** — if you say yes (or use `--push`), opens the Rank Players popup and reorders it to match your Excel, then saves.

```bash
python main.py apply-order                          # uses latest file in outputs/
python main.py apply-order --push                   # sort and push without prompting
python main.py "outputs/Season 30 amateur draft 2026-03-14_20-52-50.xlsx" apply-order
```

| Option | Default | Description |
|--------|---------|-------------|
| `--push` | off | Push to web after sorting (skip the prompt) |
| `--headless` | off | Run Chrome without a visible window |
| `--chrome-profile PATH` | none | Chrome user-data dir for saved login |

If no file is specified, the most recently modified `.xlsx` in `outputs/` is used automatically.

## GUI and executable

A simple GUI runs the same workflow from buttons (no command line needed):

```bash
python gui_app.py
```

- **Fetch data** — Opens the browser, scrapes the draft pool, and writes only the data (Hitters, Pitchers, Background) to a timestamped Excel file in `outputs/`.
- **Sort master list** — Applies the algorithm (formulas, penalties) to the workbook, builds the Master List, and sorts it.
- **Push to Hardball Dynasty** — Pushes the current Excel order to the site (Rank Players).

Log output appears in the window. You can still use `main.py` from the command line for scripts or automation.

### Building a Windows executable

1. Install PyInstaller: `pip install pyinstaller`
2. From the project root, run: `build.bat` (or `pyinstaller --noconfirm --distpath . hardball_draft.spec`)
3. The executable is created at `HardballDraftOptimizer.exe` in the project root.

**Using the executable:** Place it in a folder alongside:
- `credentials.env` (your login; copy from `credentials.env.example`)
- `config.json` (game config; copy from `config.json.example`)
- Your Excel template: `Season x amateur draft-template.xlsx`

The first time you run Fetch, a browser window opens; log in to WhatIfSports if prompted. The app will create an `outputs` folder next to the exe for the generated Excel files. You can override the bundled `algorithm.json` by placing your own `algorithm.json` in the same folder as the exe.

## Workflow

1. **Fetch** — Run `fetch` to pull the draft pool from the site; only Hitters, Pitchers, and Background Info are written (no formulas or Master List).
2. **Sort master list** (or **apply-order**) — Applies `algorithm.json` (formulas, Control override, durability/stamina penalty), builds the Master List, and sorts it by Adjusted Score.
3. **Push** — Optionally push that order to the site's Rank Players list (via apply-order prompt or the GUI Push button).

## Excel Sheet Layout

### Hitters (header row 6)

| Region | Columns | Contents |
|--------|---------|----------|
| Projection | A | Overall Projection (formula generated from `algorithm.json`; wrapped with `IFERROR`) |
| Hitting | B–P | Rnk, Player, Pos, B, T, Age, Contact, Power, vs L, vs R, Batting Eye, Baserunning, Arm, Bunt, Overall |
| Fielding | Q onward | Rank, Player, Pos, B, T, Age, Range, Glove, Arm Strength, Arm Accuracy, Pitch Calling, Durability, Health, Speed, Patience, Temper, Makeup, Overall |
| Weights | Row 1 | Individual rating weights from `algorithm.json` (at each rating column) |
| Catcher weights | Row 2 | Alternate fielding weights for catchers |
| Group weights | Row 3 | Group weights at intermediate columns (AI–AM) |
| Intermediates | AI–AM | Computed group scores: hitting, baserunning, fielding, durability/health, intangibles |

### Pitchers (header row 5)

| Region | Columns | Contents |
|--------|---------|----------|
| Projection | A | Overall Projection (formula generated from `algorithm.json`; wrapped with `IFERROR`) |
| Ratings | B–S | Rank, Player, Position, B, T, Age, Durability, Stamina, Control, vsL, vsR, Velocity, GB/FB Tendency, Pitch 1–5, Overall |
| Weights | Row 1 | Individual rating weights from `algorithm.json` (at each rating column) |
| Group weights | Row 2 | Group weights at intermediate columns (U–W) |
| Intermediates | U–W | Computed group scores: pitching, pitches, durability/stamina |

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

The Master List is sorted by Adjusted Score descending when you run **apply-order** (Excel COM recalculates and sorts). Players with a zero score (unscouted, formula errors) are excluded.

### Background Info (auto-generated)

Rnk, Player, Pos, B, T, Age, Hometown, School, Class, Signability.

## Credentials (`credentials.env`)

Login and browser settings only. Copy `credentials.env.example` to `credentials.env`.

| Key | Description |
|-----|-------------|
| `USERNAME` | WhatIfSports email |
| `PASSWORD` | WhatIfSports password |
| `HEADLESS` | `true` / `false` — run browser without a window (default: `false`) |

## Game Configuration (`config.json`)

Scouting and signability settings for your team. Copy `config.json.example` to `config.json`. All values have sensible defaults — omit any key to use the default.

### Scouting

```json
"scouting": {
    "college": 0,
    "high_school": 10,
    "min_trust": 0.75,
    "curve": 0.17
}
```

| Key | Default | Description |
|-----|---------|-------------|
| `college` | `0` | Your college scouting budget ($M, 0–20) |
| `high_school` | `0` | Your high school scouting budget ($M, 0–20) |
| `min_trust` | `0.10` | Floor trust at $0 scouting (how much you trust unscouted ratings) |
| `curve` | `0.17` | Exponent — lower values mean trust ramps up faster at low budgets |

Whichever category has the higher budget gets a trust factor of 1.0 (no penalty). The other is penalized relative to it using: `trust = min_trust + (1 - min_trust) × (budget / 20) ^ curve`. Players are classified by age: 18 = high school, 19+ = college. The max budget in HBD is always $20M.

### Signability

```json
"signability": {
    "will_sign": 1.0,
    "first_round": 0.90,
    "first_round_threshold": 70,
    "first_five": 0.80,
    "first_five_threshold": 60,
    "may_sign": 0.60,
    "undecided": 0.40,
    "probably_wont": 0.05,
    "unknown": 0.0,
    "fallback": 0.50
}
```

Each value is a multiplier (0.0–1.0) applied to the player's Adjusted Score. 1.0 = no penalty, 0.0 = effectively excluded.

The `first_round` and `first_five` penalties are conditional: if the player's raw overall rating meets the threshold, no penalty is applied (they're good enough to justify the pick).

## Projection Algorithm (`algorithm.json`)

Everything that controls the Overall Projection formula lives in `algorithm.json`: the polynomial coefficients, every individual rating weight, group weights, and the method used for each group. If you delete the file, the script preserves whatever formulas are already in your template.

### How it works

1. Each rating is transformed through a **3rd-order polynomial**: `f(x) = a3·x³ + a2·x² + a1·x + a0`
2. The transformed rating is multiplied by its **individual weight**.
3. Weighted ratings are summed into **groups** (e.g. hitting, fielding, pitching).
4. Groups are combined using **group weights** and normalized against a "perfect player" reference row (all 100s) to produce the Overall Projection in column A.

Groups with `"method": "simple"` skip the polynomial and use a plain weighted average instead.

### Config structure

```json
{
    "polynomial": {
        "a3": -0.000002,
        "a2": 0.00032,
        "a1": -0.0021,
        "a0": 0
    },
    "hitters": {
        "groups": {
            "hitting": {
                "group_weight": 2.5,
                "method": "polynomial",
                "ratings": { "Contact": 1.2, "Power": 2.0, ... }
            },
            "fielding": {
                "group_weight": 2.0,
                "method": "polynomial",
                "ratings": { "Range": 1.0, "Glove": 1.0, ... },
                "catcher_condition": "Pitch Calling",
                "catcher_threshold": 50,
                "catcher_ratings": { "Glove": 0.2, "Arm Strength": 0.5, ... }
            },
            ...
        }
    },
    "pitchers": {
        "groups": { ... }
    }
}
```

### Polynomial coefficients

| Key | Default | Description |
|-----|---------|-------------|
| `a3` | `-0.000002` | Cubic coefficient |
| `a2` | `0.00032` | Quadratic coefficient |
| `a1` | `-0.0021` | Linear coefficient |
| `a0` | `0` | Constant term |

Default curve: `f(x) = -0.000002·x³ + 0.00032·x² - 0.0021·x` — compresses high ratings toward the top so that 80 and 90 are both close to the maximum (86% and 95% of perfect), while lower ratings like 60 and 70 are penalized more steeply (60% and 74%). The effect is that elite ratings are all treated as "good enough" while mediocre ones are clearly separated. Coefficients are also written to the Algorithm tab (columns M–P) for reference.

### Rating names

These must match the names used in `algorithm.json`:

**Hitters:** Contact, Power, vs L, vs R, Batting Eye, Baserunning, Arm, Bunt, Range, Glove, Arm Strength, Arm Accuracy, Pitch Calling, Durability, Health, Speed, Patience, Temper, Makeup

**Pitchers:** Durability, Stamina, Control, vsL, vsR, Velocity, GB/FB, Pitch 1, Pitch 2, Pitch 3, Pitch 4, Pitch 5

### Group properties

| Property | Required | Description |
|----------|----------|-------------|
| `group_weight` | Yes | How much this group contributes to the final score (only ratios matter between groups) |
| `method` | No | `"polynomial"` (default) or `"simple"` (plain weighted average, no polynomial) |
| `ratings` | Yes | `{name: weight}` — individual rating weights within the group |
| `catcher_condition` | No | Rating name used for the IF condition (e.g. `"Pitch Calling"`) |
| `catcher_threshold` | No | Threshold value for the condition (default 50) |
| `catcher_ratings` | No | Alternate `{name: weight}` used when the condition rating ≥ threshold |

The `catcher_*` fields let you define alternate fielding weights for catchers. When a player's Pitch Calling is ≥ the threshold, the `catcher_ratings` weights are used instead of `ratings`.

### Where weights are written

The script writes all weights from `algorithm.json` into the Excel sheet so they're visible:

- **Row 1:** Individual rating weights at their column (e.g. Contact weight → H1)
- **Row 2 (hitters only):** Catcher-specific weights
- **Row 3 (hitters) / Row 2 (pitchers):** Group weights at the intermediate columns (AI–AM for hitters, U–W for pitchers)

### Examples

Make fielding matter more for hitters:

```json
"hitting":  { "group_weight": 1.5, ... },
"fielding": { "group_weight": 3.0, ... }
```

Double the weight of Power:

```json
"ratings": { "Contact": 1.2, "Power": 4.0, "vs L": 1.0, ... }
```

Use a linear algorithm (disable the polynomial):

```json
"polynomial": { "a3": 0, "a2": 0, "a1": 1, "a0": 0 }
```

## Project Structure

```
hardball-dynasty-draft-optimizer/
├── main.py                 # CLI entry point (fetch / apply-order)
├── web_draft.py            # Selenium scraping, login, and Rank Players automation
├── excel_draft.py          # Excel reading/writing, Master List generation, COM sorting
├── credentials.py          # Loads credentials + config from their respective files
├── algorithm.json          # Projection algorithm: polynomial + all rating/group weights
├── config.json             # Scouting budgets, trust formula, signability penalties (gitignored)
├── config.json.example     # Template for config.json
├── requirements.txt        # Python dependencies
├── credentials.env.example # Template for credentials.env
├── credentials.env         # Login credentials (gitignored)
├── *.xlsx                  # Your Excel template (gitignored)
└── outputs/                # Generated output files (gitignored)
```

## Troubleshooting

- **Login:** The draft page requires a logged-in session. Run without `--headless` at least once and log in when the browser opens, or use `--chrome-profile` with a profile that's already logged in (close that profile's Chrome window first).
- **"Could not find draft prospects table":** The site's HTML may have changed. Check `web_draft.py` selectors (e.g. `table#dgPlayers`, button XPaths).
- **`#VALUE!` errors in Excel:** Column A formulas are automatically wrapped with `IFERROR(..., 0)` so errors display as 0 rather than `#VALUE!`.
- **Master List not sorted:** Run `apply-order` to recalculate and sort (uses Excel COM on Windows). If `pywin32` is unavailable, open the file in Excel and sort the Master List by column A descending manually.
- **Slow apply-order:** The script uses JavaScript to reorder the Rank Players list instantly. If JS reorder fails, it falls back to button-clicking (slower but still works).
