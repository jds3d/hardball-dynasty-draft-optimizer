"""
Browser automation for Hardball Dynasty Amateur Draft:
- Fetch draft pool table from whatifsports.com
- Reorder players in the Rank Players popup to match Excel order
"""
import logging
import re
import time
from datetime import datetime
from typing import Any

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

DRAFT_POOL_URL = "https://www.whatifsports.com/hbd/Pages/GM/AmateurDraftPlayerPool.aspx"

# Real column names for rating columns when the page gives no title/alt (icons only).
# Keys are fallback names like "Rating_7" (7 = 1-based column index in the table).
# Fill these in after a fetch: check the Excel for any "Rating_N" headers and map them here.
# Hitting view: columns after Rnk, Player, Pos, B, T, Age (so first rating column is Rating_7).
HITTER_RATING_DISPLAY_NAMES: dict[str, str] = {}
# Fielding/General view: same idea; merge adds only columns not already in hitting row.
FIELDING_RATING_DISPLAY_NAMES: dict[str, str] = {}
# Pitching view: rating columns for pitchers.
PITCHER_RATING_DISPLAY_NAMES: dict[str, str] = {}


def _get_chrome_driver(headless: bool = False, user_data_dir: str | None = None) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def _wait(driver: webdriver.Chrome, timeout: float = 15) -> WebDriverWait:
    return WebDriverWait(driver, timeout)


def _cell_header_label(cell) -> str:
    """Get header label from cell: visible text, then title attribute, then img title/alt (for icon columns)."""
    text = (cell.text or "").strip()
    if text:
        return text
    title = (cell.get_attribute("title") or "").strip()
    if title:
        return title
    try:
        for img in cell.find_elements(By.TAG_NAME, "img"):
            t = (img.get_attribute("title") or img.get_attribute("alt") or "").strip()
            if t:
                return t
    except Exception:
        pass
    return ""


def _get_first_row_headers(table) -> list[str]:
    """Get header labels from the table's first row (th or td). Uses title/img alt when cell text is empty (icons)."""
    try:
        all_tr = table.find_elements(By.TAG_NAME, "tr")
        for tr in all_tr[:2]:  # first or second row
            cells = tr.find_elements(By.TAG_NAME, "th")
            if not cells:
                cells = tr.find_elements(By.TAG_NAME, "td")
            labels = [_cell_header_label(c) for c in cells]
            # Use this row if it has Rnk/Rank and Player
            if any("rank" in (l or "").lower() or l in ("Rnk", "Rank") for l in labels):
                if any("player" in (l or "").lower() for l in labels):
                    return labels
        # Fallback: return first row's labels anyway
        if all_tr:
            cells = all_tr[0].find_elements(By.TAG_NAME, "th") or all_tr[0].find_elements(By.TAG_NAME, "td")
            return [_cell_header_label(c) for c in cells]
    except Exception:
        pass
    return []


def _find_draft_table(driver: webdriver.Chrome):
    """Find the Draft Prospects table by locating a table whose first row contains Rnk and Player."""
    tables = driver.find_elements(By.TAG_NAME, "table")
    for table in tables:
        try:
            header_cells = _get_first_row_headers(table)
            if not header_cells:
                continue
            # Accept Rnk/Rank and Player/Player Name (case-insensitive, allow partial)
            header_lower = [h.lower() for h in header_cells if h]
            has_rank = any(h in ("rnk", "rank") or "rank" in h for h in header_lower)
            has_player = any(h == "player" or "player" in h for h in header_lower)
            if has_rank and has_player:
                log.info("Found draft table (%s header cells).", len(header_cells))
                return table
        except Exception:
            continue
    return None


def _table_to_rows(
    driver: webdriver.Chrome,
    table_selector: str = "table#dgPlayers",
    rating_display_names: dict[str, str] | None = None,
    key_prefix: str = "Rating",
) -> list[dict[str, Any]]:
    """Parse the draft prospects table into a list of dicts (header -> cell text)."""
    rows: list[dict[str, Any]] = []
    table = _find_draft_table(driver)
    if not table:
        try:
            table = _wait(driver).until(EC.presence_of_element_located((By.CSS_SELECTOR, table_selector)))
        except Exception:
            for sel in ["table[id*='dgPlayers']", "table[id*='Players']", "table.grid"]:
                try:
                    table = driver.find_element(By.CSS_SELECTOR, sel)
                    break
                except Exception:
                    continue
        if not table:
            raise RuntimeError("Could not find draft prospects table (no table with Rnk and Player columns).")
    # Use same logic as _find_draft_table: first row = header (th or td)
    header_cells = _get_first_row_headers(table)
    if not header_cells:
        raise RuntimeError("Draft table has no header row.")
    # All rows except the first are data (in case there's no tbody/thead, or header is in tbody)
    all_rows = table.find_elements(By.TAG_NAME, "tr")
    body_rows = all_rows[1:] if len(all_rows) > 1 else []
    # If the site uses a two-row header (text in row 1, icons in row 2), we only got 6 labels.
    # Pad to the data row's column count so we scrape every column (Rating_7, Rating_8, ...).
    for tr in body_rows:
        cells = tr.find_elements(By.TAG_NAME, "td")
        if len(cells) < len(header_cells):
            continue
        first_text = (cells[0].text.strip() if cells else "").lower()
        second_text = (cells[1].text.strip() if len(cells) > 1 else "").lower()
        if first_text in ("rnk", "rank") or second_text == "player":
            continue
        if len(cells) > len(header_cells):
            header_cells = list(header_cells) + [""] * (len(cells) - len(header_cells))
        break
    for tr in body_rows:
        cells = tr.find_elements(By.TAG_NAME, "td")
        if len(cells) < len(header_cells):
            continue
        # Skip header row if it appears in tbody (e.g. first cell "Rnk"/"Rank" or second "Player")
        first_text = (cells[0].text.strip() if cells else "").lower()
        second_text = (cells[1].text.strip() if len(cells) > 1 else "").lower()
        if first_text in ("rnk", "rank") or second_text == "player":
            continue
        row = {}
        for i, h in enumerate(header_cells):
            if i >= len(cells):
                continue
            raw = cells[i].text.strip()
            parsed = _parse_cell(raw)
            positional_key = f"{key_prefix}_{i + 1}"
            row[positional_key] = parsed
            # Also store under the header-derived name if we have one
            name_key = (h or "").strip()
            if name_key and name_key != positional_key:
                row[name_key] = parsed
            # Apply display name overrides
            if rating_display_names and positional_key in rating_display_names:
                row[rating_display_names[positional_key]] = parsed
        # Skip rows with no player name (e.g. second header row with icons)
        if not row or (not row.get("Player") and not row.get("Player Name")):
            continue
        if not rows:
            log.info("First scraped row keys: %s", list(row.keys()))
            log.info("First scraped row values: %s", row)
        rows.append(row)
    log.info("Header labels from page: %s", header_cells)
    return rows


def get_season_from_page(driver: webdriver.Chrome) -> int | None:
    """
    Parse the current season number from the page (e.g. 'Strawberry-Gooden (30) - Scottsdale' -> 30).
    Tries visible text first, then individual elements, then raw page source HTML.
    Returns None if not found.
    """
    # Strategy 1: visible body text
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        text = body.text or ""
        match = re.search(r"\((\d{1,3})\)", text)
        if match and 1 <= int(match.group(1)) <= 999:
            return int(match.group(1))
    except Exception:
        pass
    # Strategy 2: individual elements with parenthesized numbers
    try:
        for el in driver.find_elements(By.XPATH, "//*[contains(text(),'(') and contains(text(),')')]"):
            t = (el.text or "").strip()
            m = re.search(r"\((\d{1,3})\)", t)
            if m and 1 <= int(m.group(1)) <= 999:
                return int(m.group(1))
    except Exception:
        pass
    # Strategy 3: raw page source (works in headless when visible text is incomplete)
    try:
        source = driver.page_source or ""
        # Look for team/league pattern like "TeamName (30) - CityName"
        m = re.search(r'[A-Z][a-zA-Z\-]+\s+\((\d{1,3})\)\s*-\s*[A-Z]', source)
        if m and 1 <= int(m.group(1)) <= 999:
            log.info("Season found in page source: %s", m.group(1))
            return int(m.group(1))
        # Broader: any (NN) in source that's near "Season" or league context
        for pattern in [r'Season\s+(\d{1,3})', r'\((\d{1,3})\)']:
            m = re.search(pattern, source)
            if m and 1 <= int(m.group(1)) <= 999:
                log.info("Season found in page source (broad): %s", m.group(1))
                return int(m.group(1))
    except Exception:
        pass
    return None


def _parse_cell(raw: str) -> str | int | float:
    if not raw:
        return raw
    try:
        return int(raw)
    except ValueError:
        pass
    try:
        return float(raw)
    except ValueError:
        pass
    return raw


def _click_go(driver: webdriver.Chrome) -> bool:
    """Click the GO button to apply Show/View/Position filters. Returns True if clicked."""
    # Collect all candidate buttons/inputs and click the one that looks like GO
    try:
        for tag in ["input", "button", "a"]:
            for el in driver.find_elements(By.TAG_NAME, tag):
                try:
                    if not el.is_displayed() or not el.is_enabled():
                        continue
                    val = (el.get_attribute("value") or "").strip()
                    txt = (el.text or "").strip()
                    if val.upper() == "GO" or txt.upper() == "GO":
                        el.click()
                        log.info("Clicked GO button.")
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    # XPath fallbacks (value, id, name, placeholder)
    for by, sel in [
        (By.XPATH, "//input[@value='GO' or @value='Go']"),
        (By.XPATH, "//button[normalize-space(.)='GO' or normalize-space(.)='Go']"),
        (By.XPATH, "//*[normalize-space(.)='GO']"),
        (By.XPATH, "//input[contains(translate(@id,'GO','go'),'go') or contains(translate(@name,'GO','go'),'go')]"),
        (By.XPATH, "//a[contains(.,'GO') or contains(.,'Go')]"),
        (By.XPATH, "//input[@type='image' and contains(@alt,'Go')]"),
    ]:
        try:
            for btn in driver.find_elements(by, sel):
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    log.info("Clicked GO button.")
                    return True
        except Exception:
            continue
    log.warning("GO button not found; continuing without applying filters.")
    return False


def _go_and_wait_for_table(driver: webdriver.Chrome, expected_view: str = "", timeout: float = 20) -> None:
    """
    Click GO and wait for the ASP.NET postback to complete and a fresh draft table to appear.
    Uses staleness_of on the old table to detect postback, then waits for a new table.
    """
    old_table = _find_draft_table(driver)
    old_table_stale = False

    clicked = _click_go(driver)
    if not clicked:
        log.warning("GO not clicked; attempting JavaScript form submit as fallback...")
        try:
            driver.execute_script("__doPostBack('', '');")
        except Exception:
            pass

    if old_table:
        try:
            WebDriverWait(driver, timeout).until(EC.staleness_of(old_table))
            old_table_stale = True
            log.info("Old table went stale (postback completed for %s view).", expected_view)
        except Exception:
            log.info("Old table did not go stale within %ss for %s view.", timeout, expected_view)

    # Poll for the new draft table to appear
    for attempt in range(20):
        table = _find_draft_table(driver)
        if table:
            if old_table_stale or (table is not old_table) or old_table is None:
                log.info("Draft table ready for %s view after %s polls.", expected_view, attempt + 1)
                time.sleep(0.5)
                return
        time.sleep(1)

    # Last resort: table may exist even if we couldn't confirm it refreshed
    if _find_draft_table(driver):
        log.warning("Using existing table for %s view (could not confirm postback).", expected_view)
    else:
        log.warning("No draft table found for %s view after polling.", expected_view)


def _set_dropdown(driver: webdriver.Chrome, option_text: str, select_hints: list[str]) -> bool:
    """Set a dropdown by finding a select that contains the given option. Returns True if set."""
    for hint in select_hints:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, f"select[id*='{hint}'], select[name*='{hint}']"):
                try:
                    sel = Select(el)
                    sel.select_by_visible_text(option_text)
                    return True
                except Exception:
                    continue
        except Exception:
            continue
    # Try any select that has this option
    try:
        for el in driver.find_elements(By.TAG_NAME, "select"):
            try:
                sel = Select(el)
                sel.select_by_visible_text(option_text)
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _norm_name(name: str) -> str:
    """Normalize player name for matching."""
    return " ".join((name or "").strip().split()).lower()


def fetch_draft_pool_data(
    driver: webdriver.Chrome,
    top_n: int = 500,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """
    Fetch draft pool: hitters get both Hitting and Fielding/General views merged; pitchers get Pitching view.
    Returns (hitters_rows, pitchers_rows, background_rows).
    """
    log.info("Fetch: Loading draft pool URL...")
    driver.get(DRAFT_POOL_URL)
    _wait(driver).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    log.info("Fetch: Page loaded, table element found.")
    time.sleep(0.5)

    _col_map = {
        "Player Name": "Player",
        "Name": "Player",
        "Rank": "Rnk",
        "#": "Rnk",
        "Position": "Pos",
        "POS": "Pos",
        "Bats": "B",
        "Throws": "T",
    }

    def norm(r: dict) -> dict:
        return {_col_map.get(k.strip(), k.strip()): v for k, v in r.items() if k and str(k).strip()}

    # --- Hitting view (first table for hitters) ---
    log.info("Fetch: Setting Show to Top %s...", top_n)
    _set_dropdown(driver, f"Top {top_n}", ["Top", "ddlTop", "Show"])
    log.info("Fetch: Setting View to Projected Hitting Ratings, clicking GO...")
    _set_dropdown(driver, "Projected Hitting Ratings", ["View", "ddlView"])
    _go_and_wait_for_table(driver, "Hitting")
    try:
        hitting_view_rows = _table_to_rows(driver, rating_display_names=HITTER_RATING_DISPLAY_NAMES)
    except Exception as e:
        log.warning("Fetch: Could not parse hitting table: %s", e)
        hitting_view_rows = []
    hitting_view_rows = [norm(r) for r in hitting_view_rows]
    log.info("Fetch: Parsed %s hitting view rows.", len(hitting_view_rows))

    # --- Fielding/General view (merge with hitting for hitters) ---
    log.info("Fetch: Setting View to Projected Fielding/General Ratings, clicking GO...")
    _set_dropdown(driver, "Projected Fielding/General Ratings", ["View", "ddlView"])
    _go_and_wait_for_table(driver, "Fielding")
    try:
        fielding_view_rows = _table_to_rows(driver, rating_display_names=FIELDING_RATING_DISPLAY_NAMES, key_prefix="Fielding")
    except Exception as e:
        log.warning("Fetch: Could not parse fielding table: %s", e)
        fielding_view_rows = []
    fielding_view_rows = [norm(r) for r in fielding_view_rows]
    log.info("Fetch: Parsed %s fielding/general view rows.", len(fielding_view_rows))

    # Merge hitting + fielding by player name (hitting base; add fielding columns that don't conflict)
    fielding_by_name = {_norm_name(_player_name(r)): r for r in fielding_view_rows}
    merged_hitter_view: list[dict[str, Any]] = []
    for row in hitting_view_rows:
        merged = dict(row)
        name_key = _norm_name(_player_name(row))
        if name_key in fielding_by_name:
            for k, v in fielding_by_name[name_key].items():
                if k not in merged:
                    merged[k] = v
        merged_hitter_view.append(merged)
    log.info("Fetch: Merged hitting + fielding into %s rows.", len(merged_hitter_view))

    # --- Pitching view ---
    log.info("Fetch: Setting View to Projected Pitching Ratings (pitchers), clicking GO...")
    _set_dropdown(driver, "Projected Pitching Ratings", ["View", "ddlView"])
    _go_and_wait_for_table(driver, "Pitching")
    try:
        pitching_rows = _table_to_rows(driver, rating_display_names=PITCHER_RATING_DISPLAY_NAMES)
    except Exception as e:
        log.warning("Fetch: Could not parse pitching table: %s", e)
        pitching_rows = []
    pitching_rows = [norm(r) for r in pitching_rows]
    log.info("Fetch: Parsed %s pitching view rows.", len(pitching_rows))

    # --- Background Info view (signability, school, class) ---
    log.info("Fetch: Setting View to Background Info, clicking GO...")
    _set_dropdown(driver, "Background Info", ["View", "ddlView"])
    _go_and_wait_for_table(driver, "Background")
    try:
        background_rows = _table_to_rows(driver, key_prefix="BG")
    except Exception as e:
        log.warning("Fetch: Could not parse background info table: %s", e)
        background_rows = []
    background_rows = [norm(r) for r in background_rows]
    log.info("Fetch: Parsed %s background info rows.", len(background_rows))

    # Merge signability + class into hitter/pitcher rows by player name
    bg_by_name = {_norm_name(_player_name(r)): r for r in background_rows}
    for row in merged_hitter_view:
        bg = bg_by_name.get(_norm_name(_player_name(row)), {})
        row["Signability"] = bg.get("Signability", "")
        row["Class"] = bg.get("Class", "")
        row["School"] = bg.get("School", "")
        row["Hometown"] = bg.get("Hometown", "")
    for row in pitching_rows:
        bg = bg_by_name.get(_norm_name(_player_name(row)), {})
        row["Signability"] = bg.get("Signability", "")
        row["Class"] = bg.get("Class", "")
        row["School"] = bg.get("School", "")
        row["Hometown"] = bg.get("Hometown", "")

    # Split by position: Hitters = non-P from merged view, Pitchers = P from pitching view
    hitters = [r for r in merged_hitter_view if _pos(r) != "P"]
    pitchers = [r for r in pitching_rows if _pos(r) == "P"]
    log.info("Fetch: Split into %s hitters and %s pitchers.", len(hitters), len(pitchers))
    return hitters, pitchers, background_rows


def _pos(row: dict) -> str:
    for key in ("Pos", "Position", "POS"):
        if key in row and row[key]:
            return str(row[key]).strip().upper()
    return ""


def _player_name(row: dict) -> str:
    for key in ("Player", "Player Name", "Name"):
        if key in row and row[key]:
            return str(row[key]).strip()
    return ""


def get_current_rank_order_from_popup(driver: webdriver.Chrome) -> list[str]:
    """Get the current ordered list of player names from the Rank Players popup. Call when popup is open."""
    names: list[str] = []
    # List might be in a listbox, ul/li, or table
    for sel in [
        "select[id*='Rank'] option",
        "ul[id*='Rank'] li",
        "[id*='Rank'] li",
        ".rank-list li",
        "select option",
        "table tbody tr",
    ]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                text = el.text.strip()
                # Match "1. Steve Wilson (2B)" or "Steve Wilson (2B)"
                if text and re.match(r"^(\d+\.\s*)?(.+)\s*\([A-Z0-9]+\)\s*$", text):
                    name_part = re.sub(r"^\d+\.\s*", "", text)
                    name_part = re.sub(r"\s*\([A-Z0-9]+\)\s*$", "", name_part).strip()
                    if name_part:
                        names.append(name_part)
            if names:
                return names
        except Exception:
            continue
    # Fallback: any list-like text (number. Name (Pos))
    for el in driver.find_elements(By.XPATH, "//*[contains(text(), '(')]"):
        t = el.text.strip()
        m = re.match(r"^(\d+\.\s*)?(.+?)\s*\([A-Z0-9]+\)", t)
        if m:
            names.append(m.group(2).strip())
    return names


def _normalize_name_for_match(name: str) -> str:
    return " ".join(name.split()).strip()


def apply_draft_order_in_popup(
    driver: webdriver.Chrome,
    desired_order: list[str],
) -> None:
    """
    In the Rank Players popup, reorder the list to match desired_order.
    desired_order[0] = first pick, etc.
    """
    desired_norm = [_normalize_name_for_match(n) for n in desired_order]
    # Find move up / move down buttons by text or value
    def get_current_list() -> list[str]:
        return get_current_rank_order_from_popup(driver)

    for target_index in range(len(desired_norm)):
        desired_name = desired_norm[target_index]
        current = get_current_list()
        current_norm = [_normalize_name_for_match(n) for n in current]
        try:
            current_pos = current_norm.index(desired_name)
        except ValueError:
            continue
        if current_pos <= target_index:
            continue
        # Select the player (click list item)
        try:
            items = driver.find_elements(By.CSS_SELECTOR, "select option, ul li, table tbody tr")
            if current_pos < len(items):
                items[current_pos].click()
                time.sleep(0.15)
        except Exception:
            pass
        # Click move up (current_pos - target_index) times
        for _ in range(current_pos - target_index):
            try:
                btn = driver.find_element(By.XPATH, "//input[@value='↑'] | //button[contains(.,'↑')] | //a[contains(.,'↑')]")
                btn.click()
                time.sleep(0.2)
            except Exception:
                break


def open_rank_players_popup(driver: webdriver.Chrome) -> None:
    """Click the Rank Players button (grey button below Formula Builder) and wait for the popup."""
    for by, sel in [
        (By.CSS_SELECTOR, "input[value='Rank Players']"),
        (By.XPATH, "//input[@value='Rank Players']"),
        (By.XPATH, "//input[contains(@value,'Rank')]"),
        (By.XPATH, "//button[contains(.,'Rank Players') or contains(.,'Rank players')]"),
        (By.XPATH, "//a[contains(.,'Rank Players') or contains(.,'Rank players')]"),
        (By.XPATH, "//*[contains(normalize-space(.),'Rank Players')]"),
        (By.LINK_TEXT, "Rank Players"),
        (By.PARTIAL_LINK_TEXT, "Rank Players"),
    ]:
        try:
            btn = driver.find_element(by, sel)
            if btn.is_displayed() and btn.is_enabled():
                btn.click()
                time.sleep(1)
                return
        except Exception:
            continue
    raise RuntimeError("Could not find 'Rank Players' button. Check the Draft Prospects page has the grey Rank Players button.")


def save_rank_players_popup(driver: webdriver.Chrome) -> None:
    """Click Save in the Rank Players popup."""
    for by, sel in [
        (By.CSS_SELECTOR, "input[value='Save']"),
        (By.XPATH, "//input[@value='Save']"),
        (By.XPATH, "//button[contains(.,'Save')]"),
        (By.XPATH, "//*[contains(text(),'Save')]"),
    ]:
        try:
            btn = driver.find_element(by, sel)
            btn.click()
            time.sleep(0.5)
            return
        except Exception:
            continue
    raise RuntimeError("Could not find 'Save' button in popup.")


def _try_auto_login(driver: webdriver.Chrome) -> bool:
    """
    WhatIfSports uses a two-step login: (1) email + Continue, (2) password + submit.
    If the page shows either step, fill and proceed. Returns True if login was attempted.
    """
    log.info("Checking for login form and credentials...")
    creds = None
    try:
        from credentials import get_hbd_credentials
        creds = get_hbd_credentials()
    except Exception as e:
        log.debug("Could not load credentials: %s", e)
    if not creds:
        log.info("No credentials found; skipping auto-login.")
        return False
    username, password = creds
    log.info("Credentials loaded. Looking for email field...")
    try:
        # Step 1: Find email field ("Enter your email address" / Email address label)
        user_input = None
        for sel in [
            "input[type='email']",
            "input[placeholder*='email']",
            "input[placeholder*='Email']",
            "input[name*='email']",
            "input[name*='user']",
            "input[id*='email']",
        ]:
            try:
                for el in driver.find_elements(By.CSS_SELECTOR, sel):
                    if el.is_displayed() and el.is_enabled():
                        user_input = el
                        break
                if user_input:
                    break
            except Exception:
                continue
        if not user_input:
            log.info("No email field found; not on login step 1.")
            return False
        log.info("Step 1: Filled email, clicking Continue...")
        user_input.clear()
        user_input.send_keys(username)
        # Click "Continue" (orange button) to go to password step
        continue_clicked = False
        for by, sel in [
            (By.XPATH, "//button[contains(.,'Continue')]"),
            (By.XPATH, "//*[contains(.,'Continue') and (self::button or self::input)]"),
            (By.CSS_SELECTOR, "button[type='submit']"),
        ]:
            try:
                for btn in driver.find_elements(by, sel):
                    if "Continue" in (btn.text or "") and btn.is_displayed() and btn.is_enabled():
                        btn.click()
                        continue_clicked = True
                        break
                if continue_clicked:
                    break
            except Exception:
                continue
        if not continue_clicked:
            from selenium.webdriver.common.keys import Keys
            user_input.send_keys(Keys.RETURN)
        log.info("Waiting for password step...")
        time.sleep(2)
        # Step 2: Password field appears after Continue
        pass_input = None
        for sel in ["input[type='password']", "input[name*='pass']", "input[id*='pass']"]:
            try:
                for el in driver.find_elements(By.CSS_SELECTOR, sel):
                    if el.is_displayed() and el.is_enabled():
                        pass_input = el
                        break
                if pass_input:
                    break
            except Exception:
                continue
        if not pass_input:
            log.info("No password field found after Continue (page may still be loading).")
            return True  # We at least submitted email; might be loading or different flow
        log.info("Step 2: Filled password, submitting...")
        pass_input.clear()
        pass_input.send_keys(password)
        # Click Sign in / Login / Submit
        submitted = False
        for by, sel in [
            (By.XPATH, "//button[contains(.,'Sign in') or contains(.,'Login') or contains(.,'Log in')]"),
            (By.XPATH, "//input[@value='Sign in' or @value='Login' or @value='Log in']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "input[type='submit']"),
        ]:
            try:
                for btn in driver.find_elements(by, sel):
                    if btn.is_displayed() and btn.is_enabled():
                        btn.click()
                        submitted = True
                        break
                if submitted:
                    break
            except Exception:
                continue
        if not submitted:
            from selenium.webdriver.common.keys import Keys
            pass_input.send_keys(Keys.RETURN)
        log.info("Login form submitted.")
        time.sleep(2)
        return True
    except Exception as e:
        log.warning("Auto-login failed: %s", e)
        return False


def _click_link_or_button(driver: webdriver.Chrome, text: str) -> bool:
    """Click a link or button that contains the given text. Returns True if found and clicked."""
    text = text.strip()
    for by, sel in [
        (By.LINK_TEXT, text),
        (By.PARTIAL_LINK_TEXT, text),
        (By.XPATH, f"//a[contains(.,'{text}')]"),
        (By.XPATH, f"//button[contains(.,'{text}')]"),
        (By.XPATH, f"//*[contains(normalize-space(.),'{text}') and (self::a or self::button or self::input)]"),
    ]:
        try:
            for el in driver.find_elements(by, sel):
                if el.is_displayed() and el.is_enabled():
                    el.click()
                    return True
        except Exception:
            continue
    return False


def _navigate_to_draft_pool(driver: webdriver.Chrome) -> None:
    """
    After login, the site may show World Center or Franchise Center.
    Click "View Your Franchises" then "Visit Team Office!" then load draft pool URL.
    """
    time.sleep(1.5)
    # Step 1: If "View Your Franchises" is visible (World Center), click it
    if _click_link_or_button(driver, "View Your Franchises"):
        log.info("Navigation: Clicked 'View Your Franchises', waiting for Franchise Center...")
        time.sleep(2)
    # Step 2: On Franchise Center, click "Visit Team Office!" (or "Visit Team Office")
    if _click_link_or_button(driver, "Visit Team Office!"):
        log.info("Navigation: Clicked 'Visit Team Office!', waiting...")
        time.sleep(2)
    elif _click_link_or_button(driver, "Visit Team Office"):
        log.info("Navigation: Clicked 'Visit Team Office', waiting...")
        time.sleep(2)
    # Step 3: Go to draft pool page
    log.info("Navigation: Loading draft pool URL...")
    driver.get(DRAFT_POOL_URL)
    time.sleep(1.5)
    try:
        _wait(driver).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        log.info("Draft prospects page loaded (table found).")
    except Exception as e:
        log.warning("Waiting for draft table failed: %s (continuing anyway)", e)


def _wait_for_login(driver: webdriver.Chrome) -> None:
    """Open draft pool page; auto-login if credentials are set and login form is shown, else wait for user."""
    log.info("Opening draft pool URL: %s", DRAFT_POOL_URL)
    driver.get(DRAFT_POOL_URL)
    time.sleep(1.5)
    if _try_auto_login(driver):
        log.info("Auto-login completed; navigating to draft prospects page...")
        _navigate_to_draft_pool(driver)
        return
    # Manual login: after user presses Enter, we may still be on World Center / Franchise Center
    log.info("Checking if we need to navigate from World Center / Franchise Center...")
    _navigate_to_draft_pool(driver)


def run_sync_from_web_to_excel(
    excel_path: str,
    headless: bool = False,
    user_data_dir: str | None = None,
    top_n: int = 500,
    output_dir: str | None = "outputs",
) -> None:
    """
    Open browser, fetch draft pool data, write to Excel in ./outputs/.
    Season number is read from the page (e.g. "Strawberry-Gooden (30) - Scottsdale" -> 30).
    """
    from pathlib import Path
    from excel_draft import write_draft_data_to_excel

    log.info("Starting fetch: opening browser...")
    driver = _get_chrome_driver(headless=headless, user_data_dir=user_data_dir)
    try:
        _wait_for_login(driver)
        log.info("Fetch: Reading season from page (before view switches)...")
        season = get_season_from_page(driver)
        log.info("Fetch: Season detected: %s", season)
        log.info("Fetch: Getting draft pool data...")
        hitters, pitchers, background = fetch_draft_pool_data(driver, top_n=top_n)
        out_dir = Path(output_dir or "outputs")
        out_dir.mkdir(parents=True, exist_ok=True)
        season_label = str(season) if season is not None else "unknown"
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_path = out_dir / f"Season {season_label} amateur draft {timestamp}.xlsx"
        log.info("Fetch: Writing to %s...", output_path)
        write_draft_data_to_excel(excel_path, hitters, pitchers, background_rows=background, output_path=output_path)
        log.info("Fetch: Done. Wrote %s hitters and %s pitchers to %s", len(hitters), len(pitchers), output_path)
        print(f"Wrote {len(hitters)} hitters and {len(pitchers)} pitchers to {output_path}")
    finally:
        driver.quit()


def _switch_to_popup_if_new_window(driver: webdriver.Chrome) -> None:
    """If Rank Players opened a new window, switch to it."""
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])


def run_apply_excel_order_to_web(
    excel_path: str,
    headless: bool = False,
    user_data_dir: str | None = None,
) -> None:
    """Open browser, go to draft pool, open Rank Players, reorder to match Excel, Save."""
    from excel_draft import get_draft_order_from_excel

    log.info("Apply-order: Reading draft order from Excel...")
    desired = get_draft_order_from_excel(excel_path)
    if not desired:
        log.warning("No draft order found in Excel.")
        print("No draft order found in Excel.")
        return
    log.info("Apply-order: %s players in desired order. Opening browser...", len(desired))

    driver = _get_chrome_driver(headless=headless, user_data_dir=user_data_dir)
    try:
        _wait_for_login(driver)
        log.info("Apply-order: Waiting for draft table...")
        _wait(driver).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(0.5)
        log.info("Apply-order: Opening Rank Players popup...")
        open_rank_players_popup(driver)
        time.sleep(1)
        _switch_to_popup_if_new_window(driver)
        log.info("Apply-order: Reordering list to match Excel...")
        apply_draft_order_in_popup(driver, desired)
        log.info("Apply-order: Saving...")
        save_rank_players_popup(driver)
        log.info("Apply-order: Done.")
        print("Rank order applied and saved.")
    finally:
        driver.quit()
