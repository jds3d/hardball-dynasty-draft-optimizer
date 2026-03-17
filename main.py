#!/usr/bin/env python3
"""
Hardball Dynasty Amateur Draft tool.

1) Populate Excel from web: fetch draft pool from whatifsports.com and fill your Excel file.
2) Apply Excel order to web: reorder the "Rank Players" list on the site to match your Excel order.

You must be logged in to whatifsports.com in the browser (or log in when the tool opens the page).
"""
import argparse
import sys
from pathlib import Path

from app_dir import get_app_dir

# Template in project root; override by passing a path as first argument.
DEFAULT_TEMPLATE = get_app_dir() / "Season x amateur draft-template.xlsx"
OUTPUTS_DIR = get_app_dir() / "outputs"

DEFAULT_CHROME_PROFILE = None


def _latest_output() -> Path | None:
    """Return the most recently modified .xlsx file in the outputs folder, or None."""
    if not OUTPUTS_DIR.is_dir():
        return None
    files = sorted(OUTPUTS_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sync Hardball Dynasty amateur draft between Excel and whatifsports.com"
    )
    parser.add_argument(
        "excel_file",
        type=Path,
        nargs="?",
        default=None,
        help="Path to Excel file. fetch: template (default: template in project root). "
             "apply-order: output file (default: latest file in outputs/).",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # Sync FROM web TO Excel
    p_fetch = sub.add_parser(
        "fetch",
        help="Fetch draft pool from web, populate Excel, and save to ./outputs/Season N amateur draft TIMESTAMP.xlsx",
    )
    p_fetch.add_argument(
        "--top",
        type=int,
        default=500,
        help="Number of prospects to fetch (default 500)",
    )
    p_fetch.add_argument(
        "--output-dir",
        type=Path,
        default=Path("outputs"),
        help="Folder to save the populated file (default: outputs)",
    )
    p_fetch.add_argument(
        "--headless",
        action="store_true",
        help="Run browser in headless mode (you cannot log in interactively)",
    )
    p_fetch.add_argument(
        "--chrome-profile",
        type=Path,
        default=DEFAULT_CHROME_PROFILE,
        help="Chrome user data dir for saved login (optional). Without this, a new window opens; log in there when prompted.",
    )

    # Apply order: reapply formula, sort Master List, then optionally push to web
    p_apply = sub.add_parser(
        "apply-order",
        help="Reapply configured formula, sort Master List by adjusted score, then optionally push order to Hardball Dynasty",
    )
    p_apply.add_argument("--push", action="store_true", help="Push order to web after sorting (otherwise only sort and ask)")
    p_apply.add_argument("--headless", action="store_true", help="Run browser headless")
    p_apply.add_argument(
        "--chrome-profile",
        type=Path,
        default=DEFAULT_CHROME_PROFILE,
        help="Chrome user data dir for saved login (optional). Without this, a new window opens; log in there when prompted.",
    )

    args = parser.parse_args()

    if args.excel_file is not None:
        excel_path = args.excel_file.resolve()
    elif args.command == "apply-order":
        latest = _latest_output()
        if latest:
            excel_path = latest
            print(f"Using latest output: {excel_path.name}")
        else:
            print("No output files found in outputs/. Run 'fetch' first or specify a file.", file=sys.stderr)
            sys.exit(1)
    else:
        excel_path = DEFAULT_TEMPLATE

    if not excel_path.exists():
        print(f"Excel file not found: {excel_path}", file=sys.stderr)
        sys.exit(1)

    from excel_draft import validate_template
    if args.command == "fetch":
        validation = validate_template(excel_path)
        if validation:
            print("Template validation failed:", file=sys.stderr)
            for msg in validation:
                print(f"  - {msg}", file=sys.stderr)
            sys.exit(1)

    from web_draft import run_sync_from_web_to_excel, run_apply_excel_order_to_web
    from excel_draft import reapply_formula_and_sort_master_list
    from credentials import get_headless

    headless = getattr(args, "headless", False) or get_headless()

    if args.command == "fetch":
        profile = str(args.chrome_profile) if getattr(args, "chrome_profile") and args.chrome_profile else None
        run_sync_from_web_to_excel(
            str(excel_path),
            headless=headless,
            user_data_dir=profile,
            top_n=getattr(args, "top", 500),
            output_dir=str(args.output_dir) if getattr(args, "output_dir", None) else "outputs",
        )
    elif args.command == "apply-order":
        sort_ok = reapply_formula_and_sort_master_list(excel_path)
        if not sort_ok:
            print("Could not sort Master List (Excel COM required on Windows). Open the file in Excel and sort by column A.", file=sys.stderr)
            sys.exit(1)
        print("Master list updated (formula recalculated and sorted by adjusted score).")
        do_push = getattr(args, "push", False)
        if not do_push:
            try:
                reply = input("Push this order to Hardball Dynasty? [y/N]: ").strip().lower()
                do_push = reply in ("y", "yes")
            except (EOFError, KeyboardInterrupt):
                do_push = False
        if do_push:
            profile = str(args.chrome_profile) if getattr(args, "chrome_profile") and args.chrome_profile else None
            run_apply_excel_order_to_web(
                str(excel_path),
                headless=headless,
                user_data_dir=profile,
            )
        else:
            print("Skipped pushing to web.")
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
