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

# Template in project root; override by passing a path as first argument.
DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "Season x amateur draft-template.xlsx"

# No default profile: script opens a separate Chrome window so you don't have to close your main Chrome.
# To reuse an existing login, use a dedicated profile: --chrome-profile "%LOCALAPPDATA%\Google\Chrome\User Data\Automation"
# (create "Automation" in Chrome's profile list, log in to HBD once there, close that Chrome before running).
DEFAULT_CHROME_PROFILE = None


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sync Hardball Dynasty amateur draft between Excel and whatifsports.com"
    )
    parser.add_argument(
        "excel_file",
        type=Path,
        nargs="?",
        default=DEFAULT_TEMPLATE,
        help=f"Path to draft Excel template (default: {DEFAULT_TEMPLATE.name} in project root)",
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

    # Apply Excel order TO web (Rank Players popup)
    p_apply = sub.add_parser(
        "apply-order",
        help="Open Rank Players popup and reorder list to match Excel (Hitters then Pitchers sheet order)",
    )
    p_apply.add_argument("--headless", action="store_true", help="Run browser headless")
    p_apply.add_argument(
        "--chrome-profile",
        type=Path,
        default=DEFAULT_CHROME_PROFILE,
        help="Chrome user data dir for saved login (optional). Without this, a new window opens; log in there when prompted.",
    )

    args = parser.parse_args()
    excel_path = args.excel_file.resolve()
    if not excel_path.exists():
        print(f"Excel file not found: {excel_path}", file=sys.stderr)
        sys.exit(1)
    from excel_draft import validate_template
    validation = validate_template(excel_path)
    if validation:
        print("Template validation failed:", file=sys.stderr)
        for msg in validation:
            print(f"  - {msg}", file=sys.stderr)
        sys.exit(1)

    from web_draft import run_sync_from_web_to_excel, run_apply_excel_order_to_web
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
        profile = str(args.chrome_profile) if getattr(args, "chrome_profile") and args.chrome_profile else None
        run_apply_excel_order_to_web(
            str(excel_path),
            headless=headless,
            user_data_dir=profile,
        )
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
