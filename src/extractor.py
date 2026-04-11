"""
Strategy Tester Report Extractor
Reads all .htm files from ../input, extracts key metrics, and writes to ../output/summary.xlsx
"""

import os
import re
from pathlib import Path

from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


COLUMNS = [
    "Filename",
    "Symbol",
    "Period",
    "Modelling quality",
    "Spread",
    "Initial deposit",
    "Total net profit",
    "Gross profit",
    "Gross loss",
    "Profit factor",
    "Expected payoff",
    "Absolute drawdown",
    "Maximal drawdown",
    "Relative drawdown",
    "Total trades",
    "Short positions",
    "Short positions won %",
    "Long positions",
    "Long positions won %",
    "Largest profit trade",
    "Largest loss trade",
    "Average profit trade",
    "Average loss trade",
    "Parameters",
]


def get_td_text(td):
    return td.get_text(separator=" ", strip=True) if td else ""


def parse_positions(raw):
    """Parse '35 (100.00%)' into ('35', '100.00%')."""
    m = re.match(r"(\d+)\s*\(([^)]+)\)", raw)
    if m:
        return m.group(1), m.group(2)
    return raw, ""


def parse_period(period_raw: str, filename: str) -> str:
    """Return 'YYYY.MM.DD - YYYY.MM.DD'.
    Start: first date inside the trailing bracket of the Period cell.
    End:   second 8-digit date in the filename (encodes the actual test window).
    """
    start_date = ""
    m = re.search(r"\((\d{4}\.\d{2}\.\d{2})\s*-\s*\d{4}\.\d{2}\.\d{2}\)\s*$", period_raw)
    if m:
        start_date = m.group(1)

    end_date = ""
    fn_m = re.search(r"-(\d{8})-(\d{8})-", filename)
    if fn_m:
        raw_end = fn_m.group(2)
        end_date = f"{raw_end[:4]}.{raw_end[4:6]}.{raw_end[6:8]}"

    if start_date and end_date:
        return f"{start_date} - {end_date}"
    return start_date or end_date or period_raw


def maximal_drawdown_pct(value: str) -> float:
    """Extract the percentage from '464.06 (4.49%)' → 4.49."""
    m = re.search(r"\(([\d.]+)%\)", value)
    return float(m.group(1)) if m else 0.0


def parse_report(file_path: str) -> dict:
    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    soup = BeautifulSoup(content, "html.parser")
    flat = {}  # label -> value collected from the HTML

    for row in soup.find_all("tr"):
        tds = row.find_all("td")
        if not tds:
            continue

        first = tds[0]
        first_colspan = first.get("colspan", "1")
        first_text = get_td_text(first)

        # ── Two-cell header rows: Symbol / Period / Model / Parameters ──────
        # Structure: <td colspan=2>Label</td><td colspan=4>Value</td>
        if first_colspan == "2" and len(tds) == 2:
            second = tds[1]
            second_colspan = second.get("colspan", "1")
            if second_colspan == "4":
                flat[first_text] = get_td_text(second)
            continue

        # ── "Largest" / "Average" rows ───────────────────────────────────────
        # Structure: <td colspan=2 align=right>Largest|Average</td>
        #            <td>profit trade</td><td>value</td>
        #            <td>loss trade</td><td>value</td>
        if first_colspan == "2" and first.get("align") == "right" and len(tds) >= 5:
            context = first_text  # "Largest", "Average", "Maximum", "Maximal"
            if context in ("Largest", "Average", "Maximum", "Maximal"):
                # Use sub-labels (tds[1], tds[3]) as part of the key to avoid collisions
                # e.g. "Average profit trade" vs "Average consecutive wins"
                sub1 = get_td_text(tds[1])
                sub2 = get_td_text(tds[3])
                flat[f"{context} {sub1}"] = get_td_text(tds[2])
                flat[f"{context} {sub2}"] = get_td_text(tds[4])
            continue

        # ── Skip rows whose first cell is colspan=2 but don't match above ───
        if first_colspan == "2":
            continue

        # ── Normal rows: pairs of (label, value) across up to 6 cells ───────
        i = 0
        while i + 1 < len(tds):
            label = get_td_text(tds[i])
            value = get_td_text(tds[i + 1])
            if label:
                flat[label] = value
            i += 2

    # ── Map flat dict to the required columns ────────────────────────────────
    result = {"Filename": Path(file_path).name}

    result["Symbol"] = flat.get("Symbol", "")
    result["Period"] = parse_period(flat.get("Period", ""), Path(file_path).name)
    result["Modelling quality"] = flat.get("Modelling quality", "")
    result["Spread"] = flat.get("Spread", "")
    result["Initial deposit"] = flat.get("Initial deposit", "")
    result["Total net profit"] = flat.get("Total net profit", "")
    result["Gross profit"] = flat.get("Gross profit", "")
    result["Gross loss"] = flat.get("Gross loss", "")
    result["Profit factor"] = flat.get("Profit factor", "")
    result["Expected payoff"] = flat.get("Expected payoff", "")
    result["Absolute drawdown"] = flat.get("Absolute drawdown", "")
    result["Maximal drawdown"] = flat.get("Maximal drawdown", "")
    result["Relative drawdown"] = flat.get("Relative drawdown", "")
    result["Total trades"] = flat.get("Total trades", "")

    short_raw = flat.get("Short positions (won %)", "")
    long_raw = flat.get("Long positions (won %)", "")
    short_count, short_pct = parse_positions(short_raw)
    long_count, long_pct = parse_positions(long_raw)
    result["Short positions"] = short_count
    result["Short positions won %"] = short_pct
    result["Long positions"] = long_count
    result["Long positions won %"] = long_pct

    result["Largest profit trade"] = flat.get("Largest profit trade", "")
    result["Largest loss trade"] = flat.get("Largest loss trade", "")
    result["Average profit trade"] = flat.get("Average profit trade", "")
    result["Average loss trade"] = flat.get("Average loss trade", "")
    result["Parameters"] = flat.get("Parameters", "")

    return result


RED_FILL = PatternFill(fill_type="solid", fgColor="FF0000")


def write_excel(rows: list, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="4472C4")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Write header
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Write data rows
    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            value = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            # No wrap for any cell (Parameters included)
            cell.alignment = Alignment(vertical="top", wrap_text=False)

            # ── Conditional red fill ─────────────────────────────────────────
            # Maximal drawdown: red if the bracketed % > 30
            if col_name == "Maximal drawdown" and value:
                if maximal_drawdown_pct(str(value)) > 30:
                    cell.fill = RED_FILL

            # Total net profit: red if value is negative
            if col_name == "Total net profit" and value:
                try:
                    if float(str(value).replace(",", "")) < 0:
                        cell.fill = RED_FILL
                except ValueError:
                    pass

    # Auto-size columns (capped at 40; Parameters capped at 60)
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        max_len = len(col_name)
        for row_idx in range(2, len(rows) + 2):
            val = ws.cell(row=row_idx, column=col_idx).value or ""
            max_len = max(max_len, min(len(str(val)), 60))
        cap = 60 if col_name == "Parameters" else 40
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, cap)

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"Saved: {output_path}")


def main():
    base_dir = Path(__file__).parent.parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)

    htm_files = sorted(input_dir.glob("*.htm")) + sorted(input_dir.glob("*.html"))
    if not htm_files:
        print(f"No .htm / .html files found in {input_dir}")
        return

    rows = []
    for htm_file in htm_files:
        print(f"Processing: {htm_file.name}")
        try:
            row = parse_report(str(htm_file))
            rows.append(row)
        except Exception as exc:
            print(f"  ERROR: {exc}")

    output_path = str(output_dir / "summary.xlsx")
    write_excel(rows, output_path)
    print(f"\nDone. {len(rows)} report(s) extracted.")


if __name__ == "__main__":
    main()
