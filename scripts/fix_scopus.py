#!/usr/bin/env python
# coding: utf-8
"""Repair a malformed Scopus export and write the parsed result out.

The exported data has broken quoting/field boundaries, so we read it
line-by-line (each line = one record) and re-split it manually instead of
letting pandas' CSV parser silently re-interpret the fields.

The input may be the raw CSV or an Excel workbook holding the same broken
records (one record per row, possibly shattered across several cells —
the cells of each row are joined back into one line, as in the original
notebook). In an interactive terminal, the sheets of an Excel input are
listed and one can be picked by number or name (Enter keeps the first
sheet); piped runs and --input-sheet skip the prompt. Paths typed at the
interactive prompts can be completed with Tab (readline).

Examples
--------
    # Fully non-interactive
    python fix_scopus.py scopus_export_2.csv scopus_export_fixed.xlsx

    # Excel input (sheet prompted in interactive runs)
    python fix_scopus.py scopus_export_2.xlsx fixed.xlsx

    # CSV in -> CSV out, custom encoding
    python fix_scopus.py scopus_export_2.csv fixed.csv --encoding cp1251

    # Interactive (prompts for missing paths)
    python fix_scopus.py
    python fix_scopus.py scopus_export_2.csv
"""

import argparse
import os
import sys
from contextlib import contextmanager
from pathlib import Path

import pandas as pd


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Repair a malformed Scopus export (raw CSV or an Excel "
            "workbook holding the broken records) and write the cleaned "
            "table to Excel or CSV, based on the output extension. "
            "Missing positional arguments are prompted for interactively "
            "(paths can be completed with Tab); "
            "Excel inputs get an interactive sheet menu in a terminal."
        )
    )
    parser.add_argument(
        "input_file",
        type=Path,
        nargs="?",
        default=None,
        help="Path to the input CSV or Excel file (.csv/.xlsx/.xls). "
             "Prompted for if omitted.",
    )
    parser.add_argument(
        "output_file",
        type=Path,
        nargs="?",
        default=None,
        help="Path to the output file (.xlsx or .csv). Prompted for if omitted.",
    )
    parser.add_argument(
        "-e", "--encoding",
        default="utf-8",
        help="Encoding of the input CSV (default: utf-8).",
    )
    parser.add_argument(
        "-s", "--sheet-name",
        default="scopus_export",
        help="Sheet name for Excel output (default: scopus_export).",
    )
    parser.add_argument(
        "--input-sheet",
        default=None,
        help="Sheet to read when the input is an Excel workbook "
             "(prompted for if omitted in interactive runs; "
             "default: first sheet).",
    )
    return parser.parse_args(argv)


def _prompt(prompt: str) -> str:
    """`input()` wrapper that exits cleanly on Ctrl+C / EOF."""
    try:
        return input(prompt).strip().replace('"', '')
    except (KeyboardInterrupt, EOFError):
        print("\nAborted.", file=sys.stderr)
        raise SystemExit(1)


# ---------------------------------------------------------------------------
# Readline Tab-completion for path prompts
# ---------------------------------------------------------------------------

try:
    import readline  # GNU readline, or libedit (e.g. Homebrew/uv Python)
except ImportError:
    try:
        import readline  # pyreadline3 provides this shim on Windows
    except ImportError:
        readline = None

# libedit (used by some Python builds) needs its own key-binding syntax;
# GNU readline and pyreadline3 both understand the classic one.
_IS_LIBEDIT = readline is not None and "libedit" in (
    getattr(readline, "__doc__", "") or ""
)
_QUOTE = '"'


class _PathCompleter:
    """Readline completer that finishes file and directory paths.

    ``text`` is the word under the cursor; word delimiters are set to
    space/tab/newline, so slashes stay inside the word and a whole path
    completes as one unit (this also makes the returned matches line up
    across GNU readline, libedit and pyreadline3). Completed directories
    get a trailing separator so the next Tab descends into them.
    Dotfiles are only offered when the typed prefix starts with a dot.
    A leading double quote is honoured and closed on finished file names
    (type paths containing spaces inside quotes).
    """

    def __init__(self):
        self._matches = []

    def complete(self, text, state):
        if state == 0:
            self._matches = self._matches_for(text)
        try:
            return self._matches[state]
        except IndexError:
            return None

    @staticmethod
    def _matches_for(word: str) -> list:
        quote = ""
        if word.startswith(_QUOTE):
            quote = _QUOTE
            word = word[1:]
        fragment = os.path.expanduser(word)
        if not fragment:
            fragment = "." + os.sep
        dirname = os.path.dirname(fragment)
        prefix = os.path.basename(fragment)
        try:
            names = os.listdir(dirname or ".")
        except OSError:
            return []
        if not prefix.startswith("."):
            names = [n for n in names if not n.startswith(".")]
        matches = []
        for name in sorted(n for n in names if n.startswith(prefix)):
            full = os.path.join(dirname, name) if dirname else name
            if os.path.isdir(full):
                full += os.sep
            elif quote:
                full += quote  # close the quote on completed file names
            matches.append(quote + full)
        return matches


@contextmanager
def _path_completion():
    """Turn on readline Tab-completion of paths inside the block.

    A no-op without a usable readline module (Windows without
    ``pyreadline3``) or when stdin is not a TTY, so piped runs keep
    their exact current behaviour.
    """
    if readline is None or not sys.stdin.isatty():
        yield
        return
    old_completer = readline.get_completer()
    old_delims = readline.get_completer_delims()
    # The default delims split words at '/', '-' and quotes, which would
    # chop every path into unusable fragments.
    readline.set_completer_delims(" \t\n")
    if _IS_LIBEDIT:
        # libedit silently ignores the classic "tab: complete" syntax.
        readline.parse_and_bind("bind ^I rl_complete")
    else:
        readline.parse_and_bind("tab: complete")
    try:
        # Don't append a space after a unique match; the completer already
        # adds a separator to completed directories itself.
        readline.set_completion_append_character("")
    except AttributeError:
        pass  # not available on libedit / pyreadline3
    readline.set_completer(_PathCompleter().complete)
    try:
        yield
    finally:
        readline.set_completer(old_completer)
        readline.set_completer_delims(old_delims)


def prompt_input_file(current: Path | None) -> Path:
    if current is not None:
        return current
    with _path_completion():
        while True:
            raw = _prompt("Input CSV or Excel file: ")
            if not raw:
                print("  A path is required.", file=sys.stderr)
                continue
            path = Path(raw).expanduser()
            if not path.is_file():
                print(f"  File not found: {path}", file=sys.stderr)
                continue
            if path.suffix.lower() in (".csv", ".xlsx", ".xls"):
                return path
            print("  Expected .csv, .xlsx or .xls.", file=sys.stderr)


def prompt_output_file(current: Path | None, input_file: Path) -> Path:
    if current is not None:
        return current
    # Offer a sensible default derived from the input name.
    default = input_file.with_name(input_file.stem + "_fixed.xlsx")
    with _path_completion():
        while True:
            raw = _prompt(f"Output file [{default}]: ")
            if not raw:
                return default
            path = Path(raw).expanduser()
            if path.suffix.lower() in (".csv", ".xlsx", ".xls"):
                return path
            print(
                "  Output must end in .csv, .xlsx, or .xls.",
                file=sys.stderr,
            )


def _list_sheets(path: Path) -> list:
    """Return the sheet names of an Excel file, or [] if unreadable."""
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return []
    try:
        return list(xls.sheet_names)
    finally:
        xls.close()


def prompt_input_sheet(path: Path) -> str | None:
    """Ask which sheet of an Excel input to read.

    Single-sheet workbooks are used without prompting; multi-sheet ones
    get a numbered menu (Enter keeps the first sheet). Returns the chosen
    sheet name, or ``None`` for the default first sheet.
    """
    sheets = _list_sheets(path)
    if not sheets:
        return None
    if len(sheets) == 1:
        print(f"  {path.name}: single sheet {sheets[0]!r}")
        return sheets[0]

    print(f"\n📋 Sheets in {path.name}:")
    for i, name in enumerate(sheets, 1):
        mark = "   <- default" if i == 1 else ""
        print(f"  {i}. {name}{mark}")

    while True:
        raw = _prompt(
            f"Input sheet [1-{len(sheets)} or name, "
            f"Enter = {sheets[0]!r}]: "
        )
        if not raw:
            print(f"  Using sheet {sheets[0]!r}")
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(sheets):
            chosen = sheets[int(raw) - 1]
        elif raw in sheets:
            chosen = raw
        else:
            lowered = raw.lower()
            hits = [s for s in sheets if s.strip().lower() == lowered]
            chosen = hits[0] if len(hits) == 1 else None
        if chosen is not None:
            print(f"  Using sheet {chosen!r}")
            return chosen
        print(
            f"  No sheet named {raw!r}; enter a number from the menu "
            f"or a sheet name.",
            file=sys.stderr,
        )


def load_lines(path: Path, encoding: str = "utf-8"):
    """Read the CSV as raw text lines.

    Reading via pandas here would already split on commas and destroy the
    quoting pattern the downstream logic relies on.
    """
    if not path.is_file():
        raise FileNotFoundError(f"Input file not found: {path}")
    with path.open("r", encoding=encoding, newline="") as f:
        return [line.rstrip("\r\n") for line in f if line.strip()]


def _read_sheet_raw(path: Path, sheet_name) -> pd.DataFrame:
    """Read one worksheet as a frame of raw cells (no header row).

    pandas' parser crashes with ``ValueError: substring not found`` on
    workbooks whose first cell carries a UTF-8 BOM; if that happens, fall
    back to reading the sheet with openpyxl directly.
    """
    if sheet_name is None:
        sheet_name = 0  # pandas would read *all* sheets on None
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception as exc:
        pandas_error = exc

    try:
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            if isinstance(sheet_name, str):
                ws = wb[sheet_name]
            else:
                idx = sheet_name if isinstance(sheet_name, int) else 0
                ws = wb[wb.sheetnames[idx]]
            data = [list(row) for row in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    except Exception as exc:
        raise ValueError(
            f"Could not read {path.name} (sheet {sheet_name!r}): "
            f"pandas said: {pandas_error!r}; openpyxl said: {exc!r}"
        ) from pandas_error

    print(
        f"ℹ️  {path.name}: read via openpyxl directly after pandas' "
        f"Excel parser failed ({pandas_error})",
        file=sys.stderr,
    )
    return pd.DataFrame(data)


def load_lines_from_excel(path: Path, sheet_name=None):
    """Read an Excel export as raw record lines (one line per row).

    The broken export often reaches us as a workbook in which each CSV
    record was shattered across many cells (split at ``;`` boundaries).
    The cells of every row are joined back into one line. The space after
    a boundary semicolon was consumed along with it whenever the next
    fragment does not start with one, so the glue is chosen per boundary —
    same trick as the original notebook, but without losing separators.
    """
    df = _read_sheet_raw(path, sheet_name)
    if df.empty:
        return []
    df = df.fillna("")
    df = df.astype(str)

    lines = []
    for row in df.itertuples(index=False, name=None):
        # Drop trailing empty padding cells so they don't leave ';'-runs.
        cells = list(row)
        while cells and cells[-1] == "":
            cells.pop()
        if not cells:
            continue
        line = cells[0]
        for cell in cells[1:]:
            line += ";" if cell.startswith(" ") else "; "
            line += cell
        lines.append(line)
    return lines


def fix_scopus_export(
    input_file: Path,
    output_file: Path,
    encoding: str = "utf-8",
    sheet_name: str = "scopus_export",
    input_sheet=None,
) -> pd.DataFrame:
    if input_file.suffix.lower() == ".csv":
        lines = load_lines(input_file, encoding)
    else:
        lines = load_lines_from_excel(input_file, input_sheet)
    if not lines:
        raise ValueError(f"Input file is empty: {input_file}")

    # A leading UTF-8 BOM would glue itself to the first header and later
    # crash pandas' Excel reader ("substring not found").
    lines[0] = lines[0].lstrip("\ufeff")

    df = pd.DataFrame({0: lines})

    # Original transformation (kept as-is)
    df_fixed = df.apply(" ".join, axis=1)

    df_split = df_fixed.str.split('","', expand=True)
    split_data = df_split[0].str.split(',"', expand=True)

    # First row acts as the header for both halves
    split_data.columns = split_data.iloc[0]
    split_data = split_data[1:]

    df_split.columns = df_split.iloc[0]
    df_split = df_split.iloc[:, 1:]

    result = split_data.join(df_split)

    # The manual split keeps the outer quotes of the first and last
    # fields (the BOM may also lead the first header). Strip them so the
    # output stays readable by pandas and the headers still match the
    # Scopus column names.
    result.columns = [
        str(c).strip().lstrip("\ufeff").strip('"') for c in result.columns
    ]
    first_col, last_col = result.columns[0], result.columns[-1]
    result[first_col] = result[first_col].map(
        lambda v: v.strip().lstrip("\ufeff").lstrip('"')
        if isinstance(v, str) else v
    )
    result[last_col] = result[last_col].map(
        lambda v: v.strip().rstrip('"') if isinstance(v, str) else v
    )

    # Write output
    if output_file.suffix.lower() == ".csv":
        result.to_csv(output_file, index=False)
    else:
        result.to_excel(output_file, sheet_name=sheet_name, index=False)

    print(
        f"Wrote {len(result)} rows x {len(result.columns)} columns to "
        f"{output_file}",
        file=sys.stderr,
    )
    return result


def main(argv=None):
    args = parse_args(argv)

    # Fill in anything missing interactively.
    input_file = prompt_input_file(args.input_file)

    input_sheet = args.input_sheet
    if input_sheet is None and input_file.suffix.lower() in (".xlsx", ".xls"):
        if sys.stdin.isatty():
            input_sheet = prompt_input_sheet(input_file)
        else:
            sheets = _list_sheets(input_file)
            if len(sheets) > 1:
                print(
                    f"ℹ️  {input_file.name}: multiple sheets, using first "
                    f"{sheets[0]!r} (pass --input-sheet to override)",
                    file=sys.stderr,
                )

    output_file = prompt_output_file(args.output_file, input_file)

    try:
        fix_scopus_export(
            input_file,
            output_file,
            encoding=args.encoding,
            sheet_name=args.sheet_name,
            input_sheet=input_sheet,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())