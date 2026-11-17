#!/usr/bin/env python
# coding: utf-8
"""Combine Scopus and Web of Science exports into one Excel workbook.

Detects duplicate papers across the two databases by DOI first, then by
title similarity (with year / first-author boosts). Writes a workbook with
four sheets: Combined_Data (unique papers), Scopus_Original (all Scopus),
WoS_Original (all WoS), and Statistics.

Input sheets are auto-detected from their headers, and the reader tolerates
Scopus/WoS quirks such as HTML tables saved with an .xlsx extension or
slightly malformed xlsx files. Missing positional arguments are prompted
for interactively, so the script can be run with no arguments. In an
interactive terminal, the sheets of each input file are listed and one can
be picked by number or name (Enter keeps the auto-detected choice); piped
runs and explicit --scopus-sheet/--wos-sheet flags skip that prompt.
Paths typed at the interactive prompts can be completed with Tab
(readline).

Examples
--------
    # Fully interactive
    python combine_papers.py

    # Fully non-interactive
    python combine_papers.py scopus.xlsx wos.xlsx -o combined.xlsx

    # Custom title threshold, output prompted
    python combine_papers.py scopus.xlsx wos.xlsx --title-threshold 0.9

    # Override auto-detection if the file has a weird layout
    python combine_papers.py scopus.xlsx wos.xlsx --scopus-sheet "Sheet3"
"""

import argparse
import os
import re
import sys
from contextlib import contextmanager
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill


# ---------------------------------------------------------------------------
# Column mappings
# ---------------------------------------------------------------------------

scopus_column_mapping = {
    'Unnamed: 0': 'index_original',
    'Автор (ы)': 'authors',
    'Author full names': 'authors_full',
    'Идентификатор автора(ов)': 'author_ids',
    'Название документа': 'title',
    'Год': 'year',
    'Название источника': 'source_title',
    'Том': 'volume',
    'Выпуск': 'issue',
    'Статья №': 'article_number',
    'Страница начала': 'start_page',
    'Страница окончания': 'end_page',
    'Цитирования': 'citations',
    'DOI': 'doi',
    'Ссылка': 'link',
    'Организации': 'affiliations',
    'Авторы организаций': 'authors_affiliations',
    'Краткое описание': 'abstract',
    'Ключевые слова автора': 'author_keywords',
    'Ключевые слова указателя': 'index_keywords',
    'Адрес для корреспонденции': 'correspondence_address',
    'Редакторы': 'editors',
    'Издатель': 'publisher',
    'ISSN': 'issn',
    'ISBN': 'isbn',
    'CODEN': 'coden',
    'Идентификатор PubMed': 'pubmed_id',
    'Язык оригинального документа': 'language',
    'Сокращенное название источника': 'source_abbreviation',
    'Тип документа': 'document_type',
    'СТАДИЯ ПУБЛИКАЦИИ': 'publication_stage',
    'Open Access (открытый доступ)': 'open_access',
    'Источник': 'source',
    'EID': 'eid',
    'Authors': 'authors',
    'Author full names': 'authors_full',
    'Author(s) ID': 'author_ids',
    'Title': 'title',
    'Year': 'year',
    'Source title': 'source_title',
    'Volume': 'volume',
    'Issue': 'issue',
    'Art. No.': 'article_number',
    'Page start': 'start_page',
    'Page end': 'end_page',
    'Cited by': 'citations',
    'DOI': 'doi',
    'Link': 'link',
    'Affiliations': 'affiliations',
    'Authors with affiliations': 'authors_affiliations',
    'Abstract': 'abstract',
    'Author Keywords': 'author_keywords',
    'Index Keywords': 'index_keywords',
    'Molecular Sequence Numbers': 'mol_sec_n',
    'Chemicals/CAS': 'casno',
    'Tradenames': 'trademarks',
    'Manufacturers': 'manufacturers',
    'Funding Details': 'funding',
    'Funding Texts': 'funding_texts',
    'References': 'cited_references',
    'Correspondence Address': 'correspondence_address',
    'Editors': 'editors',
    'Publisher': 'publisher',
    'Sponsors': 'sponsors',
    'Conference name': 'conference_title',
    'Conference date': 'conference_date',
    'Conference location': 'conference_location',
    'Conference code': 'conference_code',
    'ISSN': 'issn',
    'ISBN': 'isbn',
    'CODEN': 'coden',
    'PubMed ID': 'pubmed_id',
    'Language of Original Document': 'language',
    'Abbreviated Source Title': 'source_abbreviation',
    'Document Type': 'document_type',
    'Publication Stage': 'publication_stage',
    'Open Access': 'open_access',
    'Source': 'source',
}

wos_column_mapping = {
    'Publication Type': 'publication_type',
    'Authors': 'authors',
    'Book Authors': 'book_authors',
    'Book Editors': 'book_editors',
    'Book Group Authors': 'book_group_authors',
    'Author Full Names': 'authors_full',
    'Book Author Full Names': 'book_authors_full',
    'Group Authors': 'group_authors',
    'Article Title': 'title',
    'Source Title': 'source_title',
    'Book Series Title': 'book_series_title',
    'Book Series Subtitle': 'book_series_subtitle',
    'Language': 'language',
    'Document Type': 'document_type',
    'Conference Title': 'conference_title',
    'Conference Date': 'conference_date',
    'Conference Location': 'conference_location',
    'Conference Sponsor': 'conference_sponsor',
    'Conference Host': 'conference_host',
    'Author Keywords': 'author_keywords',
    'Keywords Plus': 'keywords_plus',
    'Abstract': 'abstract',
    'Addresses': 'addresses',
    'Affiliations': 'affiliations',
    'Reprint Addresses': 'reprint_addresses',
    'Email Addresses': 'email_addresses',
    'Researcher Ids': 'researcher_ids',
    'ORCIDs': 'orcids',
    'Funding Orgs': 'funding_orgs',
    'Funding Name Preferred': 'funding_name_preferred',
    'Funding Text': 'funding_text',
    'Cited References': 'cited_references',
    'Cited Reference Count': 'cited_reference_count',
    'Times Cited, WoS Core': 'times_cited_wos',
    'Times Cited, All Databases': 'times_cited_all',
    '180 Day Usage Count': 'usage_180_days',
    'Since 2013 Usage Count': 'usage_since_2013',
    'Publisher': 'publisher',
    'Publisher City': 'publisher_city',
    'Publisher Address': 'publisher_address',
    'ISSN': 'issn',
    'eISSN': 'eissn',
    'ISBN': 'isbn',
    'Journal Abbreviation': 'journal_abbreviation',
    'Journal ISO Abbreviation': 'journal_iso_abbreviation',
    'Publication Date': 'publication_date',
    'Publication Year': 'year',
    'Volume': 'volume',
    'Issue': 'issue',
    'Part Number': 'part_number',
    'Supplement': 'supplement',
    'Special Issue': 'special_issue',
    'Meeting Abstract': 'meeting_abstract',
    'Start Page': 'start_page',
    'End Page': 'end_page',
    'Article Number': 'article_number',
    'DOI': 'doi',
    'DOI Link': 'doi_link',
    'Book DOI': 'book_doi',
    'Early Access Date': 'early_access_date',
    'Number of Pages': 'pages_count',
    'WoS Categories': 'wos_categories',
    'Web of Science Index': 'wos_index',
    'Research Areas': 'research_areas',
    'IDS Number': 'ids_number',
    'Pubmed Id': 'pubmed_id',
    'Open Access Designations': 'open_access',
    'Highly Cited Status': 'highly_cited',
    'Hot Paper Status': 'hot_paper',
    'Date of Export': 'export_date',
    'UT (Unique WOS ID)': 'ut_wos_id',
    'Web of Science Record': 'wos_record',
}

common_columns = [
    'title', 'year', 'authors', 'authors_full', 'doi', 'abstract',
    'author_keywords', 'affiliations', 'source_title', 'volume',
    'issue', 'start_page', 'end_page', 'article_number', 'citations',
    'publisher', 'issn', 'isbn', 'language', 'document_type',
    'open_access', 'pubmed_id', 'source_db',
]


# ---------------------------------------------------------------------------
# Sheet auto-detection markers
# ---------------------------------------------------------------------------
# Distinctive header names used to guess which sheet holds the actual
# export. Case-insensitive. Deliberately excludes ambiguous names that
# appear in both exports (e.g. "Title" vs "Article Title", "Source title"
# vs "Source Title").

SCOPUS_SHEET_MARKERS = {
    'EID',
    'Author(s) ID',
    'Cited by',
    'Идентификатор автора(ов)',
    'Название документа',
    'Название источника',
    'Цитирования',
    'Статья №',
}

WOS_SHEET_MARKERS = {
    'UT (Unique WOS ID)',
    'Article Title',
    'Web of Science Record',
    'Publication Type',
    'WoS Categories',
    'IDS Number',
    'Keywords Plus',
    'Book Authors',
}


# ---------------------------------------------------------------------------
# Matching logic (unchanged)
# ---------------------------------------------------------------------------

def clean_title(title):
    """Clean and normalize title for comparison"""
    if pd.isna(title):
        return ""
    title = str(title).lower()
    title = re.sub(r'[^\w\s]', ' ', title)
    title = ' '.join(title.split())
    stop_words = ['the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at',
                  'to', 'for', 'of', 'with']
    words = title.split()
    words = [w for w in words if w not in stop_words]
    return ' '.join(words)


def title_similarity(title1, title2):
    """Calculate similarity ratio between two titles"""
    if pd.isna(title1) or pd.isna(title2):
        return 0
    clean1 = clean_title(title1)
    clean2 = clean_title(title2)
    return SequenceMatcher(None, clean1, clean2).ratio()


def find_duplicates_by_title(scopus_df, wos_df, similarity_threshold=0.85):
    """Find duplicate papers based on DOI first, then title similarity."""
    scopus_df = scopus_df.copy()
    wos_df = wos_df.copy()

    scopus_df['clean_title'] = scopus_df['title'].apply(clean_title)
    wos_df['clean_title'] = wos_df['title'].apply(clean_title)

    scopus_df['doi_clean'] = scopus_df['doi'].str.lower().str.strip() \
        if 'doi' in scopus_df.columns else None
    wos_df['doi_clean'] = wos_df['doi'].str.lower().str.strip() \
        if 'doi' in wos_df.columns else None

    scopus_dois = set(scopus_df[scopus_df['doi_clean'].notna()]['doi_clean'])
    wos_dois = set(wos_df[wos_df['doi_clean'].notna()]['doi_clean'])
    duplicate_dois = scopus_dois.intersection(wos_dois)

    matched_scopus_indices = set()
    matched_wos_indices = set()
    matches = []

    for doi in duplicate_dois:
        scopus_idx = scopus_df[scopus_df['doi_clean'] == doi].index[0]
        wos_idx = wos_df[wos_df['doi_clean'] == doi].index[0]
        matches.append((scopus_idx, wos_idx, 'doi', 1.0))
        matched_scopus_indices.add(scopus_idx)
        matched_wos_indices.add(wos_idx)

    scopus_unmatched = scopus_df[~scopus_df.index.isin(matched_scopus_indices)]
    wos_unmatched = wos_df[~wos_df.index.isin(matched_wos_indices)]

    title_matches = []
    for scopus_idx, scopus_row in scopus_unmatched.iterrows():
        best_match = None
        best_similarity = 0
        for wos_idx, wos_row in wos_unmatched.iterrows():
            if wos_idx in matched_wos_indices:
                continue
            similarity = title_similarity(scopus_row['title'], wos_row['title'])
            if scopus_row['year'] == wos_row['year']:
                similarity += 0.05
            if pd.notna(scopus_row.get('authors')) and pd.notna(wos_row.get('authors')):
                scopus_first_author = str(scopus_row['authors']).split(';')[0].strip().lower()
                wos_first_author = str(wos_row['authors']).split(';')[0].strip().lower()
                if scopus_first_author == wos_first_author:
                    similarity += 0.05
            if similarity > best_similarity and similarity >= similarity_threshold:
                best_similarity = similarity
                best_match = wos_idx
        if best_match is not None:
            title_matches.append((scopus_idx, best_match, 'title', best_similarity))
            matched_scopus_indices.add(scopus_idx)
            matched_wos_indices.add(best_match)

    all_matches = matches + title_matches
    return all_matches, matched_scopus_indices, matched_wos_indices


def load_and_translate_scopus(df_scopus):
    df_clean = df_scopus.rename(columns=scopus_column_mapping)
    df_clean.columns = df_clean.columns.str.strip()
    return df_clean


def load_and_translate_wos(df_wos):
    df_clean = df_wos.rename(columns=wos_column_mapping)
    df_clean.columns = df_clean.columns.str.strip()
    return df_clean


def create_combined_dataframe(scopus_df, wos_df, title_threshold=0.85):
    scopus_clean_all = load_and_translate_scopus(scopus_df)
    wos_clean_all = load_and_translate_wos(wos_df)

    matches, matched_scopus, matched_wos = find_duplicates_by_title(
        scopus_clean_all.copy(), wos_clean_all.copy(), title_threshold
    )

    combined_papers = []
    for scopus_idx, wos_idx, match_type, similarity in matches:
        scopus_paper = scopus_clean_all.loc[scopus_idx]
        wos_paper = wos_clean_all.loc[wos_idx]

        combined = {'source_db': 'both', 'match_type': match_type,
                    'match_similarity': similarity}

        for col in common_columns:
            if col in ['source_db', 'match_type', 'match_similarity']:
                continue
            scopus_val = scopus_paper.get(col, None)
            wos_val = wos_paper.get(col, None)
            if pd.notna(scopus_val) and pd.notna(wos_val):
                if col in ['abstract', 'author_keywords', 'affiliations']:
                    combined[col] = f"{scopus_val}; {wos_val}"
                else:
                    combined[col] = scopus_val
            elif pd.notna(scopus_val):
                combined[col] = scopus_val
            elif pd.notna(wos_val):
                combined[col] = wos_val
            else:
                combined[col] = None

        if 'eid' in scopus_paper:
            combined['scopus_eid'] = scopus_paper['eid']
        if 'ut_wos_id' in wos_paper:
            combined['wos_ut_id'] = wos_paper['ut_wos_id']

        combined_papers.append(combined)

    scopus_clean_all['source_db'] = 'scopus'
    wos_clean_all['source_db'] = 'wos'

    scopus_unique = scopus_clean_all[~scopus_clean_all.index.isin(matched_scopus)].copy()
    wos_unique = wos_clean_all[~wos_clean_all.index.isin(matched_wos)].copy()

    scopus_common = scopus_unique[
        [c for c in common_columns if c in scopus_unique.columns]
    ].copy()
    wos_common = wos_unique[
        [c for c in common_columns if c in wos_unique.columns]
    ].copy()

    combined_df = pd.concat(
        [scopus_common, wos_common, pd.DataFrame(combined_papers)],
        ignore_index=True, sort=False,
    )

    for col in ('scopus_eid', 'wos_ut_id', 'match_type', 'match_similarity'):
        if col not in combined_df.columns:
            combined_df[col] = None

    return combined_df, scopus_clean_all, wos_clean_all


# ---------------------------------------------------------------------------
# Excel export & formatting (unchanged)
# ---------------------------------------------------------------------------

def export_to_excel_with_sheets(combined_df, scopus_original, wos_original,
                                filename='combined_papers.xlsx'):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='Combined_Data', index=False)
        scopus_original.to_excel(writer, sheet_name='Scopus_Original', index=False)
        wos_original.to_excel(writer, sheet_name='WoS_Original', index=False)

        stats_data = {
            'Metric': [
                'Total Papers in Combined Sheet (Unique)',
                'Papers from Scopus Only',
                'Papers from WoS Only',
                'Papers in Both Databases (Total)',
                '- Matched by DOI',
                '- Matched by Title',
                'Total Scopus Papers (All - Original Sheet)',
                'Total WoS Papers (All - Original Sheet)',
                'Duplicate Papers Found (in both)',
            ],
            'Count': [
                len(combined_df),
                len(combined_df[combined_df['source_db'] == 'scopus']),
                len(combined_df[combined_df['source_db'] == 'wos']),
                len(combined_df[combined_df['source_db'] == 'both']),
                len(combined_df[(combined_df['source_db'] == 'both')
                                & (combined_df['match_type'] == 'doi')]),
                len(combined_df[(combined_df['source_db'] == 'both')
                                & (combined_df['match_type'] == 'title')]),
                len(scopus_original),
                len(wos_original),
                len(combined_df[combined_df['source_db'] == 'both']),
            ],
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='Statistics', index=False)

        format_excel_sheets(writer, filename)

    print(f"\n✅ Successfully exported to {filename}")
    print("\n📊 Summary:")
    print(f"  • Combined sheet: {len(combined_df)} papers (unique papers only)")
    print(f"    - Scopus only: {len(combined_df[combined_df['source_db'] == 'scopus'])}")
    print(f"    - WoS only: {len(combined_df[combined_df['source_db'] == 'wos'])}")
    print(f"    - Both databases: {len(combined_df[combined_df['source_db'] == 'both'])}")
    print(f"      * Matched by DOI: {len(combined_df[(combined_df['source_db'] == 'both') & (combined_df['match_type'] == 'doi')])}")
    print(f"      * Matched by Title: {len(combined_df[(combined_df['source_db'] == 'both') & (combined_df['match_type'] == 'title')])}")
    print(f"  • Scopus original sheet: {len(scopus_original)} papers (ALL papers, including duplicates)")
    print(f"  • WoS original sheet: {len(wos_original)} papers (ALL papers, including duplicates)")


def format_excel_sheets(writer, filename):
    workbook = writer.book
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        worksheet.freeze_panes = 'A2'

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092",
                                    fill_type="solid")
            cell.alignment = Alignment(horizontal="center")


# ---------------------------------------------------------------------------
# I/O helpers (type sniffing, sheet auto-detection, parser fallbacks)
# ---------------------------------------------------------------------------

def _sniff_file_kind(path: Path) -> str:
    """Return 'zip' (real xlsx), 'html', 'xml', or 'text'.

    Real .xlsx files are ZIP archives (magic bytes ``PK``). Scopus/WoS
    sometimes hand back an HTML table with an .xlsx extension, which is
    what makes openpyxl die with "substring not found".
    """
    try:
        with path.open("rb") as f:
            head = f.read(8192)
    except OSError:
        return "text"

    if head[:2] == b"PK":
        return "zip"

    lowered = head.lstrip().lower()
    if (lowered.startswith(b"<")
            or b"<html" in lowered
            or b"<table" in lowered
            or b"<table " in lowered):
        return "html"
    if lowered.startswith(b"<?xml"):
        return "xml"
    return "text"


def _detect_sheet(path: Path, markers) -> str:
    """Return the sheet whose headers best match ``markers``.

    Per-sheet read errors are tolerated; if every sheet fails, the
    exception is re-raised so ``read_table`` can try an alternative parser.
    """
    markers_lower = {str(m).strip().lower() for m in markers}

    xls = pd.ExcelFile(path)
    try:
        sheet_names = list(xls.sheet_names)
        if not sheet_names:
            return 0
        if len(sheet_names) == 1:
            return sheet_names[0]

        best_sheet = sheet_names[0]
        best_score = -1
        last_error = None

        for sheet in sheet_names:
            try:
                header = pd.read_excel(xls, sheet_name=sheet, nrows=0)
            except Exception as exc:
                last_error = exc
                print(
                    f"⚠️  Cannot read headers of sheet {sheet!r} in "
                    f"{path.name}: {exc}",
                    file=sys.stderr,
                )
                continue
            cols = {str(c).strip().lower() for c in header.columns}
            score = len(cols & markers_lower)
            if score > best_score:
                best_sheet, best_score = sheet, score

        if best_score <= 0:
            if last_error is not None and best_score < 0:
                raise last_error
            print(
                f"⚠️  No known header markers in {path.name}; using "
                f"first sheet {best_sheet!r}.",
                file=sys.stderr,
            )
        else:
            print(
                f"🔎 {path.name}: using sheet {best_sheet!r} "
                f"(matched {best_score} marker(s))",
                file=sys.stderr,
            )
        return best_sheet
    finally:
        xls.close()


def _read_excel_via_openpyxl(path: Path, sheet_name=None) -> pd.DataFrame:
    """Read a workbook with openpyxl directly, bypassing pandas' parser.

    pandas routes every sheet through its CSV-like TextParser, whose BOM
    handling assumes a quote character follows a leading U+FEFF in the
    first cell and dies with ``ValueError: substring not found``
    otherwise. Exports written from a BOM-prefixed CSV hit this on every
    engine (openpyxl, calamine). openpyxl itself parses such files fine,
    so extract the rows by hand; header cells are cleaned of the BOM and
    stray wrapping quotes so column mapping still matches.
    """
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        wb = load_workbook(path, data_only=True)
    try:
        if isinstance(sheet_name, str):
            ws = wb[sheet_name]
        else:
            idx = sheet_name if isinstance(sheet_name, int) else 0
            ws = wb[wb.sheetnames[idx]]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not rows:
        return pd.DataFrame()
    header = [
        "" if h is None else str(h).strip().lstrip("\ufeff").strip('"')
        for h in rows[0]
    ]
    return pd.DataFrame(rows[1:], columns=header)


def read_table(path: Path, sheet_name=None, encoding: str = "utf-8",
               markers=None) -> pd.DataFrame:
    """Read an Excel or CSV file into a DataFrame.

    Falls back to ``calamine`` for broken .xlsx files and to
    ``pd.read_html`` for HTML tables disguised as .xlsx (a common quirk
    of Scopus exports).
    """
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()

    # ---- CSV ------------------------------------------------------------
    if suffix == ".csv":
        return pd.read_csv(path, encoding=encoding)

    if suffix not in (".xlsx", ".xls"):
        raise ValueError(
            f"Unsupported input extension: {suffix!r} "
            f"(use .xlsx/.xls/.csv)"
        )

    kind = _sniff_file_kind(path)

    # ---- HTML disguised as Excel (Scopus does this) --------------------
    if kind == "html":
        print(
            f"ℹ️  {path.name}: file is HTML with an Excel extension — "
            f"reading with pd.read_html()",
            file=sys.stderr,
        )
        try:
            tables = pd.read_html(str(path))
        except ImportError as exc:
            raise ValueError(
                f"{path.name} is an HTML table, but parsing it requires "
                f"lxml or html5lib (pip install lxml)."
            ) from exc
        except Exception as exc:
            raise ValueError(
                f"{path.name} looks like HTML but pd.read_html() failed: "
                f"{exc}"
            ) from exc
        if not tables:
            raise ValueError(f"No tables found in HTML file {path}")
        # The largest table is virtually always the actual export.
        return max(tables, key=len)

    # ---- Normal .xlsx / .xls -------------------------------------------
    # First pass: auto-detect sheet + openpyxl (the existing behaviour).
    try:
        if sheet_name is None:
            sheet_name = _detect_sheet(path, markers) if markers else 0
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as first_error:
        # Second pass: try the calamine engine. Handles many malformed
        # xlsx files that trip up openpyxl with "substring not found".
        try:
            target_sheet = sheet_name if sheet_name is not None else 0
            df = pd.read_excel(path, sheet_name=target_sheet,
                               engine="calamine")
            print(
                f"ℹ️  {path.name}: read via the 'calamine' engine after "
                f"openpyxl failed ({first_error})",
                file=sys.stderr,
            )
            return df
        except Exception as second_error:
            # Third pass: maybe it really is HTML that just didn't sniff
            # as such (e.g. served inside a ZIP). Give read_html a shot.
            try:
                tables = pd.read_html(str(path))
                if tables:
                    print(
                        f"ℹ️  {path.name}: read as HTML table after Excel "
                        f"parsers failed",
                        file=sys.stderr,
                    )
                    return max(tables, key=len)
            except Exception:
                pass

            # Fourth pass: read the workbook with openpyxl directly.
            # Skips pandas' TextParser, which crashes on files whose
            # first header cell carries a BOM ("substring not found").
            direct_error = None
            try:
                df = _read_excel_via_openpyxl(path, sheet_name)
                print(
                    f"ℹ️  {path.name}: read via openpyxl directly after "
                    f"pandas' Excel parsers failed ({first_error})",
                    file=sys.stderr,
                )
                return df
            except Exception as exc:
                direct_error = exc

            raise ValueError(
                f"Could not read {path.name} as Excel. "
                f"openpyxl said: {first_error!r}. "
                f"calamine said: {second_error!r}. "
                f"direct openpyxl said: {direct_error!r}. "
                f"If this is a Scopus/WoS export, it is often an HTML "
                f"table saved with an .xlsx extension — open it in a "
                f"browser to confirm, and if so either save it as .html "
                f"or install lxml (pip install lxml)."
            ) from first_error


def process_papers(scopus_file, wos_file, output_file='combined_papers.xlsx',
                   title_threshold=0.85, scopus_sheet=None, wos_sheet=None,
                   encoding="utf-8"):
    scopus_df = read_table(Path(scopus_file), sheet_name=scopus_sheet,
                           encoding=encoding, markers=SCOPUS_SHEET_MARKERS)
    wos_df = read_table(Path(wos_file), sheet_name=wos_sheet,
                        encoding=encoding, markers=WOS_SHEET_MARKERS)

    print(f"📚 Loaded {len(scopus_df)} papers from Scopus")
    print(f"📚 Loaded {len(wos_df)} papers from WoS")
    print(f"🔍 Using title similarity threshold: {title_threshold}")

    combined_df, scopus_original, wos_original = create_combined_dataframe(
        scopus_df, wos_df, title_threshold
    )

    export_to_excel_with_sheets(combined_df, scopus_original, wos_original,
                                str(output_file))

    return combined_df, scopus_original, wos_original


# ---------------------------------------------------------------------------
# argparse + interactive prompting
# ---------------------------------------------------------------------------

def _threshold_type(value: str) -> float:
    try:
        f = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"invalid float: {value!r}") from exc
    if not 0.0 <= f <= 1.0:
        raise argparse.ArgumentTypeError(
            f"threshold must be between 0 and 1, got {f}"
        )
    return f


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Combine Scopus and Web of Science exports into a single Excel "
            "workbook with deduplication by DOI and title similarity. "
            "Input sheets are auto-detected from their headers; missing "
            "positional arguments are prompted for interactively "
            "(paths can be completed with Tab)."
        )
    )
    parser.add_argument(
        "scopus_file", type=Path, nargs="?", default=None,
        help="Path to the Scopus export (.xlsx/.xls/.csv). Prompted if omitted.",
    )
    parser.add_argument(
        "wos_file", type=Path, nargs="?", default=None,
        help="Path to the WoS export (.xlsx/.xls/.csv). Prompted if omitted.",
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output .xlsx file. Prompted if omitted (defaults to "
             "<scopus_stem>_combined.xlsx).",
    )
    parser.add_argument(
        "-t", "--title-threshold", type=_threshold_type, default=0.85,
        help="Title similarity threshold in [0, 1] (default: 0.85).",
    )
    parser.add_argument(
        "--scopus-sheet", default=None,
        help="Override Scopus sheet auto-detection with this sheet name "
             "(its selection is prompted for if omitted in interactive "
             "runs).",
    )
    parser.add_argument(
        "--wos-sheet", default=None,
        help="Override WoS sheet auto-detection with this sheet name "
             "(its selection is prompted for if omitted in interactive "
             "runs).",
    )
    parser.add_argument(
        "-e", "--encoding", default="utf-8",
        help="Encoding for CSV inputs (default: utf-8).",
    )
    return parser.parse_args(argv)


def _prompt(prompt: str) -> str:
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


def _prompt_existing_file(label: str) -> Path:
    with _path_completion():
        while True:
            raw = _prompt(f"{label} file: ")
            if not raw:
                print("  A path is required.", file=sys.stderr)
                continue
            path = Path(raw).expanduser()
            if path.is_file():
                if path.suffix.lower() in (".xlsx", ".xls", ".csv"):
                    return path
                print(
                    "  Expected .xlsx, .xls or .csv.", file=sys.stderr,
                )
                continue
            print(f"  File not found: {path}", file=sys.stderr)


def _prompt_output(default: Path) -> Path:
    with _path_completion():
        while True:
            raw = _prompt(f"Output Excel file [{default}]: ")
            path = Path(raw).expanduser() if raw else default
            if path.suffix.lower() in (".xlsx", ".xls"):
                return path
            print("  Output must end in .xlsx or .xls.", file=sys.stderr)


def _list_sheets(path: Path) -> list:
    """Return the sheet names of an Excel file, or [] if unreadable.

    CSV files and HTML tables disguised as .xlsx have no sheets to list;
    the empty list tells the caller to fall back to auto-detection.
    """
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return []
    try:
        return list(xls.sheet_names)
    finally:
        xls.close()


def _prompt_sheet(path: Path, label: str, markers) -> str | None:
    """Ask the user which sheet of ``path`` to read.

    Returns the chosen sheet name, or ``None`` to keep the existing
    auto-detection. Files with a single sheet (or that cannot be opened
    as Excel workbooks) skip the prompt.
    """
    sheets = _list_sheets(path)
    if not sheets:
        return None
    if len(sheets) == 1:
        print(f"  {path.name}: single sheet {sheets[0]!r}")
        return sheets[0]

    default_sheet = sheets[0]
    try:
        default_sheet = _detect_sheet(path, markers)
    except Exception:
        pass

    print(f"\n📋 Sheets in {path.name}:")
    for i, name in enumerate(sheets, 1):
        mark = "   <- default (auto-detected)" if name == default_sheet else ""
        print(f"  {i}. {name}{mark}")

    while True:
        raw = _prompt(
            f"{label} sheet [1-{len(sheets)} or name, "
            f"Enter = auto-detect]: "
        )
        if not raw:
            print(f"  Using auto-detected sheet {default_sheet!r}")
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


def _resolve_sheet(path: Path, label: str, markers, override):
    """Pick the sheet for one input file.

    An explicit ``--*-sheet`` override wins; otherwise the user is asked
    in interactive terminals (Excel files only). Piped runs and CSV
    inputs keep silent auto-detection.
    """
    if override is not None:
        return override
    if path.suffix.lower() not in (".xlsx", ".xls"):
        return None
    if not sys.stdin.isatty():
        return None
    return _prompt_sheet(path, label, markers)


def main(argv=None):
    args = parse_args(argv)

    scopus_file = args.scopus_file or _prompt_existing_file("Scopus")
    wos_file = args.wos_file or _prompt_existing_file("WoS")

    scopus_sheet = _resolve_sheet(
        scopus_file, "Scopus", SCOPUS_SHEET_MARKERS, args.scopus_sheet
    )
    wos_sheet = _resolve_sheet(
        wos_file, "WoS", WOS_SHEET_MARKERS, args.wos_sheet
    )

    if args.output is not None:
        output_file = args.output
        if output_file.suffix.lower() not in (".xlsx", ".xls"):
            print(
                f"Error: output must end in .xlsx or .xls, got "
                f"{output_file.suffix!r}",
                file=sys.stderr,
            )
            return 1
    else:
        default_out = scopus_file.with_name(scopus_file.stem + "_combined.xlsx")
        output_file = _prompt_output(default_out)

    try:
        combined, scopus_orig, wos_orig = process_papers(
            scopus_file,
            wos_file,
            output_file=output_file,
            title_threshold=args.title_threshold,
            scopus_sheet=scopus_sheet,
            wos_sheet=wos_sheet,
            encoding=args.encoding,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("\n📄 Combined Data Preview (common columns):")
    print(combined[['title', 'year', 'source_db', 'match_type']].head(10))

    print(f"\n📊 Scopus original sheet has {len(scopus_orig.columns)} columns (all translated)")
    print(f"📊 Scopus original sheet contains ALL {len(scopus_orig)} original papers")
    print(f"📊 WoS original sheet has {len(wos_orig.columns)} columns (all translated)")
    print(f"📊 WoS original sheet contains ALL {len(wos_orig)} original papers")

    title_matches = combined[combined['match_type'] == 'title']
    if len(title_matches) > 0:
        print(f"\n📌 Found {len(title_matches)} papers matched by title "
              f"(similarity ≥ {args.title_threshold})")
        print("\nSample title matches:")
        for _, row in title_matches.head(5).iterrows():
            print(f"  • {row['title']} ({row['year']}) - "
                  f"similarity: {row['match_similarity']:.2%}")

    return 0


if __name__ == "__main__":
    sys.exit(main())