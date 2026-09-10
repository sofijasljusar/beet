"""
ESAIC 2025 Congress Abstract Extractor
--------------------------------------
Extracts scientific conference presentation metadata and abstracts from the 
European Society of Anaesthesiology and Intensive Care (ESAIC) 2025 Abstract Book
and structures them into the Monocl Excel data schema.

Monocl Target Schema:
  1. Name (incl. titles)
  2. Affiliation/Organisation and location (multiple joined with ' ___ ')
  3. Role ('Poster presenter' for underlined author, 'Abstract author' for co-authors)
  4. Session Name (e.g. 'Best Abstract Prize Competition (BAPC)', 'Perioperative Care')
  5. Session Description (Presentation ID, e.g. 'BAPC-01', '10AP01-2')
  6. Presentation Title
  7. Presentation Abstract (Cleaned text with de-hyphenation)
  8. Abstract URL (Static source link)
"""

import os
import sys
import time
import re
import argparse
import logging
from typing import List, Dict, Any, Tuple, Optional
from tqdm import tqdm
import pymupdf
import openpyxl

logger = logging.getLogger("monocl_extractor")

DEFAULT_ABSTRACT_URL = "https://esaic.org/wp-content/uploads/2025/07/ESAIC2025_Abstracts-1.pdf"

# Monocl Target Excel Schema Headers
EXCEL_HEADERS = [
    'Name (incl. titles)',
    'Affiliation/Organisation and location',
    'Role',
    'Session Name',
    'Session Description',
    'Presentation Title',
    'Presentation Abstract',
    'Abstract URL'
]

# Consortium / study group keywords that should not be treated as individual person authors
NON_PERSON_KEYWORDS = [
    "group", "investigators", "consortium", "study group", "collaborators",
    "network", "committee", "taskforce"
]

# Regex pattern matching ESAIC presentation IDs across all 811 pages (e.g., 'BAPC-01', '10AP01-2')
PRESENTATION_ID_PATTERN = re.compile(r'^(BAPC-\d+|\d+AP\d+-\d+)\s*(.*)', re.DOTALL)

# Regex pattern matching standalone visual artifact tags (e.g., 'Fig 1.', 'Table 2.', 'Figure 1.')
STANDALONE_FIG_PATTERN = re.compile(r'^(Fig(\.|ure)?|Table)\s*\d+[\.:]?$', re.IGNORECASE)

# Regex pattern matching back-matter index sections (Author Index, Subject Index) where extraction should stop
BACK_MATTER_PATTERN = re.compile(r'^(author|subject)\s+index\b', re.IGNORECASE)


def dehyphenate(text: str) -> str:
    """Remove hyphenated line breaks: 'preopera- \n tive' -> 'preoperative'."""
    if not text:
        return ""
    return re.sub(r'(\b\w+)-\s*\n\s*(\w+)', r'\1\2', text)


def clean_title(text: str) -> str:
    """Normalize presentation title or session heading into a clean single-line string."""
    if not text:
        return ""
    text = dehyphenate(text)
    # Collapse all whitespace and newlines to a single space
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def clean_text(text: str) -> str:
    """Fix hyphenated line breaks in body text while preserving paragraph structure."""
    if not text:
        return ""
    return dehyphenate(text).strip()


def is_author_underlined(author_spans: List[Tuple[str, List[float]]], underlines: List[pymupdf.Rect]) -> bool:
    """Check if an author's text span is underlined by a vector line below it."""
    for text, bbox in author_spans:
        sb = pymupdf.Rect(bbox)
        clean_t = text.lstrip(' ,')
        if not clean_t:
            continue
        for u in underlines:
            # Check vertical proximity (underline directly beneath text)
            if abs(sb.y1 - u.y0) < 4:
                # Check horizontal overlap
                overlap = max(0, min(sb.x1, u.x1) - max(sb.x0, u.x0))
                if overlap >= 15:
                    return True
    return False


def parse_authors_spans(author_block: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Parse an author text block using font spans to accurately separate
    author names from their superscript affiliation indices and detect underlines.
    """
    authors = []
    current_name = []
    current_spans = []
    current_indices = []
    page_underlines = author_block.get('page_underlines', [])

    for line in author_block.get('lines', []):
        for span in line.get('spans', []):
            text = span.get('text', '')
            size = span.get('size', 9.0)
            bbox = span.get('bbox', [0, 0, 0, 0])

            # Superscript markers are smaller font size (< 7.0pt)
            if size < 7.0 and re.search(r'[\d,]', text):
                digits = re.findall(r'\d+', text)
                current_indices.extend(digits)
            else:
                # Comma in regular text usually separates authors
                if ',' in text and current_name and current_indices:
                    parts = text.split(',')
                    current_name.append(parts[0])
                    current_spans.append((parts[0], bbox))
                    name_str = re.sub(r'\s+', ' ', ''.join(current_name)).strip(' ,;\xa0')
                    if name_str and not is_consortium(name_str):
                        underlined = is_author_underlined(current_spans, page_underlines)
                        authors.append({
                            'name': name_str,
                            'indices': current_indices,
                            'is_underlined': underlined
                        })
                    # Reset for subsequent author in the same span
                    current_name = [','.join(parts[1:])]
                    current_spans = [(','.join(parts[1:]), bbox)]
                    current_indices = []
                else:
                    current_name.append(text)
                    current_spans.append((text, bbox))

    # Process remaining author buffer
    if current_name:
        name_str = re.sub(r'\s+', ' ', ''.join(current_name)).strip(' ,;\xa0')
        if name_str and not is_consortium(name_str):
            underlined = is_author_underlined(current_spans, page_underlines)
            authors.append({
                'name': name_str,
                'indices': current_indices,
                'is_underlined': underlined
            })

    return authors


def is_consortium(name: str) -> bool:
    """Check if the extracted name corresponds to a collaborative study group rather than an individual."""
    name_lower = name.lower()
    return any(keyword in name_lower for keyword in NON_PERSON_KEYWORDS)


def parse_affil_spans(affil_blocks: List[Dict[str, Any]]) -> Dict[str, str]:
    """
    Extract numbered affiliations from affiliation blocks.
    Identifies index numbers via superscript font spans (< 7.0pt).
    """
    affils = {}
    current_num = '1'
    current_text = []

    for block in affil_blocks:
        for line in block.get('lines', []):
            for span in line.get('spans', []):
                text = span.get('text', '')
                size = span.get('size', 9.0)

                # Detected a new affiliation index number
                if size < 7.0 and text.strip().isdigit():
                    if current_text:
                        full_str = re.sub(r'\s+', ' ', ''.join(current_text)).strip(' ,;\xa0')
                        if full_str:
                            affils[current_num] = full_str
                        current_text = []
                    current_num = text.strip()
                else:
                    current_text.append(text)

    if current_text:
        full_str = re.sub(r'\s+', ' ', ''.join(current_text)).strip(' ,;\xa0')
        if full_str:
            affils[current_num] = full_str

    return affils


def extract_presentation_flow(doc: pymupdf.Document, start_page: int, end_page: int) -> Tuple[List[Dict[str, Any]], int, Optional[str]]:
    """
    Extract presentations across pages respecting the multi-column reading flow:
      Left column (top-to-bottom) -> Right column (top-to-bottom) -> Next page.
    """
    presentations = []
    current_session = "Best Abstract Prize Competition (BAPC)"
    curr_pres = None
    stop_extraction = False
    stop_reason: Optional[str] = None
    last_page = start_page - 1

    page_range = range(start_page - 1, min(end_page, len(doc)))
    for pno in tqdm(page_range, desc="Extracting pages", unit="page"):
        last_page = pno + 1
        page = doc[pno]

        # Stop extraction if reaching back-matter index section (e.g. Author Index on p. 777+)
        if pno >= 8:
            first_lines = [l.strip() for l in page.get_text().split('\n') if l.strip()][:4]
            for line in first_lines:
                m_bm = BACK_MATTER_PATTERN.search(line)
                if m_bm:
                    section_name = m_bm.group(0).title()
                    stop_reason = f"stopped at {section_name} on p. {pno + 1}"
                    logger.info(f"Reached back-matter index section ({section_name}) on page {pno + 1}. Stopping presentation extraction.")
                    stop_extraction = True
                    break
            if stop_extraction:
                break

        d = page.get_text('dict')
        
        # Collect underline drawings for presenting author detection
        drawings = page.get_drawings()
        page_underlines = [d_item['rect'] for d_item in drawings if abs(d_item['rect'].y1 - d_item['rect'].y0) < 3 and 15 < (d_item['rect'].x1 - d_item['rect'].x0) < 250 and d_item['rect'].y0 > 50]
        
        # Partition into two columns based on horizontal coordinate x0
        col_left = []
        col_right = []
        for b in d.get('blocks', []):
            if 'lines' not in b:
                continue
            bbox = b['bbox']
            # Exclude running page header at the top of page
            if bbox[1] < 45:
                continue
            if bbox[0] < 290:
                col_left.append((bbox[1], b))
            else:
                col_right.append((bbox[1], b))

        # Sort blocks vertically within each column
        col_left.sort(key=lambda item: item[0])
        col_right.sort(key=lambda item: item[0])
        ordered_blocks = [b for _, b in col_left] + [b for _, b in col_right]

        for b in ordered_blocks:
            first_span = b['lines'][0]['spans'][0]
            block_raw_text = '\n'.join([''.join([s['text'] for s in l['spans']]) for l in b['lines']]).strip()

            # Skip standalone visual artifact tags (e.g. 'Fig 1.', 'Table 1.')
            if STANDALONE_FIG_PATTERN.match(block_raw_text):
                continue

            # Stop extraction if an index section heading block is encountered
            m_bm = BACK_MATTER_PATTERN.match(block_raw_text)
            if m_bm:
                section_name = m_bm.group(0).title()
                stop_reason = f"stopped at {section_name} on p. {pno + 1}"
                logger.info(f"Reached back-matter index header '{block_raw_text}' on page {pno + 1}. Stopping presentation extraction.")
                stop_extraction = True
                break

            # Session Header detection (e.g. font size >= 12.0)
            if first_span['size'] >= 12.0 and not PRESENTATION_ID_PATTERN.match(block_raw_text):
                clean_sess = re.sub(r'^ESAIC\s+', '', block_raw_text).strip()
                current_session = clean_title(clean_sess)
                continue

            # Presentation ID detection (e.g., 'BAPC-03', '10AP01-2')
            m = PRESENTATION_ID_PATTERN.match(block_raw_text)
            if m:
                if curr_pres:
                    presentations.append(curr_pres)

                pres_id = m.group(1)
                title = clean_title(m.group(2))
                curr_pres = {
                    'session': current_session,
                    'id': pres_id,
                    'title': title,
                    'author_block': None,
                    'affil_blocks': [],
                    'body_blocks': [],
                    'state': 'AWAITING_AUTHORS'
                }
                continue

            if curr_pres:
                if curr_pres['state'] == 'AWAITING_AUTHORS':
                    b['page_underlines'] = page_underlines
                    curr_pres['author_block'] = b
                    curr_pres['state'] = 'AWAITING_AFFILS'
                elif curr_pres['state'] == 'AWAITING_AFFILS':
                    has_italic = any('I' in s['font'] or (s['flags'] & 2) for l in b['lines'] for s in l['spans'])
                    has_superscript = any(s['size'] < 7.0 for l in b['lines'] for s in l['spans'])
                    is_body_start = any(block_raw_text.startswith(h) for h in [
                        'Background', 'Case Report', 'Introduction', 'Objectives',
                        'Materials', 'Methods', 'Results'
                    ])

                    if (has_italic or has_superscript) and not is_body_start:
                        curr_pres['affil_blocks'].append(b)
                    else:
                        curr_pres['body_blocks'].append(block_raw_text)
                        curr_pres['state'] = 'IN_BODY'
                elif curr_pres['state'] == 'IN_BODY':
                    curr_pres['body_blocks'].append(block_raw_text)

        if stop_extraction:
            break

    if curr_pres:
        presentations.append(curr_pres)

    return presentations, last_page, stop_reason


def build_dataframe_rows(presentations: List[Dict[str, Any]], url: str = DEFAULT_ABSTRACT_URL) -> List[List[str]]:
    """Transform extracted presentations into flattened author-level spreadsheet rows."""
    rows = []

    for pres in presentations:
        session_name = pres['session']
        session_desc = pres['id']
        title = pres['title']

        # Format full abstract text
        raw_body = '\n'.join(pres['body_blocks'])
        abstract_text = clean_text(raw_body)

        # Parse authors and affiliations
        authors = parse_authors_spans(pres['author_block']) if pres['author_block'] else []
        affils = parse_affil_spans(pres['affil_blocks'])

        if not authors:
            logger.warning(f"[{session_desc}] No individual authors detected. Populating placeholder '-'.")
            rows.append([
                "-",
                "-",
                "-",
                session_name,
                session_desc,
                title,
                abstract_text,
                url
            ])
            continue

        # Check if an author is underlined to designate as Poster presenter
        has_underlined = any(a.get('is_underlined') for a in authors)

        for i, author_info in enumerate(authors):
            author_name = author_info['name']
            idx_list = author_info['indices']
            is_underlined = author_info.get('is_underlined', False)

            # If an author is underlined, they are Poster presenter; otherwise all are Abstract authors
            if has_underlined:
                role = "Poster presenter" if is_underlined else "Abstract author"
            else:
                role = "Abstract author"

            if not idx_list:
                logger.warning(f"[{session_desc}] Author '{author_name}' has no affiliation index -> assigned '-'")
                affiliation_str = "-"
            else:
                # Map author indices to institutions
                author_affils = [affils[idx] for idx in idx_list if idx in affils and affils[idx]]
                if author_affils:
                    affiliation_str = " ___ ".join(author_affils)
                else:
                    logger.warning(f"[{session_desc}] Affiliation index {idx_list} for '{author_name}' not found -> assigned '-'")
                    affiliation_str = "-"

            rows.append([
                author_name,
                affiliation_str,
                role,
                session_name,
                session_desc,
                title,
                abstract_text,
                url
            ])

        presenters = [a['name'] for a in authors if a.get('is_underlined')]
        presenter_str = ", ".join(presenters) if presenters else "None (all Abstract authors)"
        logger.info(f"[OK] {session_desc}: '{title[:40]}...' (Presenter: {presenter_str}, {len(authors)} authors, {len(affils)} affils)")

    return rows


def export_to_excel(rows: List[List[str]], output_path: str, template_path: str = None, include_template_rows: bool = False):
    """Save extracted rows into an Excel file, creating a new workbook or appending to template if requested."""
    if template_path and include_template_rows:
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file '{template_path}' not found.")
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Abstracts"
        ws.append(EXCEL_HEADERS)

    for row in rows:
        ws.append(row)

    wb.save(output_path)
    logger.info(f"[OK] Successfully wrote {len(rows)} author rows to {output_path}")


def print_ingestion_summary(
    start_page: int,
    end_page_scanned: int,
    presentations: List[Dict[str, Any]],
    rows: List[List[str]],
    output_path: str,
    elapsed_time: float,
    included_template_rows: bool = False,
    stop_reason: Optional[str] = None
):
    """Print an executive summary of the extraction process and data quality audit."""
    total_pres = len(presentations)
    total_authors = len(rows)
    presenter_count = sum(1 for r in rows if r[2] == 'Poster presenter')
    coauthor_count = sum(1 for r in rows if r[2] == 'Abstract author')
    missing_affil_count = sum(1 for r in rows if r[1] == '-')
    pages_count = max(0, end_page_scanned - start_page + 1)
    rate = pages_count / elapsed_time if elapsed_time > 0 else 0

    early_stop_note = f" ({stop_reason})" if stop_reason else ""

    print("\n" + "=" * 68)
    print("                    INGESTION & AUDIT SUMMARY")
    print("=" * 68)
    print(f"  Pages Scanned:            {start_page} to {end_page_scanned}{early_stop_note} ({pages_count} pages)")
    print(f"  Total Presentations:      {total_pres}")
    print(f"  Total Author Records:     {total_authors}")
    print(f"  Poster Presenters:        {presenter_count}")
    print(f"  Abstract Co-Authors:      {coauthor_count}")
    print(f"  Missing Affiliations:     {missing_affil_count} ({'All mapped [OK]' if missing_affil_count == 0 else f'{missing_affil_count} set to \"-\"'})")
    print(f"  Template Rows Included:   {'Yes (BAPC-01, BAPC-02)' if included_template_rows else 'No (clean extraction)'}")
    print(f"  Execution Time:           {elapsed_time:.2f} seconds ({rate:.1f} pages/sec)")
    print(f"  Target Output File:       {output_path}")
    print("=" * 68 + "\n")


def main():
    parser = argparse.ArgumentParser(description="Extract ESAIC 2025 congress abstracts into Monocl Excel schema.")
    parser.add_argument("--pdf", default="ESAIC2025_Abstracts-1.pdf", help="Path to ESAIC PDF")
    parser.add_argument("--template", default="Output_example.xlsx", help="Path to reference Excel template")
    parser.add_argument("--output", default="Output_completed.xlsx", help="Output Excel file path")
    parser.add_argument("--start-page", type=int, default=10, help="Start page (1-based PDF page number)")
    parser.add_argument("--end-page", type=int, default=14, help="End page (1-based PDF page number)")
    parser.add_argument("--keep-existing", action="store_true", default=False, help="Include existing rows from template (BAPC-01 and BAPC-02)")
    parser.add_argument("--debug", action="store_true", help="Enable verbose debug logging")
    args = parser.parse_args()

    # Fail fast: validate template upfront if user explicitly asked to keep it
    if args.keep_existing and not os.path.exists(args.template):
        logger.error(f"Template file '{args.template}' not found. Cannot keep existing rows.")
        sys.exit(1)

    # Configure logging
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S"
    )

    start_time = time.time()

    logger.info(f"Opening PDF: {args.pdf}")
    doc = pymupdf.open(args.pdf)
    logger.info(f"Extracting presentation flow for pages {args.start_page} through {args.end_page}...")

    presentations, last_page, stop_reason = extract_presentation_flow(doc, args.start_page, args.end_page)
    logger.info(f"[OK] Extracted {len(presentations)} presentations from pages {args.start_page}-{last_page}.")

    rows = build_dataframe_rows(presentations)
    logger.info(f"[OK] Generated {len(rows)} structured author records.")

    export_to_excel(rows, args.output, template_path=args.template, include_template_rows=args.keep_existing)

    elapsed_time = time.time() - start_time
    print_ingestion_summary(
        start_page=args.start_page,
        end_page_scanned=last_page,
        presentations=presentations,
        rows=rows,
        output_path=args.output,
        elapsed_time=elapsed_time,
        included_template_rows=args.keep_existing,
        stop_reason=stop_reason
    )


if __name__ == "__main__":
    main()
