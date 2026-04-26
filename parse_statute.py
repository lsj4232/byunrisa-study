#!/usr/bin/env python3
"""Extract statute articles from the 민사소송법 PDF, write them into
data.json under top-level key STATUTE so the page bundles them with the
existing EXAM_DATA payload, and re-embed into index.html."""
import subprocess
import re
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# (subject, relative_pdf_path)
STATUTES = [
    ('민사소송법', '민사소송법(법률)(제19516호)(20250712).pdf'),
    ('특허법',     '특허법/특허법(법률)(제21134호)(20251111).pdf'),
]


def extract_statute_text(rel_path):
    out = subprocess.run(['pdftotext', '-layout',
                          os.path.join(BASE_DIR, rel_path), '-'],
                         capture_output=True, text=True, encoding='utf-8')
    return out.stdout


def clean(text, subject):
    # Strip page-footer banner and page numbers
    text = re.sub(r'법제처\s+\d+\s+국가법령정보센터', '', text)
    text = re.sub(rf'\s*{re.escape(subject)}\s*\n', '\n', text)
    text = text.replace('\x0c', '\n')
    return text


def parse_articles(text, subject):
    """Returns list of {no:int, no_str:str, title:str, body:str}.
    When two versions of the same article exist (current + future-effective),
    keep both, distinguished by `variant` index."""
    text = clean(text, subject)

    # Split on each "제N조" header, capturing the marker so we can re-attach.
    # Article header looks like: 제3조(사람의 보통재판적)  or  제3조의2(...) or 제13조(...) without title
    pattern = re.compile(
        r'(제(\d+)조(?:의\s*\d+)?)'           # 1=full marker, 2=number
        r'(?:\(([^)]*)\))?'                   # 3=optional title
        r'(.*?)'                              # 4=body
        r'(?=\n제\d+조|\n\s*제\d+편|\n\s*제\d+장|\n\s*제\d+절|\n\s*\[시행일|\Z)',
        re.DOTALL,
    )

    by_no = {}
    order = []
    for m in pattern.finditer(text):
        marker = m.group(1).replace(' ', '')
        try:
            no = int(m.group(2))
        except ValueError:
            continue
        title = (m.group(3) or '').strip()
        body = (m.group(4) or '').strip()

        # Collapse runaway whitespace; keep paragraph breaks.
        body = re.sub(r'[ \t]+', ' ', body)
        body = re.sub(r' *\n *', '\n', body)
        body = re.sub(r'\n{2,}', '\n', body).strip()

        article = {
            'no': no,
            'no_str': marker,
            'title': title,
            'body': body,
        }

        if marker in by_no:
            # Multiple versions (e.g. 제13조 has current + future-effective text).
            # Keep them concatenated with a separator so the side panel shows all.
            existing = by_no[marker]
            existing['body'] += '\n\n— 시행 예정 —\n' + body
        else:
            by_no[marker] = article
            order.append(marker)

    return [by_no[k] for k in order]


def main():
    by_subject = {}
    for subject, rel_path in STATUTES:
        raw = extract_statute_text(rel_path)
        arts = parse_articles(raw, subject)
        print(f"Parsed {len(arts)} articles ({subject})", file=sys.stderr)
        by_subject[subject] = arts

    out_path = os.path.join(BASE_DIR, 'statute.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(by_subject, f, ensure_ascii=False, indent=2)
    print(f"Saved {out_path}", file=sys.stderr)

    # Embed into index.html as STATUTE_DATA — now an object keyed by subject
    html_path = os.path.join(BASE_DIR, 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    json_str = json.dumps(by_subject, ensure_ascii=False)
    block = f'const STATUTE_DATA = {json_str};'

    # Replace either the old array form or the new object form
    if re.search(r'const STATUTE_DATA = (?:\[|\{)', html, re.DOTALL):
        html_new = re.sub(
            r'const STATUTE_DATA = (?:\[.*?\]|\{.*?\});',
            lambda mm: block,
            html,
            flags=re.DOTALL,
        )
    else:
        html_new = re.sub(
            r'(const EXAM_DATA = \[.*?\];)',
            lambda mm: mm.group(1) + '\n' + block,
            html,
            count=1,
            flags=re.DOTALL,
        )

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_new)
    print(f"Embedded statute into {html_path}", file=sys.stderr)


if __name__ == '__main__':
    main()
