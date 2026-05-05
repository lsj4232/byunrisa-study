#!/usr/bin/env python3
import subprocess
import re
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Hand-supplied answer text that overrides whatever pdftotext / docx
# extraction produces for the listed (exam, problemNo, sqNo). Use for cases
# where the source file is missing a section.
ANSWER_OVERRIDES = {
    ('62회', 1, 2): """1. 문제의 소재
丙은 소송계속 중 계쟁물 승계인으로 인수승계의 요건은 만족하나, 인수승계의 절차와 관련하여 추가적 인수는 불가하여 법원은 각하 결정을 내릴 것임을 알아본다.
2. 인수승계의 의의 및 취지(제82조)
소송계속 중 소송목적인 권리의무 승계가 있는 경우 종전당사자 신청에 의해 승계인을 새로운 당사자로 끌어들이는 것으로서 소송경제를 도모하기 위함이다.
3. 요건
(1) ⅰ) 타인 간 소송 계속 중 ⅱ) 제3자에게 소송의 목적인 권리,의무의 승계가 있어야 한다.
(2) 사안
ⅰ) 甲,乙간의 소송 계속 중 ⅱ) 丙은 계쟁물인 건물을 승계하였고, 甲,乙간의 건물철거 소송은 물권적 청구권에 기한 것이므로 구이론(判例)에 따라 丙은 승계인에 해당한다. 다만, 丙에 대해서는 등기말소를 구했는 바, 추가적 인수 가부가 문제된다.
4. 절차 - 추가적 인수 허용여부
(1) 문제점
인수승계의 절차에는, ⅰ) 소송 목적인 채무 자체를 승계하는 교환적 인수와 ⅱ) 소송 목적인 채무 승계가 아니라, 그 채무를 전제로 새로운 채무가 생기는 추가적 인수가 있다.
(2) 학설
① 긍정설은 이는 분쟁주체 지위의 이전으로서 소송경제상 가능하다고 보나 ② 부정설은 소송목적인 권리 또는 의무를 승계한 때라고 규정한 제82조에 반하므로 부정한다.
(3) 판례
① 당사자가 제3자로 하여금 소송인수를 하기 위해서는 제3자에 대해 인수한 소송의 목적이 된 채무이행을 구하는 경우만 허용되고 별개 채무의 이행을 구함은 허용될 수 없다고 하나 ② 공유물분할 소송에서 지분 일부양수인을 추가하는 경우는 허용했다.
(4) 검토
추가적 인수를 부정한 듯한 판례는 존재하나, 이는 소송목적인 권리·의무를 승계했다고 볼 수 없다는 이유로 부정한 것인바, 소송목적의 권리·의무를 승계했다고 볼 수 있다면 허용함이 타당하다.
(5) 사안
乙이 건물철거의무를 진다고 하여, 丙이 등기말소의무를 부담하는 것이 아니므로, 丙은 소송목적의 의무를 승계했다고 볼 수 없다. 따라서, 추가적 인수는 불허된다.
5. 설문해결
주장 자체로 인수승계 신청이 부적법한 경우이므로, 법원은 각하결정을 해야 한다.""",
}

# Hand-supplied question text overrides for cases where extraction loses
# part of the prompt.
QUESTION_OVERRIDES = {
    ('61회', 3, 3): "항소법원은 심리 후 두 개의 청구를 전부 인용하는 판결을 하였다. 乙은 이 가운데 제2의 소에 대한 판결에만 상고하였다. 이러한 경우 甲은 '제1의 소에 대한 판결에는 상고하지 않았으므로 이 판결은 확정되었다'는 이유로 A토지에 대한 소유권이전등기절차를 실행할 수 있는지 검토하시오.",
}

# (subject, exam, year, question_file, answer_file). Paths are relative to BASE_DIR.
EXAMS = [
    ('민사소송법', '56회', 2019, '56민사소송법.docx', '제56회_답안.docx'),
    ('민사소송법', '57회', 2020, '57민사소송법.docx', '제57회_답안.docx'),
    ('민사소송법', '58회', 2021, '58회민사소송법[3].pdf', '제58회_민사소송법_답안.docx'),
    ('민사소송법', '59회', 2022, '59회민사소송법.pdf', '제59회_답안.docx'),
    ('민사소송법', '60회', 2023, '60회민사소송법.pdf', '제60회_민사소송법_답안.docx'),
    ('민사소송법', '61회', 2024, '61회민사소송법.pdf', '제61회_답안.docx'),
    ('민사소송법', '62회', 2025, '62회민사소송법.pdf', '제62회_답안.docx'),
    ('특허법',     '62회', 2025, '특허법/2025년 제62회 변리사 2차 시험 1교시(특허법).pdf', '특허법/제62회_특허법_답안.docx'),
]

def extract_docx_text(filepath):
    import zipfile
    from xml.etree import ElementTree as ET
    NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    with zipfile.ZipFile(filepath) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    parts = []
    for para in root.iter(NS + 'p'):
        line = ''.join(t.text or '' for t in para.iter(NS + 't'))
        parts.append(line)
    return '\n'.join(parts)

def extract_pdf_text(filepath, layout=False):
    if filepath.endswith('.docx'):
        return extract_docx_text(filepath)

    # If a sibling .ocr.txt exists, prefer it (PDFs whose embedded fonts
    # don't extract via pdftotext but were OCR'd via macOS Vision API)
    base = os.path.basename(filepath)
    # Strip [N] disambiguation suffix from the OCR cache filename
    cache_base = re.sub(r'\[\d+\]\.pdf$', '.pdf', base)
    cache_path = os.path.join(os.path.dirname(filepath),
                              cache_base.replace('.pdf', '.ocr.txt'))
    if os.path.exists(cache_path):
        with open(cache_path, 'r', encoding='utf-8') as f:
            text = f.read()
        # OCR consistently misreads 乙 as the Latin letter Z; restore it
        # only when followed by a Hangul josa (safe — real Latin Z won't appear
        # immediately before Korean characters in a question stem).
        text = re.sub(r'Z(?=[가-힣])', '乙', text)
        return text

    cmd = ['pdftotext']
    if layout:
        cmd.append('-layout')
    cmd += [filepath, '-']
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if result.returncode != 0:
        print(f"ERROR extracting {filepath}: {result.stderr}", file=sys.stderr)
    return result.stdout

def clean_text(text, normalize_lines=False):
    # Normalize form feeds (PDF page breaks) to newlines
    text = text.replace('\x0c', '\n')
    # When using -layout mode, normalize each line (strip + collapse internal spaces)
    if normalize_lines:
        lines = []
        for line in text.split('\n'):
            stripped = line.strip()
            # Collapse multiple spaces within a line to single space
            stripped = re.sub(r' {2,}', ' ', stripped)
            lines.append(stripped)
        text = '\n'.join(lines)
    # Normalize problem-header variants to the canonical "【 문제-N 】" form:
    #   "[ 문제-1 ]" / "[ 문제-1 )" / "[문제 - 1]"  (OCR / docx)
    #   "〈 문제-1 〉"                              (특허법 docx)
    text = re.sub(r'\[\s*문제\s*-\s*(\d+)\s*[\])]', r'【 문제-\1 】', text)
    text = re.sub(r'[〈<]\s*문제\s*-\s*(\d+)\s*[〉>]', r'【 문제-\1 】', text)
    # Normalize spacing in points notation: "(10점 )" → "(10점)"
    text = re.sub(r'\(\s*(\d+)\s*점\s*\)', r'(\1점)', text)
    # Remove page markers (question PDFs) — covers 민사소송법 and 특허법
    text = re.sub(r'\d{4}년 제\d+회 변리사 2차 - (?:민사소송법|특허법) - ?\d교시 \(.*?\)', '', text)
    # Remove answer file headers
    text = re.sub(r'윤곽 민사소송법 기출문제 \d+회', '', text)
    # Remove page number lines (e.g. "1쪽", "12쪽")
    text = re.sub(r'^\s*\d+쪽\s*$', '', text, flags=re.MULTILINE)
    # Remove single/double Hangul character lines (decorative sidebar chars like 단,공,력,인,업,산,국,한)
    text = re.sub(r'^\s*[가-힣]{1,2}\s*$', '', text, flags=re.MULTILINE)
    # Collapse 3+ blank lines to 2
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def parse_questions(text, exam_name, year, normalize_lines=False, subject='민사소송법'):
    text = clean_text(text, normalize_lines=normalize_lines)

    # Split by problem header: 【 문제-N 】 (Ypoints) — allow optional space
    # before the points parenthesis (62회 has "【 문제-1 】(30점)" with no space).
    parts = re.split(r'【\s*문제-(\d+)\s*】\s*\(\s*(\d+)\s*점\s*\)', text)
    # parts[0] = text before first problem (discard)
    # then groups of 3: [number, points, content]

    problems = []
    for i in range(1, len(parts), 3):
        prob_no = int(parts[i])
        prob_points = int(parts[i + 1])
        prob_body = parts[i + 2].strip()

        # Separate passage from sub-questions
        # Sub-questions start with (N) at the beginning of a line
        # Use lookahead to split without losing the (N) prefix
        sq_split = re.split(r'\n(?=\(\d\) )', prob_body)

        passage_raw = sq_split[0]
        # Strip "(다음 각 물음/설문은 독립적임/이다.)" / "(다음 각 물음은 독립적이다.)" etc.
        passage = re.sub(
            r'\(\s*다음 각 (?:물음|설문)(?:은)?\s*독립적(?:임|이다\.?)\s*\)\s*$',
            '', passage_raw).strip()
        # Also handle version without the separator (passage just ends before (1))
        passage = clean_passage(passage)

        subquestions = []
        for sq_part in sq_split[1:]:
            m = re.match(r'\((\d)\)\s+(.+)', sq_part, re.DOTALL)
            if not m:
                continue
            sq_no = int(m.group(1))
            sq_text = m.group(2).strip()

            # Extract points — check end first, then anywhere in text
            pts_m = re.search(r'\((\d+)점\)\s*$', sq_text)
            if not pts_m:
                pts_m = re.search(r'\((\d+)점\)', sq_text)
            sq_points = int(pts_m.group(1)) if pts_m else 0
            # Remove ALL occurrences of the points notation
            sq_text = re.sub(r'\s*\(\d+점\)', '', sq_text).strip()

            # Normalize whitespace within text
            sq_text = normalize_whitespace(sq_text)

            subquestions.append({
                'no': sq_no,
                'points': sq_points,
                'question': sq_text,
                'answer': ''
            })

        problems.append({
            'subject': subject,
            'exam': exam_name,
            'year': year,
            'problemNo': prob_no,
            'points': prob_points,
            'passage': normalize_whitespace(passage),
            'subquestions': subquestions
        })

    return problems

def clean_passage(text):
    # Remove any trailing sub-question-like lines that shouldn't be in passage
    # (shouldn't happen after splitting, but just in case)
    return text.strip()

def normalize_whitespace(text):
    # Collapse multiple blank lines to one
    text = re.sub(r'\n{2,}', '\n', text)
    # Remove trailing spaces on each line
    lines = [line.rstrip() for line in text.split('\n')]
    return '\n'.join(lines).strip()

def parse_answers(text, exam_name):
    text = clean_text(text)

    # Split by problem header: [문제-N] or 【 문제-N 】 (latter from normalization)
    parts = re.split(r'(?:\[문제-|【\s*문제-)(\d+)(?:\]|\s*】)', text)
    # parts[0] = before first (discard)
    # then groups of 2: [number, content]

    answer_map = {}  # key: (prob_no, sq_no) -> answer text

    for i in range(1, len(parts), 2):
        prob_no = int(parts[i])
        prob_body = parts[i + 1].strip()

        # Split by roman-numeral sub-question marker. Two flavors:
        #   Unicode roman:    "Ⅰ. 설문(1)", "Ⅱ. 설문(2)"
        #   Latin letters:    "I. 설문(1)", "II. 설문(2)" (62회 docx)
        # Anchor at line start to avoid stray "I." in body text.
        sq_parts = re.split(
            r'(?:^|\n)\s*(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨ]|I{1,3}|IV|V|VI{0,3}|IX|X)\s*\.\s*설문\s*\(\s*(\d+)\s*\)',
            prob_body)
        # sq_parts[0] = before first (discard)
        # then groups of 2: [number, content]

        for j in range(1, len(sq_parts), 2):
            sq_no = int(sq_parts[j])
            sq_answer = sq_parts[j + 1].strip()
            sq_answer = normalize_whitespace(sq_answer)
            key = (prob_no, sq_no)
            if key in answer_map:
                # Same sub-question split into multiple roman-numeral sections
                # (e.g. "I. 설문(1) - 1)" + "II. 설문(1) - 2)") — concatenate.
                answer_map[key] = answer_map[key] + '\n\n' + sq_answer
            else:
                answer_map[key] = sq_answer

    return answer_map

def main():
    all_problems = []

    # Preserve hand-curated `categories` and `hints` (LLM-style data) from
    # any prior data.json so re-running the parser doesn't wipe them.
    out_path = os.path.join(BASE_DIR, 'data.json')
    existing_categories = {}
    existing_hints = {}
    if os.path.exists(out_path):
        try:
            with open(out_path, 'r', encoding='utf-8') as f:
                existing = json.load(f)
            for prob in existing:
                subj = prob.get('subject', '민사소송법')
                for sq in prob.get('subquestions', []):
                    key = (subj, prob['exam'], prob['problemNo'], sq['no'])
                    if 'categories' in sq:
                        existing_categories[key] = sq['categories']
                    if 'hints' in sq:
                        existing_hints[key] = sq['hints']
        except Exception as e:
            print(f"  WARN: couldn't load existing curated data: {e}", file=sys.stderr)

    for subject, exam_name, year, q_file, a_file in EXAMS:
        q_path = os.path.join(BASE_DIR, q_file)
        a_path = os.path.join(BASE_DIR, a_file)

        print(f"Processing {subject} {exam_name}...", file=sys.stderr)
        q_text = extract_pdf_text(q_path, layout=True)
        a_text = extract_pdf_text(a_path, layout=False)

        problems = parse_questions(q_text, exam_name, year, normalize_lines=True, subject=subject)
        answers = parse_answers(a_text, exam_name)

        # Merge answers + restore preserved category data + apply overrides
        for prob in problems:
            for sq in prob['subquestions']:
                key = (prob['problemNo'], sq['no'])
                if key in answers:
                    sq['answer'] = answers[key]
                else:
                    print(f"  WARNING: no answer for {subject} {exam_name} 문제-{prob['problemNo']} 설문({sq['no']})",
                          file=sys.stderr)
                ov_key = (exam_name, prob['problemNo'], sq['no'])
                if ov_key in ANSWER_OVERRIDES:
                    sq['answer'] = ANSWER_OVERRIDES[ov_key]
                if ov_key in QUESTION_OVERRIDES:
                    sq['question'] = QUESTION_OVERRIDES[ov_key]
                cat_key = (subject, exam_name, prob['problemNo'], sq['no'])
                if cat_key in existing_categories:
                    sq['categories'] = existing_categories[cat_key]
                if cat_key in existing_hints:
                    sq['hints'] = existing_hints[cat_key]

        all_problems.extend(problems)
        total_sq = sum(len(p['subquestions']) for p in problems)
        print(f"  {len(problems)} problems, {total_sq} sub-questions", file=sys.stderr)

    # Save data.json
    out_path = os.path.join(BASE_DIR, 'data.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(all_problems, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {out_path}", file=sys.stderr)

    # Embed into index.html as a JS variable
    html_path = os.path.join(BASE_DIR, 'index.html')
    if os.path.exists(html_path):
        with open(html_path, 'r', encoding='utf-8') as f:
            html = f.read()

        json_str = json.dumps(all_problems, ensure_ascii=False)
        new_data_block = f'const EXAM_DATA = {json_str};'

        # Replace the data block — use lambda to avoid re.sub backslash processing
        if re.search(r'const EXAM_DATA = \[', html, re.DOTALL):
            html_new = re.sub(
                r'const EXAM_DATA = \[.*?\];',
                lambda m: new_data_block,
                html,
                flags=re.DOTALL
            )
        else:
            # Placeholder not yet replaced — find the placeholder comment
            html_new = html.replace('/* DATA_PLACEHOLDER */', new_data_block)

        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_new)
        print(f"Embedded data into {html_path}", file=sys.stderr)
    else:
        print("index.html not found — skipping embed", file=sys.stderr)

    total = sum(len(p['subquestions']) for p in all_problems)
    print(f"\nDone: {len(all_problems)} problems, {total} sub-questions total", file=sys.stderr)

if __name__ == '__main__':
    main()
