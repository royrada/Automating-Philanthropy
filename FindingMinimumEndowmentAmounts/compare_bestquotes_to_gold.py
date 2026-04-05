
# This script compares BestQuotes.json to the gold standard and writes a Markdown summary for easy human review.
# Usage: python compare_bestquotes_to_gold.py


import os
import re
import json

testfile = 'testing/testURLs.txt'
md_lines = ['# BestQuotes vs Gold Standard Summary', '']
if not os.path.exists('test_outputs'):
    os.makedirs('test_outputs')

with open(testfile, encoding='utf-8') as f:
    for line in f:
        if not line.strip() or line.startswith('#'):
            continue
        parts = line.strip().split('\t')
        if len(parts) < 4:
            continue
        entity_url, quote_url, min_amt, gold_quote = parts[:4]
        entity = re.search(r'//([^/]+)/?', entity_url)
        if not entity:
            continue
        entity = entity.group(1).split('.')[-2]
        bestquotes_path = os.path.join('outputs', entity, 'BestQuotes.json')
        md_lines.append(f'\n## Entity: {entity}')
        md_lines.append(f'### URL: {quote_url}')
        md_lines.append(f'- **Gold Quote:** {gold_quote}')
        if not os.path.exists(bestquotes_path):
            md_lines.append(f'- **Result:** No BestQuotes file found.')
            continue
        # Find all BestQuotes for this URL
        matches = []
        with open(bestquotes_path, encoding='utf-8') as bqf:
            for line in bqf:
                try:
                    q = json.loads(line)
                except Exception:
                    continue
                if q.get('url') == quote_url:
                    matches.append(q)
        # Compute overlap score for each match
        def overlap_score(a, b):
            a_tokens = set(re.findall(r'\w+', a.lower()))
            b_tokens = set(re.findall(r'\w+', b.lower()))
            if not b_tokens:
                return 0.0
            return len(a_tokens & b_tokens) / len(b_tokens)
        best_score = 0.0
        best_quote = None
        for q in matches:
            score = overlap_score(q.get('quote',''), gold_quote)
            if score > best_score:
                best_score = score
                best_quote = q.get('quote','')
        # Write summary
        if matches:
            hitmiss = 'HIT' if best_score > 0.5 else 'MISS'
            md_lines.append(f'- **Best Overlap:** {best_score:.2f}')
            md_lines.append(f'- **HIT/MISS:** {hitmiss}')
            md_lines.append(f'- **Best Program Quote:**')
            md_lines.append(f'    > {best_quote.strip()}')
        else:
            md_lines.append(f'- **Result:** No BestQuotes found for this URL.')

with open('test_outputs/bestquotes_overlap_summary.md', 'w', encoding='utf-8') as outf:
    outf.write('\n'.join(md_lines))

print('Markdown summary written to test_outputs/bestquotes_overlap_summary.md')
