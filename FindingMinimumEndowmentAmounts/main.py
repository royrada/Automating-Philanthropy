import os
import re
import json
import requests
from urllib.parse import urlparse, urljoin
from collections import deque
from bs4 import BeautifulSoup
import time
import random
from scraper import (
    normalize_url, load_names, contains_name, delay,
    make_candidate_url, make_redflag_url, make_visited_url,
    ThrottleMax, AllowedURLfileExtension, URL_Bad, URL_Good
)
from quote_extraction import extract_visible_text, extract_quotes_from_text

# --- Sitemap and Homepage Extraction ---
def extract_sitemap_urls(entity_url, session):
    # Try to find sitemap.xml or robots.txt
    parsed = urlparse(entity_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    sitemap_urls = [urljoin(base, 'sitemap.xml'), urljoin(base, 'robots.txt')]
    found_urls = set()
    for sitemap_url in sitemap_urls:
        try:
            resp = session.get(sitemap_url, timeout=10)
            if resp.status_code == 200:
                if sitemap_url.endswith('robots.txt'):
                    for line in resp.text.splitlines():
                        if 'sitemap:' in line.lower():
                            found_urls.add(line.split(':',1)[1].strip())
                else:
                    found_urls.add(sitemap_url)
        except Exception:
            continue
    # Parse each sitemap (one level deep)
    all_urls = set()
    for sm_url in found_urls:
        try:
            resp = session.get(sm_url, timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'xml')
                for loc in soup.find_all('loc'):
                    all_urls.add(loc.text.strip())
        except Exception:
            continue
    return list(all_urls)[:3000]

def extract_homepage_urls(entity_url, session):
    try:
        resp = session.get(entity_url, timeout=10)
        if resp.status_code != 200:
            return []
        soup = BeautifulSoup(resp.text, 'html.parser')
        links = set()
        for a in soup.find_all('a', href=True):
            href = a['href']
            if href.startswith(('mailto:', 'tel:', 'javascript:')) or href.startswith('#'):
                continue
            full_url = urljoin(entity_url, href)
            links.add(full_url)
        return list(links)
    except Exception:
        return []

# --- RedFlag Filtering ---
def filter_candidate_urls(candidate_urls, entity_url, names):
    redflags = []
    openurls = []
    parsed_entity = urlparse(entity_url)
    entity_domain = parsed_entity.netloc.split('.')[-2]
    seen_norm = set()
    for url in candidate_urls:
        parsed = urlparse(url)
        norm = normalize_url(url)
        if norm in seen_norm:
            redflags.append(make_redflag_url(url, 'duplicate'))
            continue
        seen_norm.add(norm)
        if parsed.scheme not in ('http', 'https'):
            redflags.append(make_redflag_url(url, 'bad_protocol'))
            continue
        if any(x in parsed.path for x in ['#', '?', '%']) or re.search(r'\d', parsed.path):
            redflags.append(make_redflag_url(url, 'bad_path'))
            continue
        ext = os.path.splitext(parsed.path)[1]
        if ext and ext not in AllowedURLfileExtension:
            redflags.append(make_redflag_url(url, 'bad_extension'))
            continue
        url_domain = parsed.netloc.split('.')[-2]
        if url_domain != entity_domain:
            redflags.append(make_redflag_url(url, 'domain_mismatch'))
            continue
        # Fix: check for bad terms in the full normalized path, not just domain
        if any(bad in norm for bad in URL_Bad):
            redflags.append(make_redflag_url(url, 'bad_string'))
            continue
        if contains_name(norm, names, use_regex=False):
            redflags.append(make_redflag_url(url, 'name'))
            continue
        # Passed all filters
        score = sum(1 for good in URL_Good if good in norm)
        openurls.append((-score, len(url), url))
    return redflags, openurls

# --- Quote Extraction ---
def extract_quotes(text, url, names):
    quotes = []
    rejected = []
    for min_amt in AllowedMins:
        for m in re.finditer(re.escape(min_amt), text):
            start = max(0, m.start() - 120)
            end = min(len(text), m.end() + 120)
            window = text[start:end]
            # Merge overlapping windows is handled by set
            if contains_name(window, names, use_regex=True):
                rejected.append(make_rejected_quote(window, url, 'name'))
                continue
            if any(bad in window for bad in badSet):
                rejected.append(make_rejected_quote(window, url, 'badSet'))
                continue
            creationScore = sum(1 for c in creationSet if c in window)
            fundScore = sum(1 for f in fundSet if f in window)
            minScore = sum(1 for mn in minSet if mn in window)
            sumScore = creationScore + fundScore + minScore
            quotes.append(make_candidate_quote(window, url, creationScore, fundScore, minScore, sumScore))
    return quotes, rejected

# --- Main Scraper Function ---

# --- Helper for deduplication ---
def add_url_if_unique(url, target_list, all_seen_urls, make_func, *args):
    norm = normalize_url(url)
    if norm in all_seen_urls:
        return False
    all_seen_urls.add(norm)
    target_list.append(make_func(url, *args))
    return True

def process_entity(entity_url, names):
    session = requests.Session()
    # Use a single, modern, hardcoded user-agent for all requests
    session.headers['User-Agent'] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    # Data structures
    CandidateURLs = []
    RedFlagURLs = []
    VisitedURLs = []
    CandidateQuotes = []
    RejectedQuotes = []
    OpenQuotes = []
    BestQuotes = []
    # Global set for all normalized URLs
    all_seen_urls = set()
    # Collect URLs
    urls = extract_sitemap_urls(entity_url, session)
    print(f"Initial sitemap URLs: {len(urls)}")
    if not urls:
        urls = extract_homepage_urls(entity_url, session)
        print(f"Homepage URLs: {len(urls)}")
        source = 'homepage'
    else:
        source = 'sitemap'
    for url in urls:
        add_url_if_unique(url, CandidateURLs, all_seen_urls, make_candidate_url, source)
    print(f"Initial CandidateURLs: {len(CandidateURLs)}")
    # RedFlag filtering
    redflags, openurls = filter_candidate_urls([c['url'] for c in CandidateURLs], entity_url, names)
    # Only add unique redflags
    for rf in redflags:
        add_url_if_unique(rf['url'], RedFlagURLs, all_seen_urls, make_redflag_url, rf['criterion'])
    print(f"Initial RedFlagURLs: {len(RedFlagURLs)}; openurls: {len(openurls)}")
    openurls = deque(sorted(openurls))
    openurls_norm_set = set(normalize_url(url) for _, _, url in openurls)
    visited_count = 0
    while openurls and visited_count < ThrottleMax:
        print(f"Openurls left: {len(openurls)}; Visited: {visited_count}")
        score_tuple = openurls.popleft()
        score, length, url = score_tuple
        norm = normalize_url(url)
        openurls_norm_set.discard(norm)
        if norm in all_seen_urls:
            continue
        all_seen_urls.add(norm)
        # RedFlag check before visiting
        redflagged = False
        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https'):
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'bad_protocol')
            redflagged = True
        if any(x in parsed.path for x in ['#', '?', '%']) or re.search(r'\d', parsed.path):
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'bad_path')
            redflagged = True
        ext = os.path.splitext(parsed.path)[1]
        if ext and ext not in AllowedURLfileExtension:
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'bad_extension')
            redflagged = True
        url_domain = parsed.netloc.split('.')[-2]
        entity_domain = urlparse(entity_url).netloc.split('.')[-2]
        if url_domain != entity_domain:
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'domain_mismatch')
            redflagged = True
        norm_url = normalize_url(url)
        if any(bad in norm_url for bad in URL_Bad):
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'bad_string')
            redflagged = True
        if contains_name(norm_url, names, use_regex=False):
            add_url_if_unique(url, RedFlagURLs, all_seen_urls, make_redflag_url, 'name')
            redflagged = True
        if redflagged:
            continue
        status = None
        error_message = None
        try:
            delay()
            resp = session.get(url, timeout=10)
            status = resp.status_code
            if resp.status_code == 200:
                # Extract more URLs (fix: resolve relative URLs and add as candidates)
                soup = BeautifulSoup(resp.text, 'html.parser')
                new_candidates = []
                for a in soup.find_all('a', href=True):
                    href = a['href']
                    if href.startswith(('mailto:', 'tel:', 'javascript:')) or href.startswith('#'):
                        continue
                    full_url = urljoin(url, href)
                    add_url_if_unique(full_url, CandidateURLs, all_seen_urls, make_candidate_url, 'crawl')
                    new_candidates.append(full_url)
                # After adding, filter and score new candidates
                redflags, new_openurls = filter_candidate_urls(new_candidates, entity_url, names)
                for rf in redflags:
                    add_url_if_unique(rf['url'], RedFlagURLs, all_seen_urls, make_redflag_url, rf['criterion'])
                for item in new_openurls:
                    _, _, new_url = item
                    norm_new = normalize_url(new_url)
                    if norm_new not in all_seen_urls:
                        openurls.append(item)
                        openurls_norm_set.add(norm_new)
                openurls = deque(sorted(openurls))
                # Extract visible text and quotes using new logic
                visible_text = extract_visible_text(resp.text)
                quotes, rejected, openquotes = extract_quotes_from_text(visible_text, url, set(names))
                CandidateQuotes.extend(quotes)
                RejectedQuotes.extend(rejected)
                OpenQuotes.extend(openquotes)
        except Exception as e:
            error_message = str(e)
        # Store the correct score (negate back to positive)
        VisitedURLs.append({
            "url": url,
            "score": -score,  # score is stored as negative for priority queue
            "http_status": status,
            "error": error_message
        })
        visited_count += 1
        print(f"Visited {url} (status: {status}, error: {error_message}, score: {-score})")
    # BestQuotes selection from OpenQuotes
    if len(OpenQuotes) <= 5:
        BestQuotes = OpenQuotes.copy()
    else:
        filtered = [q for q in OpenQuotes if q['minScore'] > 0]
        if not filtered:
            filtered = OpenQuotes.copy()
        filtered.sort(key=lambda q: -q['sumScore'])
        BestQuotes = filtered[:5]
    print(f"Final: CandidateURLs={len(CandidateURLs)}, RedFlagURLs={len(RedFlagURLs)}, VisitedURLs={len(VisitedURLs)}, CandidateQuotes={len(CandidateQuotes)}, OpenQuotes={len(OpenQuotes)}, BestQuotes={len(BestQuotes)}")
    return CandidateURLs, RedFlagURLs, VisitedURLs, CandidateQuotes, RejectedQuotes, OpenQuotes, BestQuotes

# --- Output Functions ---
def write_outputs(entity, outdir, CandidateURLs, RedFlagURLs, VisitedURLs, CandidateQuotes, RejectedQuotes, OpenQuotes, BestQuotes):
    os.makedirs(outdir, exist_ok=True)
    def write_json(name, data):
        with open(os.path.join(outdir, name + '.json'), 'w', encoding='utf-8') as f:
            for item in data:
                f.write(json.dumps(item, ensure_ascii=False) + '\n')
    def write_md(name, data):
        if not data:
            return
        keys = data[0].keys()
        with open(os.path.join(outdir, name + '.md'), 'w', encoding='utf-8') as f:
            f.write('| ' + ' | '.join(keys) + ' |\n')
            f.write('| ' + ' | '.join(['---']*len(keys)) + ' |\n')
            for item in data:
                f.write('| ' + ' | '.join(str(item[k]) for k in keys) + ' |\n')
    write_json('CandidateURLs', CandidateURLs)
    write_json('RedFlagURLs', RedFlagURLs)
    write_json('VisitedURLs', VisitedURLs)
    write_json('CandidateQuotes', CandidateQuotes)
    write_json('RejectedQuotes', RejectedQuotes)
    write_json('OpenQuotes', OpenQuotes)
    write_json('BestQuotes', BestQuotes)
    write_md('CandidateURLs', CandidateURLs)
    write_md('RedFlagURLs', RedFlagURLs)
    write_md('VisitedURLs', VisitedURLs)
    write_md('CandidateQuotes', CandidateQuotes)
    write_md('RejectedQuotes', RejectedQuotes)
    write_md('OpenQuotes', OpenQuotes)
    # Custom BestQuotes.md output (hierarchy and bullets)
    best_md_path = os.path.join(outdir, 'BestQuotes.md')
    with open(best_md_path, 'w', encoding='utf-8') as f:
        for idx, q in enumerate(BestQuotes[:5], 1):
            f.write(f"\n## Best Quote #{idx}\n")
            f.write(f"### URL on which quote was found\n")
            f.write(f"{q.get('url','')}\n")
            f.write(f"#### Minimal dollar amount\n")
            f.write(f"{q.get('minAmount','')}\n")
            f.write(f"- Quote verbatim: {q.get('quote','').strip()}\n")
            f.write(f"- creationScore: {q.get('creationScore','')}\n")
            f.write(f"- fundScore: {q.get('fundScore','')}\n")
            f.write(f"- minScore: {q.get('minScore','')}\n")
            f.write(f"- sumScore: {q.get('sumScore','')}\n")

# --- Main Entrypoint ---
def main():
    print('>>> main() entry')
    with open('inputs/URLs.txt', encoding='utf-8') as f:
        entity_urls = [line.strip() for line in f if line.strip()]
    names = load_names('inputs/names.txt')
    for entity_url in entity_urls:
        entity = urlparse(entity_url).netloc.split('.')[-2]
        outdir = os.path.join('outputs', entity)
        CandidateURLs, RedFlagURLs, VisitedURLs, CandidateQuotes, RejectedQuotes, OpenQuotes, BestQuotes = process_entity(entity_url, names)
        write_outputs(entity, outdir, CandidateURLs, RedFlagURLs, VisitedURLs, CandidateQuotes, RejectedQuotes, OpenQuotes, BestQuotes)

if __name__ == '__main__':
    main()
