import re
from typing import List, Dict, Set

# --- Constants (should be imported from config or main if needed) ---
creationSet = {'agree', 'award', 'charit', 'create', 'donor', 'endow', 'establish', 'fellow', 'fund', 'give', 'giving', 'gift', 'legacy', 'paid'}
fundSet = {'fellowship', 'field of interest', 'scholar'}
minSet = {'minim', 'or more', 'at least', 'begin', 'start', 'initial'}
badSet = {'loan', 'maximum', 'tier'}
AllowedMins = {'$5,000', '$10,000', '$15,000', '$20,000', '$25,000', '$30,000', '$40,000', '$50,000', '$75,000', '$100,000', '$120,000', '$125,000', '$150,000', '$170,000', '$175,000', '$200,000', '$220,000', '$250,000', '$300,000'}


def extract_visible_text(html: str) -> str:
    """Extract visible text from HTML (strip tags/scripts/styles)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    # Remove script and style elements (no longer remove tables)
    for tag in soup(['script', 'style']):
        tag.decompose()
    return soup.get_text(separator=' ', strip=True)


from typing import Tuple
def extract_quotes_from_text(text: str, url: str, names: Set[str]) -> Tuple[List[Dict], List[Dict], List[Dict]]:
    """
    Implements the Plan.md quote extraction logic:
    - For each AllowedMins instance, extract 120 chars left/right (QuoteString)
    - Reject if QuoteString contains a name (>3 chars) from names.txt
    - Reject if QuoteString contains a badSet word
    - Score and move to OpenQuotes if valid
    Returns: CandidateQuotes, RejectedQuotes, OpenQuotes
    """
    CandidateQuotes = []
    RejectedQuotes = []
    OpenQuotes = []
    diagnostics = []
    found_any = False
    # Find all AllowedMins matches and group those within 75 chars
    min_matches = []
    for min_amt in sorted(AllowedMins, key=lambda x: (-len(x), x)):
        pattern = re.compile(r'(?<!\d)' + re.escape(min_amt) + r'(?!\d)', re.IGNORECASE)
        for m in pattern.finditer(text):
            min_matches.append((m.start(), m.end(), min_amt))
    min_matches.sort()
    # Group matches within 75 chars
    groups = []
    i = 0
    while i < len(min_matches):
        group = [min_matches[i]]
        j = i + 1
        while j < len(min_matches) and min_matches[j][0] - group[-1][1] <= 75:
            group.append(min_matches[j])
            j += 1
        groups.append(group)
        i = j
    for group in groups:
        found_any = True
        start = max(0, group[0][0] - 120)
        end = min(len(text), group[-1][1] + 120)
        # Adjust boundaries to not split words
        while start < group[0][0] and text[start].isalpha():
            start += 1
        while end > group[-1][1] and text[end-1].isalpha():
            end -= 1
        quote_str = text[start:end]
        min_amounts = [m[2] for m in group]
        # aNum logic: string with 4+ digits, 1-2 commas, optional $ at start
        aNum_pat = re.compile(r'\$?\d{1,3}(?:,\d{3}){1,2}(?!\d)')
        aNums = list(aNum_pat.finditer(quote_str))
        # If three consecutive aNum with only whitespace/punctuation between, reject
        rejected_table = False
        if len(aNums) >= 3:
            for i in range(len(aNums)-2):
                between1 = quote_str[aNums[i].end():aNums[i+1].start()]
                between2 = quote_str[aNums[i+1].end():aNums[i+2].start()]
                if not re.search(r'\w', between1) and not re.search(r'\w', between2):
                    rejected_table = True
                    break
        possible_names = set(w.lower() for w in re.findall(r'\b[a-zA-Z]{4,}\b', quote_str))
        diag = {"min_amounts": min_amounts, "group_start": group[0][0], "group_end": group[-1][1], "quote": quote_str}
        if possible_names & names:
            diag["result"] = "rejected_name"
            RejectedQuotes.append({"quote": quote_str, "url": url, "reason": "name"})
            diagnostics.append(diag)
            continue
        if any(bad in quote_str for bad in badSet):
            diag["result"] = "rejected_badSet"
            RejectedQuotes.append({"quote": quote_str, "url": url, "reason": "badSet"})
            diagnostics.append(diag)
            continue
        if rejected_table:
            diag["result"] = "rejected_table_aNum"
            RejectedQuotes.append({"quote": quote_str, "url": url, "reason": "table_aNum"})
            diagnostics.append(diag)
            continue
        creationScore = sum(1 for c in creationSet if c in quote_str)
        fundScore = sum(1 for f in fundSet if f in quote_str)
        minScore = sum(1 for mn in minSet if mn in quote_str)
        sumScore = creationScore + fundScore + minScore
        qdict = {
            "quote": quote_str,
            "url": url,
            "creationScore": creationScore,
            "fundScore": fundScore,
            "minScore": minScore,
            "sumScore": sumScore,
            "minAmounts": min_amounts
        }
        diag["result"] = "accepted"
        diag["creationScore"] = creationScore
        diag["fundScore"] = fundScore
        diag["minScore"] = minScore
        diag["sumScore"] = sumScore
        CandidateQuotes.append(qdict)
        OpenQuotes.append(qdict)
        diagnostics.append(diag)
    if not found_any:
        diagnostics.append({"result": "no_matches_found", "AllowedMins": list(AllowedMins)})
    # Always write diagnostics for Wayne test
    import os
    import json
    diag_path = os.path.join(os.path.dirname(__file__), 'test_outputs/wayne_quote_test/wayne_quote_diag.json')
    try:
        os.makedirs(os.path.dirname(diag_path), exist_ok=True)
        with open(diag_path, "w", encoding="utf-8") as f:
            for d in diagnostics:
                f.write(json.dumps(d, ensure_ascii=False) + "\n")
    except Exception:
        pass
    return CandidateQuotes, RejectedQuotes, OpenQuotes
