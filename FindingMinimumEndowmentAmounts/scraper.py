# WebScraper System Main Module

"""
Implements the web scraper system as specified in plan/Plan.md.
- Modular, parameterized functions for each major task
- No virtual environment required; assumes all libraries are installed globally
- All output in Markdown or JSON as specified
- Test and production logic separated
- Mimics human browsing (user-agent, delay)
"""

import os
import re
import time
import json
import random
import requests
from urllib.parse import urlparse, urljoin
from collections import deque

# --- User-Defined Constants ---
ThrottleMax = 30
AllowedURLfileExtension = {'.xml', '.txt', '.html', '.htm'}
URL_Bad = {'access', 'blog', 'blogs', 'calendar', 'catalog', 'campus', 'career', 'child', 'course', 'directory', 'disability', 'event', 'events', 'history', 'loan', 'login', 'magazine', 'media', 'news', 'podcast', 'story', 'stories'}
URL_Good = {'about', 'advance', 'agree', 'alumn', 'award', 'charit', 'creat', 'donat', 'donor', 'endow', 'establish', 'faq', 'financ', 'fund', 'gift', 'give', 'giving', 'help', 'legacy', 'recognition', 'scholar', 'support'}
AllowedMins = {'$5,000', '$10,000', '$15,000', '$20,000', '$30,000', '$40,000', '$50,000', '$75,000', '$100,000', '$120,000', '$125,000', '$150,000', '$170,000', '$175,000', '$200,000', '$220,000', '$250,000', '$300,000'}
creationSet = {'agree', 'award', 'charit', 'create', 'donor', 'endow', 'establish', 'fellow', 'fund', 'give', 'giving', 'gift', 'legacy'}
fundSet = {'fellowship', 'field of interest', 'scholar'}
minSet = {'minim', 'or more', 'at least', 'begin', 'start', 'initial'}
badSet = {'loan', 'maximum', 'tier'}

# --- Utility Functions ---
def normalize_url(url):
    url = url.lower()
    # Remove protocol
    if '://' in url:
        url = url.split('://', 1)[1]
    # Split domain and path
    if '/' in url:
        domain, path = url.split('/', 1)
        path = '/' + path
    else:
        domain = url
        path = ''
    # Remove y part of x.y.z domain (second label)
    domain_parts = domain.split('.')
    if len(domain_parts) >= 3:
        # Remove the second label (y part)
        domain = '.'.join([domain_parts[0]] + domain_parts[2:])
    # Recombine domain and path
    norm = domain + path
    # Replace non-alpha chars with space
    norm = re.sub(r'[^a-z]', ' ', norm)
    norm = re.sub(r' +', ' ', norm).strip()
    return norm

def load_names(names_path):
    with open(names_path, encoding='utf-8') as f:
        # Only include names > 3 characters
        return [line.strip().lower() for line in f if line.strip() and len(line.strip()) > 3]


# Optimized for URL path: tokenize and check set intersection
def contains_name(text, names, use_regex=False):
    if use_regex:
        # For quote extraction, use regex for context
        for name in names:
            if re.search(r'(^|[^a-z])' + re.escape(name) + r'([^a-z]|$)', text, re.I):
                return True
        return False
    # For URL path, tokenize and check
    tokens = set(text.split())
    names_set = set(names)
    return not tokens.isdisjoint(names_set)


def delay():
    time.sleep(random.uniform(1.5, 3.5))

# --- Data Structure Creators ---
def make_candidate_url(url, source):
    return {"url": url, "source": source}

def make_redflag_url(url, criterion):
    return {"url": url, "criterion": criterion}

def make_visited_url(url, score, http_status):
    return {"url": url, "score": score, "http_status": http_status}

def make_candidate_quote(quote, url, creationScore, fundScore, minScore, sumScore):
    return {"quote": quote, "url": url, "creationScore": creationScore, "fundScore": fundScore, "minScore": minScore, "sumScore": sumScore}

def make_rejected_quote(quote, url, reason):
    return {"quote": quote, "url": url, "reason": reason}

# --- Main Scraper Logic ---
# (To be continued in next file due to size)
