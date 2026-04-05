# WebScraper System: Design and Implementation Plan

## Instructions for Copilot Agent

-   The Copilot Agent must create the code for the web scraper and the
    tests of its functionality.
-   The Agent must run the web scraper and run the tests.
-   The Agent must make any fixes to the code of both until the scraper
    and the tests work properly.
-   If any ambiguity remains, then Copilot Agent is NOT to ask for
    clarification but to make the assumption most likely to lead to
    achieving the final results successfully.
-   The Agent should document any assumptions made during implementation
    if a requirement is not 100% clear.

## Objective

The objective of the WebScraper System is to determine for the
EntityURLs provided what is the minimum dollar amount to endow a named,
targeted fund. I already know for the URLs currently provided what the
correct answer is, but I want to feed the System hundreds of URLs of
other 501c3s whose minimum endowment amount I do not know. I don't
expect the System to perform with 100% accuracy but will be content with
50%. This program is part of my overall project which includes programs
remaining to be finished to download and analyze IRS data about 501c3s
and to communicate with 501c3s about endowment agreements.

## General Conditions

### No Virtual Environment

-   The Agent and code must not create or require a virtual environment.
    All necessary libraries are installed globally.

### Markdown and Formatting

-   All output must use only logical Markdown structure (headers,
    lists), no graphics, font changes, or non-text characters.

### Modularity

-   Each major function (e.g., sitemap parsing, quote extraction,
    scoring, output, testing) must be implemented as a reusable,
    parameterized function.
-   Minimal reliance on global state; data should be passed explicitly
    between functions.
-   All test and production logic must be separated and independently
    testable.

### Test Not Log

-   The Agent will write test routines that run the production code
    against provided test data and determine correctness based on those
    tests. No reliance on log file analysis for debugging.

### Mimic Human Behavior

-    Use a single, modern, hardcoded user-agent for all requests:   session.headers['User-Agent'] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"

-   Introduce a delay between requests to mimic human browsing.

### Error Handling

-   On timeout or failed response, record the error in the Trail of URLs
    and continue.

### Minimize Overhead

-   Large libraries should only be imported once at the system level,
    not within subroutines.

### Naming Conventions

-   501c3 entity: Entity
-   URL of Entity: EntityURL
-   After the URL protocol is the domain name. Since we are dealing with
    US 501c3s, the domain will typically be of the form x.y.z where y is
    the name of the Entity and Extension is either edu, org, or com. We
    refer to the y part as the URLdomainYpart. The y part is removed
    during normalization.

## Data Structure Specifications

-   **CandidateURLs**: list of dicts\
    `{ "``url``": str, "source": "sitemap" | "homepage" | "crawl" }`

-   **RedFlagURLs**: list of dicts\
    `{ "``url``": str, "criterion": str }`

-   **VisitedURLs**: list of dicts\
    `{ "``url``": str, "score": int, "``http_status``": int }`

-   **OpenURLs**: priority queue of tuples\
    `( -score, ``len``(``url``), ``url`` )`

-   **CandidateQuotes**: list of dicts

-   **OpenQuotes**: list of dicts\
    `{ "quote": str, "``url``": str, "``creationScore``": int, "``fundScore``": int, "``minScore``": int, "``sumScore``": int }`

-   **RejectedQuotes**: list of dicts\
    `{ "quote": str, "``url``": str, "reason": str }`

-   **BestQuotes**: list of dicts\
    same structure as OpenQuotes

-- PathSentence: string

## URL Normalization Process

For RedFlag tests and URL scoring, the URL is processed as follows:

1\. Remove the protocol (e.g., "https://") and the "://".

2\. Remove the y part of the domain (the entity-identifying part, which
is the second label in a standard x.y.z domain). - For
"giving.wayne.edu", remove "wayne" → "giving.edu"

3\. The remaining string, including subdomains and path, is processed so
that each non-alphabetic character (such as backslash, period, dash,
etc.) is replaced with a blank space.

4\. Collapse multiple spaces into a single space and trim
leading/trailing spaces.

5\. The result is called `pathSentence`.

**Example:**

-   Input: `https://giving.wayne.edu/endowments/minimum`\
    Step 1: Remove protocol → `giving.wayne.edu/endowments/minimum`\
    Step 2: Remove y part → `giving.edu/endowments/minimum`\
    Step 3: Replace non-alphabetic →
    `giving ``edu`` endowments minimum`\
    Output: `giving ``edu`` endowments minimum`

## User-Defined Constants

### Preamble

In a set of strings, any two elements of the set are delimited by a
comma (,). Thus { events, financial aid, history} has 3 elements. In
subsequent instructions of the sort PathSentence contains any URL_Bad
string → RedFlag, the meaning is that an element in URL_Bad is a
substring of any string in PathSentence. Thus, for PathSentence to cause
a URL to be moved to RedFlag requires that PathSentence contains
\"financial aid\" and not just \"financial\".

### Constants

-   ThrottleMax = 30
-   AllowedURLfileExtension = { xml, .txt, .html, .htm }
-   URL_Bad = { access, blog, blogs, calendar, catalog, campus, career, child, course, directory, disability, event, events, financial aid, history, loan, login, magazine, media, news, podcast, story, stories }
-   URL_Good = { about, advance, agree, alumn, award, charit, creat, donat, donor, endow, establish, faq, financ, fund, gift, give, giving, help, legacy, recognition, scholar, support }
-   AllowedMins = { $5,000, $10,000, $15,000, $20,000, $30,000, $40,000, $50,000, $75,000, $100,000, $120,000, $125,000, $150,000, $170,000, $175,000, $200,000, $220,000, $250,000, $300,000 }
-   creationSet = { agree, award, charit, create, donor, endow, establish, fellow, fund, give, giving, gift, legacy, paid }
-   fundSet = { fellowship, field of interest, scholar }
-   minSet = { minim, or more, at least, begin, start, initial }
-   badSet = { financial aid, loan, maximum, tier }

## Inputs

-   Read list of URLs from input/URLs.txt (one URL per line).
-   Read names.txt (one name per line, alphabetically sorted) for name
    detection.

## Outer Loop

For each EntityURL:

Set these data structures to empty:

-   AllSeenURLs
-   CandidateURLs
-   OpenURLs
-   RedFlagURLs
-   VisitedURLs
-   CandidateQuotes
-   RejectedQuotes
-   BestQuotes

## DeDuplication Rule

Introduce a global set (e.g., \`AllSeenURLs\`) to track all normalized
URLs added to any list.

Before adding a URL to any list, check if its normalized form is in \`
AllSeenURLs \`.

\- If yes, skip adding it to any list (including RedFlagURLs).

\- If no, add it to the target list and record it in \`AllSeenURLs\`.

## Collect First Set of URLs

### Sitemap Parsing Rules

-   If the EntityURL provides a sitemap index, follow links to child
    sitemaps **one level deep only**.
-   Do **not** follow nested sitemap indexes beyond this single level.
-   From each child sitemap, extract URLs directly (no further
    recursion).
-   Stop after collecting **up to 3,000 unique URLs**.

### Homepage URL Extraction Rules

If no sitemap exists:

-   Extract only `<a ``href``="...">` links.
-   Resolve relative URLs using the homepage as base.

## RedFlag Filtering

### First 4 Red Flags

For each CandidateURL (call it X), apply criteria in order:

1.  **Protocol not http/https** → RedFlag
2.  **Path contains #, ?, or digit** → RedFlag
3.  **File extension not in AllowedURLfileExtension**
4.  **Domain mismatch** (y-part does not equal y part of EntityURL)

### Normalize and Continue RedFlag

Further RedFlag operates on Normalized URL: Normalize URL and call
resultant URL PathSentence

#### URL_Bad 

PathSentence contains any URL_Bad string → RedFlag

#### Personal Name 

PathSentence contains a personal name of more than 3 characters in
length -\> RedFlag

-   Name match uses regex:\
    `(^|[^a-z])NAME([^a-z]|$)`

### RedFlag Enforcement

-   All RedFlag criteria (e.g., presence of ?, #, digits, %, or other
    forbidden characters/conditions) must be enforced before any URL is
    visited or added to VisitedURLs. If a URL fails any RedFlag test, it
    must not be visited or added to VisitedURLs, but instead added to
    RedFlagURLs with the appropriate criterion.
-   This logic must be applied consistently for URLs from sitemaps,
    homepage, and crawl discoveries.

## Move to OpenURLs

For each URL remaining on CandidateURLS after RedFlag screening, move
those URLs along with their PathSentence to OpenURLs.

## Heuristic Scoring

For each (URL, PathSentence) upon its addition to OpenURL it is assigned
a Score

\- Score +=1 for each string in URL_Good found in PathSentence.

\- Store score with URL.

## Visit URLs

While OpenURLs not empty and `len``(``VisitedURLs``) < ``ThrottleMax`:

Take highest scoring URL from OpenURL and move to VisitedURL. If more
than one URL has highest score, then tie-breakers:

1.  Shorter URL length
2.  Lexicographic order

Call this URL to now be processed the StudyURL

-   Visit StudyURL; record http status and score.
-   If page loads:
    -   Extract URLs from StudyURL (with same procedure as used for
        extracting URLs from a home page) and add these newly found URLs
        to CandidateURLs where they will be processed with RedFlag and
        other routines as where the original URLs from sitemaps or the
        home page.

    -   Pass StudyURL to QuoteExtract.

## Quote Extraction

Extract visible text (strip HTML/JS) StudyURL to create VisibleTextPage

### Find MinInstance and Create QuoteString

The program parses VisibleTextPage to find each occurrence (call it
MinInstance) of a string from the set AllowedMins

For each MinInstance, extract the 120 characters (including spaces) to
the left of MinInstance and the 120 characters to the right of
MinInstance (including spaces). If the left-most (or right-most) cutoff
splits a word (with word defined as alphabetic characters bounded by
either spaces or punctuation), then delete that word fragment. Store the
resultant string with MinInstance inside it, as QuoteString in
CandidateQuotes.

### Post-Processing of QuoteStrings for Multiple Minimums in Proximity

-   When extracting quotes, if two or more AllowedMins (minimum dollar
    amounts) are found within 75 characters of each other, treat them as
    part of a single quote window.
-   The extraction window should expand to include all such nearby
    AllowedMins, resulting in one combined QuoteString that covers the
    full span from the leftmost to the rightmost minimum in the group.
-   Only one quote should be created for this group, with all relevant
    minimum amounts included (e.g., as a list or as found in the text).
-   Scoring and filtering should be applied to this combined quote as
    usual.

**Example:** If "\$10,000" and "\$20,000" are found within 75
characters, extract a single quote window covering both, not two
overlapping quotes.

## Moving QuotesStrings from Candidates to Open or Rejected

### identify proper names

-   identify in QuoteString each string of alphabetic characters of
    length greater than 3 bounded by spaces and put each such string in
    a set call possibleNames

-   if any element in possibleNames is in names.txt, then move
    QuoteString from CandidateQuotes to RejectedQuotes

###  badset

if QuoteString contains an element from badSet, then move QuoteString
from CandidateQuotes to RejectedQuotes

### Numbers Near Each Other

`Define ``aNum`` as a string of characters bounded by spaces which ``has`` `

-   4 or more digits

-   at least one comma but not more than two which demarcate thousands,
    as in 100,000

-   optionally begins with \$

If QuoteString contains three consecutive aNum without any intervening
non-empty words, then move QuoteString from CandidateQuotes to
RejectedQuotes.

-   A quote should only be rejected for the "table_aNum" reason if it
    contains **three** consecutive aNum (number patterns, e.g.,
    \$10,000, \$20,000, \$25,000) without any intervening non-empty
    words.
-   If there are only two consecutive aNum, the quote should NOT be
    rejected for this reason.
-   This change ensures that valid quotes containing two minimum amounts
    in close proximity are not mistakenly rejected as table data.

**Example:**\
A quote like "...with a \$2,000 gift (\$5,000 for scholarship) with up
to five years to build to a fund minimum of \$10,000 (\$25,000 for
scholarship funds)..." should NOT be rejected for table_aNum, since
there are only two consecutive aNum at a time.

### Scoring: From CandidateQuotes to OpenQuotes

After a QuoteString has passed all RedFlag criteria and not been moved
to RejectedQuotes, Score it as follows

## Scoring QuoteStrings

For each remaining QuoteString on CandidateQuotes

Set creationScore = 0, fundScore = 0, minimumScore = 0

(In the following, the meaning of the expression "for each string in
CreationSet, if in QuoteString" means that if an element in CreationSet
is a substring of any string in QuoteString, then the condition is
satisfied. By way of illustration, the string 'charit' in creationSet
satisfies the string 'charitable' in QuoteString)

For each string in creationSet, if in QuoteString, then creationScore =
creationScore +1

For each string in fundSet, if in QuoteString, then fundScore =
fundScore +1

For each string in minSet, if in QuoteString, then minScore = minScore
+1

Create sumScore = creationScore + fundtypeScore + minScore

Move this QuoteString from CandidateQuotes (along with its scores,
element of AllowedMins, and URL) to OpenQuotes.

## Termination

Stop when:

-   `VisitedURLs`` >= ``ThrottleMax`, or
-   `OpenURLs` is empty.

## BestQuotes Selection and Output Format

### Best Five

Review all entries in OpenQuotes

-   If ≤5 entries: keep all.

-   If \>5:

    1.  Remove entries with minScore = 0.
        -   If this removes all quotes, restore original set.
    2.  Sort by sumScore descending.
    3.  Keep top 5 entries and output as BestQuotes with all its
        associated information from OpenQuotes. 

### BestQuotes.md Format

BestQuotes.md should be formatted with hierarchy and bullets as follows:

For each of the 5 best quotes:

        -   header level 2: Best Quote #k (where k goes from 1 to 5)

        -   header level 3: : URL on which quote was found

        -   header level 4: minimal dollar amount

        -   bullet item: Quote verbatim

        -   bullet item: creationScore

        -   bullet item: fundScore

        -   bullet item: minScore

        -   bullet item: sumScore

## Output for EntityURL

For each EntityURL, create a subdirectory named after the Entity and
produce **two versions** of each output file:

### JSON version

-   One JSON object per line.
-   UTF‑8 encoded.

### Markdown version

-   Plain Markdown tables or lists.
-   No graphics or emojis.

Files:

-   CandidateURLs.json / CandidateURLs.md
-   RedFlagURLs.json / RedFlagURLs.md
-   VisitedURLs.json / VisitedURLs.md
-   CandidateQuotes.json / CandidateQuotes.md
-   RejectedQuotes.json / RejectedQuotes.md
-   OpenQuotes.json / OpenQuotes.md
-   BestQuotes.json / BestQuotes.md

Ordering must follow discovery order.

## Testing

### Syntax and Execution

-   Remove all syntax/execution errors by running system against
    input/URLs.txt until error-free.

### Functional Testing

-   Use testURLs.txt as gold standard: EntityURL, target URL, minimum
    amount, desired quote.

### Subsidiary Tests

-   Sitemap overlap: compare CandidateURLs to testing/sitemapURLs/\*.txt

-   Homepage URLs: compare homepage scraper output to YaleURLs.txt

-   Scoring of URLs: confirm for URL https://giving.wayne.edu/about/faq
    that thePathSentence score is 3.

-   Quote Extraction: compare extracted quotes to
    testing/QuoteExtraction/\*.csv

-   Wayne Quote Extract Page test:

    -   confirm that what the program produces as VisibleTextPage from
        https://giving.wayne.edu/about/faq is the same as the text at
        testing/QuoteExtraction/wayne.txt

    -   confirm that the program's two extracted quotes and their scores
        for that VisibleTextPage are the same as wayneOpenQuotes.csv
        whose first row is labels.

    For catlin, pacfwv, swe, and yale, please get the precise target URL
    testing/testURLs.txt. Then confirm that the text that scraper.py
    extracts matches the relevant text file in testing/QuoteExtraction.
    Then check that the quotes and minimum dollar amounts in the
    relevant csv file match those produced by scraper.py. The columns in
    the catlin, pacfwv, swe, and yale csv files have not been updated to
    correspond to the latest changes in requirements and thus you should
    not expect an exact match between those 4 csv files and the scoring
    of scraper.py.

## Error Reporting Format

    [ERROR] <timestamp> <EntityURL> <error_type> <error_message>

Example:

    [ERROR] 2026-03-30T22:30:00Z https://example.org/endowment-page TimeoutError Failed to load within 10 seconds
