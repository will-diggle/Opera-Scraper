# Opera Job Scraper

Four stages. Stage 4 is the one you run every week.

Checks opera company websites for auditions, chorus vacancies, young artist
programmes and other openings, and puts everything into a filterable spreadsheet.

## Running it

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python scraper.py
```

222 companies across the UK, Ireland, Europe, North America and Australasia —
takes roughly 8–10 minutes. To test one company:

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python scraper.py glyndebourne
```

## What you get

- **opera_jobs.xlsx** — open in Excel/Numbers. Filters are already switched on
  and the header row is frozen.
  - Sheet **Opportunities**: one row per thing found. Columns: Company,
    Country, Role / posting, Type, Deadline found, Link, Found on page,
    Date checked, plus empty **Status** and **Notes** columns for you to fill in.
  - Sheet **Coverage**: every company and whether the scraper managed to read
    its site. Check this to see what was missed.
- **opera_jobs.csv** — same data, for importing into Google Sheets or Notion.

The **Type** column is a rough guess:
- `Singer` — mentions a voice type, chorus, audition, studio, etc.
- `Other staff` — clearly admin/technical.
- `Unclear` — worth a glance.

Filter on `Type = Singer` first, then skim `Unclear`.

## Adding companies

Open `companies.csv` and add a row: `name,country,website`. Just the homepage —
the scraper finds the jobs/auditions pages itself. Rerun and it's included.

## Known limits

- **Overwrites the output files each run.** If you've typed notes into the
  spreadsheet, rename it (e.g. `opera_jobs_working.xlsx`) before rerunning.
- Some sites block automated visits (HTTP 403 / connection refused) and always
  need checking by hand: Lyric Opera of Chicago, Washington National Opera,
  Wolf Trap Opera, Staatstheater Karlsruhe, Pegasus Opera. The Coverage sheet
  flags these every run.
- Some companies publish an **archive** of past auditions, so old postings
  (e.g. 2022 dates) appear alongside current ones. Check the date on the page.
- Sites that build their page with JavaScript may show nothing. Also flagged in
  Coverage as "no jobs/auditions link found".
- There is still some noise (cast lists, page headings). Deliberate — better a
  few extra rows you can filter out than a missed audition.
- "Deadline found" is a best guess from the page text. Always confirm on the
  actual page before relying on it.


---

# The Operabase pipeline

`companies.csv` was my hand-made list of 222 big houses. This pipeline instead
builds the list automatically from Operabase, which covers small companies,
festivals and ensembles too - about 84,000 organisations worldwide.

Operabase's own rules (`robots.txt`) allow this, and they publish machine-readable
sitemaps for exactly this purpose. We use those rather than scraping their pages
by force. Please keep it that way: don't raise the speed settings.

### Stage 1 - get the master list

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python stage1_list.py
```

17 requests, under a minute. Writes `operabase_orgs.csv` (~84,000 rows).

### Stage 2 - get each one's website and social accounts

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python stage2_details.py
```

The slow one: roughly 90 minutes for all 84,000. **You can stop it any time with
Ctrl-C and run it again later** - it remembers everything already fetched and
carries on. Writes `operabase_details.csv`: name, type, city, country, official
website, Facebook, Instagram, X, YouTube, LinkedIn, phone, email.

Only needs rerunning every few months, to pick up new companies.

### Stage 3 - choose who to check

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python stage3_filter.py GB IE DE AT CH
```

No countries given = UK/Ireland + all Europe. `ALL` = worldwide. Drops anything
with no website and de-duplicates companies sharing a site. Writes `targets.csv`.

### Stage 4 - look for jobs

```bash
cd "/Users/willdiggle/Opera Scraper" && .venv/bin/python scraper.py --file targets.csv
```

Same scraper as before, pointed at the bigger list. Writes `jobs_targets.xlsx`.
Each row carries that company's social media links in a **Social accounts**
column.

## About social media

I did not build a social media scraper, and I'd advise against one. Facebook,
Instagram, X and LinkedIn all block automated reading and require a logged-in
account; scraping them risks your personal account being restricted, and any
scraper built against them breaks within weeks.

What the pipeline does instead: it collects each company's social links from
Operabase and puts them in the spreadsheet, so you can click straight through to
the feeds of the handful of companies you actually care about. That is the part
worth doing by hand.
