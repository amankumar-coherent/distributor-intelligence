# Vendor Intelligence Scraper — Project Summary

## 1. The Task

Given a market name (e.g. "Avocado Oil Market", "High Energy Laser Market"), automatically find every company operating in that market and produce a filled-in profile for each one (role, products, headquarters, size, etc.).

The search-and-collect logic already existed and already worked, but only for **one market at a time**, run separately, with manual cleanup afterward. The goal was to scale that same logic up to run **100 to 10,000 markets in a single unattended batch**, without losing data quality.

## 2. What Has Been Built

**Before**: the search-and-collect part was already automated — the system already searched the web and crawled company websites for a market on its own. What was still manual was everything around that: it only ran **one market at a time** (a separate script per market), and after each run I had to go through and run a stack of one-off cleanup/audit scripts by hand (extra company discovery, removing bad entries, fixing duplicate/aliased names, etc.) to get the results to an acceptable quality bar. There was no way to queue up many markets, no shared cost/rate tracking, and nothing resumed automatically if a run was interrupted.

**Now**, the whole thing has been rebuilt into one automated system:
- **Same automated search-and-collect pipeline per market**, but now: understand the market → search for candidate companies → visit each company's own website → verify it's actually relevant → export a report, with no separate script per market and no manual step needed in between.
- **Batch engine**: point it at a list of markets and it runs many of them at once, keeps track of progress, and automatically resumes from where it left off if interrupted (crash, restart, etc.) — nothing has to be re-run from scratch.
- **The manual cleanup/audit scripts are now built into the pipeline itself**: extra company discovery from trade directories/LinkedIn, and a final review pass, now run automatically as part of every market instead of needing someone to run a script by hand afterward.
- **Cost & progress tracking**: while a batch runs, you can check live status — companies found so far, cost so far, estimated time remaining, and recent failures.
- **Search reliability under load**: if a search engine starts blocking requests, the system automatically slows down or switches to a backup engine, so a large batch doesn't get shut out partway through.
- **Report format adapts by market type**: consumer-goods markets get retail/e-commerce-specific columns; other markets get a general firmographic report — automatically selected per market.

## 3. How It Gets Data, and How Accurate It Is

Data collection happens in two stages:
1. **Discovery** — search engines (and business/trade directories like LinkedIn) are used to find *which companies might belong* in a market.
2. **Verification** — each candidate company's own official website is visited to pull real details (headquarters, products, size, etc.) and confirm it actually operates in that market. An AI model reads the site content and decides relevance + assigns a confidence score — companies aren't included just because the name sounds relevant.

Companies that can't be confirmed (site unreachable, relevance unclear) are listed separately as "not verified" with a reason, rather than being silently dropped or silently included.

**Accuracy status today**: the rebuilt system was checked against 6 markets we already ran and manually audited previously, aiming to match that same quality bar. A full side-by-side accuracy check at real batch scale (running 50, then 500 markets and comparing results against the known-good manual audits) is the next step before relying on it fully unattended at 10,000 markets — this is being validated in stages on purpose, not assumed.

## 4. What Else We're Doing to Make It Better

**How we now get data for 10,000 markets at a time:**
- Build one list of all the markets to run (a simple spreadsheet-style file — market name, country, market type).
- Feed that list to the batch system in one command; it loads every market into a tracked queue.
- The system works through the queue itself, running several markets in parallel, each going through the full search → verify → export process automatically.
- Companies already looked up for one market are reused if they show up again in another market, so repeated work is avoided as the list gets larger.
- Progress, cost, and any failures can be checked at any time while it's running, without stopping it.
- If it's interrupted for any reason, restarting it picks up exactly where it left off — no lost work, no duplicate runs.
- We're rolling this out in stages — first a small batch, then a mid-size batch, then the full 10,000 — checking accuracy and cost at each stage before going bigger, rather than jumping straight to full scale.

**Other improvements planned:**
- Adding optional extra enrichment (e.g. pulling company size/revenue from business databases when a company's own website has thin information).
- Tuning how many markets run at the same time for the best speed without overloading the system.
- Continuing to fold any remaining manual review steps into the automatic pipeline.
