# AI-Powered Success Plan Generator

Turns raw account data into an executive-ready Customer Success Plan, structured around TSIA's six success-plan components.

## Why this exists

Most CS teams write success plans by hand, one account at a time, in whatever format the CSM prefers. That means inconsistent quality, hours of manual slide-building, and no guarantee a plan actually covers what TSIA recommends. This project generates a consistent, executive-facing success plan deck straight from account data, so every plan looks and reads the same regardless of who's running the account.

## What it does

Given structured account data (stakeholders, usage, milestones, metrics, actions, risks), the generator produces a polished slide deck covering all six TSIA success-plan components:

1. **Customer Profile** — stakeholders, business goals, key use cases
2. **Current State Assessment** — product usage, adoption levels, known challenges
3. **Milestones & Timeline** — journey stages with target dates and status
4. **Success Metrics** — adoption, onboarding, and business-impact metrics vs. target
5. **Actions & Responsibilities** — what the CS team owns vs. what the customer owns
6. **Risk Management** — open risks, likelihood/impact, mitigation plans

## How it's built

The generator is deliberately split into two layers so it can plug into a real Customer Success Platform later without touching the rendering code:

```
CSP data (or sample .xlsx)
        │
        ▼
generator_step1_extract.py   ← data layer: reads account data, normalizes it into a JSON payload
        │
        ▼
generator_step2_build_deck.js ← rendering layer: turns the JSON payload into a .pptx deck
        │
        ▼
Executive success plan (.pptx)
```

To connect a live CSP (Gainsight, Totango, ChurnZero, etc.), only `generator_step1_extract.py` changes: swap the workbook reads for API calls that return the same field names. The rendering script never needs to know where the data came from.

## Sample data

`Success Plan Sample Customer Data.xlsx` contains three synthetic accounts, normalized into tables that mirror a typical CSP export:

| Tab | Contents | Maps to |
|---|---|---|
| Accounts | Segment, ARR, health score, renewal date, CSM owner | Account / Company object |
| Stakeholders | Contacts with role and sentiment | Contacts / Relationship map |
| Goals & Use Cases | Business goals and use cases per account | Success Plan / Objectives object |
| Usage & Adoption | Feature-level adoption and known challenges | Product usage / telemetry feed |
| Milestones | Journey stages, target dates, status | Success Plan / Timeline / CTAs |
| Success Metrics | Adoption, onboarding, business-impact metrics vs. target | Scorecard / Outcomes object |
| Actions | Owned actions (CS team vs. customer) | CTA (Call to Action) / Task object |
| Risks | Open risks, likelihood/impact, mitigation | Risk CTA / Health score drivers |

No real customer data is used anywhere in this repo.

## Sample output

Three generated decks are included, covering a healthy account, an at-risk account, and a strong account, to show the generator holds up across different data shapes:

- `Success Plan - Meridian Health Systems.pptx` (healthy, Enterprise)
- `Success Plan - Brightline Logistics.pptx` (at-risk, Mid-Market)
- `Success Plan - Foundry Analytics.pptx` (strong, Enterprise)

## Tech stack

- **Python** (openpyxl, pandas) — data extraction and normalization
- **Node.js** (pptxgenjs) — slide deck rendering
- **Excel** as the interim data source, structured to match common CSP export schemas

## Running it yourself

```bash
python generator_step1_extract.py ACC-1001      # writes payload_ACC-1001.json
node generator_step2_build_deck.js ACC-1001     # writes success_plan_ACC-1001.pptx
```

## Known limitations / next iteration

- Progress bars on the Success Metrics slide assume "higher is better." A metric like time-to-first-value, where lower is better, needs an inverted calculation, which is the next planned fix.
- Currently generates one deck per account on demand. A natural next step is batch generation across an entire book of business, paired with the CSM Workload & Book-of-Business Balancer project.

## Part of a larger portfolio

This project is one piece of a connected AI-powered Customer Success automation portfolio. See the [profile README](https://github.com/melissamcgowan) for the full roadmap.
