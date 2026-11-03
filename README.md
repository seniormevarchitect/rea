# Provider readiness drafts

Local continuation of the Grok Bot project for a free Provider Readiness Starter Pack and a later Registration Decision Workbook for small Australian providers.

The Grok Bot computer held `starter-pack/claims-tracker-90-day.xlsx` (01:53 on 2 November 2026) and `products/au-registration-workbook/outline-v2.md`. Neither file is in this GitHub repository. This branch rebuilds the tracker from the rules written down that day, and copies the approved page drafts into `drafts/`. Nothing here is published, and nothing takes payment.

## Run the checks

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m unittest discover -s starter-pack/tests -v
.venv/bin/python starter-pack/build_tracker.py
```

The tests cover the three guide dates from the 2 November review note, read as day/month: 2 September 2026, 3 September 2026 and 2 December 2026.

## Still needs you

Prices, the refund policy, Whop, any spend, GST, the seller's legal name, and the brand. The workbook outline was not recovered, so the pre-order stays closed.
