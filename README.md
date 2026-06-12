# Pricing AI Strategist

Local prototype for the CPG price discipline cockpit.

## How to run

1. Open `index.html` in Microsoft Edge or Chrome.
2. Select the Stackline first-mover `.xlsx` workbook.
3. Use the cockpit, recommendation list, and strategist panel.

You can also double-click `start-pricing-ai-strategist.bat` on Windows.

## What it analyzes

- Product map
- Daily price sheets
- Daily price-per-count sheets
- Today's rollup
- Guardrail flags against `NonPromo MAP`
- Consecutive days below guardrail
- Lowest-price retailer
- Price-per-count gaps
- OOS flags
- Retailer/channel discipline score

## UX views

- Cockpit: scan priority guardrail flags, inspect evidence, and move a recommendation into review.
- Recommendations: review one recommendation at a time with Summary, Evidence, Impact, Risks, and Approval tabs.
- SKU Explorer: inspect SKU/retailer rows, daily price history, price-per-count context, and governance notes.
- Ask Strategist: ask workbook-grounded questions using local answers.
- Approvals: track recommendations across Draft, RGM, Finance, Sales, Legal, Approved, Rejected, and Revised.

## Governance

The prototype does not execute prices or generate retailer outreach language.
Every recommendation requires RGM approval before action. Legal review is flagged
when a recommendation references guardrails, corridor language, repeated retailer
issues, channel conflict, or external-facing commercial action.

## Notes for testers

This first version is intentionally built around the Stackline first-mover
workbook structure. Margin, revenue, units, trade spend, and promo ROI are not
calculated unless those fields are added in a future data feed.
