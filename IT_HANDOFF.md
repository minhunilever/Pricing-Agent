# Pricing AI Strategist - Azure Static Hosting Handoff

## Purpose

This package contains a static browser app for CPG price discipline review. It reads a Stackline first-mover Excel workbook in the user's browser and produces cockpit views, guardrail flags, and RGM recommendation drafts.

## Data Handling

- No backend is included.
- No database is included.
- No workbook data is uploaded by this prototype.
- Users select an `.xlsx` workbook from their machine.
- Workbook parsing happens client-side in Edge or Chrome.
- Approval status is stored only in the user's browser storage.

## Recommended Hosting

Use Azure Static Web Apps or an equivalent internal static web host.

Recommended URL pattern:

`https://pricing-ai.<company-domain>`

## Files To Host

Host the contents of this folder as the web root:

- `index.html`
- `assets/app.js`
- `assets/styles.css`
- `staticwebapp.config.json`

The entry file is `index.html`.

## Access Control

The included `staticwebapp.config.json` requires an authenticated user and redirects unauthenticated users to Azure AD / Entra ID login:

```json
"allowedRoles": ["authenticated"]
```

If IT wants a short public/unrestricted smoke test first, replace `staticwebapp.config.json` with `staticwebapp.public-test.config.json`. Do not use the public config for broad company testing with sensitive workbooks.

## Entra ID Setup

For company access, configure Azure Static Web Apps authentication with Microsoft Entra ID.

Recommended policy:

- Require company tenant login.
- Restrict access to the intended test group if possible.
- Do not allow anonymous access for production-style testing.
- Use HTTPS only.

## Domain Setup

To use a company domain:

1. Create the Azure Static Web App.
2. Deploy this static folder.
3. Add the custom domain in Azure Static Web Apps.
4. Ask DNS/IT to create the required DNS record for the chosen hostname.
5. Wait for Azure domain validation and managed TLS certificate provisioning.
6. Confirm the app is reachable at the company hostname.

Suggested hostname:

`pricing-ai.<company-domain>`

## Test Plan

Use Microsoft Edge or Chrome.

1. Open the hosted app URL.
2. Confirm unauthenticated users are redirected to company login.
3. Sign in as an approved tester.
4. Upload the Stackline workbook.
5. Confirm these views load:
   - Cockpit
   - Recommendations
   - Ask Strategist
   - Data Map
6. Export the executive brief and recommendation CSV.
7. Confirm no workbook data appears in server logs or app network requests.

## Governance Notes

The app is decision support only.

- It does not execute prices.
- It does not draft retailer outreach language.
- `NonPromo MAP` is treated as a commercial guardrail.
- The app does not label rows as legal violations.
- Every recommendation requires RGM approval.
- Legal review is flagged only when a recommendation could lead to external action, guardrail language, channel conflict, or retailer-specific follow-up.

## Known MVP Limits

This package analyzes only the first-mover workbook structure. It does not calculate:

- Units
- Revenue
- Margin
- COGS
- Trade spend
- Promo ROI
- Repeat rate

Volume and margin language in the app is intentionally labeled as hypothesis-only.
