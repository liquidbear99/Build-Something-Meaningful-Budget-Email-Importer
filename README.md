# Build Something Meaningful: Budget Email Importer

A Google Apps Script budget helper that imports credit card transaction alert emails from Gmail, categorizes them with merchant rules and optional OpenWebUI AI, and gives you a mobile-friendly review screen before transactions are added to your budget sheet.

This project is meant to be copied into **your own Google Sheet** so every user owns their own data, rules, review queue, and configuration.

## What it does

- Searches Gmail for supported credit card transaction alerts.
- Parses transaction date, merchant, amount, account, card ending, source, and Gmail message ID.
- Skips duplicates using a generated duplicate key.
- Applies merchant rules for confident auto-categorization.
- Supports ignored merchant rules for transactions you do not want in the budget.
- Uses optional OpenWebUI AI categorization when merchant rules are uncertain.
- Sends uncertain items to a Pending Transactions sheet.
- Provides a fast card-style review web app for approving, learning rules, skipping, or ignoring merchants.
- Can optionally submit approved transactions to a Google Form.
- Can optionally email you when transactions need review.

## Supported email sources

The current parser is designed for:

- Chase transaction alert emails.
- American Express purchase alert emails.

Statement summaries, balance notices, payment reminders, weekly snapshots, and other non-purchase emails are intentionally excluded.

## Categories

The starter budget categories are:

- Groceries
- Restaurants
- Misc
- My Fun
- Spouse Fun
- Pet
- Hair/Cosmetics
- Home/Car Maintenance
- Entertainment/Clothing
- Medical
- Fuel
- Utilities
- Spouse Benefits
- My Benefits
- Vacation

You can rename these, but if you do, update all of these places together:

1. `CONFIG.ALLOWED_CATEGORIES`
2. `CONFIG.FORM_CATEGORY_MAP`
3. Any existing Merchant Rules categories
4. Any Google Form category choices, if Google Form submission is enabled

## Files in this repository

| File | Purpose |
| --- | --- |
| `Budget Email Importer Code.js` | Main Google Apps Script code. Paste this into `Code.gs` in Apps Script. |
| `Budget Email Importer ReviewApp.html` | Web app UI. Add this as an Apps Script HTML file named `ReviewApp`. |
| `README.md` | Setup and usage guide. |
| `LICENSE` | Apache-2.0 license. |

## Important privacy note

Do **not** paste another person's spreadsheet URL, Google Form URL, API key, or web app URL into your copy.

This app should be installed from inside the Google Sheet that will hold your budget data:

```text
Google Sheet → Extensions → Apps Script
```

Do not create it as a standalone Apps Script project unless you know how to bind it to a spreadsheet. The included code has a safety check that should stop setup if it cannot find an active spreadsheet.

## Quick start

### 1. Create your budget spreadsheet

Create a new blank Google Sheet. This will become your private budget importer workbook.

### 2. Open Apps Script

From that sheet, go to:

```text
Extensions → Apps Script
```

### 3. Add the main script

Open the default script file, usually named `Code.gs`, delete its contents, and paste in the contents of:

```text
Budget Email Importer Code.js
```

Save the project.

### 4. Add the review web app HTML

In Apps Script:

```text
+ → HTML
```

Name the file exactly:

```text
ReviewApp
```

Paste in the contents of:

```text
Budget Email Importer ReviewApp.html
```

Save again.

### 5. Run setup

In Apps Script, select and run:

```text
setupBudgetAutomation
```

Approve the requested permissions. The script needs access to the spreadsheet, Gmail labels/search, optional email notifications, optional forms, and optional external OpenWebUI calls depending on which features you enable.

Setup creates these sheets:

- Expenses
- Pending Transactions
- Merchant Rules
- Ignored Merchants
- Review Settings
- Import Log

It also creates these Gmail labels:

- `Budget/Processed`
- `Budget/Needs Review`

### 6. Test the parsers

From the Apps Script function selector, run:

```text
testParseExample
```

Then run:

```text
testParseAmexExample
```

These do not import real emails. They only confirm that the sample parser logic works.

### 7. Import transactions manually

Run:

```text
importGmailTransactions
```

The script searches recent Gmail messages for supported transaction alerts, applies rules, and either adds confident transactions to Expenses or sends uncertain transactions to Pending Transactions.

### 8. Deploy the review app

In Apps Script, go to:

```text
Deploy → New deployment → Web app
```

Recommended starting settings:

- Execute as: `Me`
- Who has access: `Only myself`

After deployment, copy the web app URL. If you want review notification emails to include the link, paste it into:

```js
WEB_APP_URL: "PASTE_YOUR_DEPLOYED_WEB_APP_URL_HERE"
```

Then create a new deployment version so the live web app uses the updated code.

### 9. Review pending transactions

Open the deployed web app. Each pending transaction appears as a card with quick actions:

- **Approve**: approve using the selected category.
- **Approve & Learn**: approve and create a future merchant rule.
- **Use AI**: approve the AI suggestion.
- **Use AI + Learn**: approve the AI suggestion and create a future merchant rule.
- **Ignore Merchant**: skip this merchant going forward.
- **Skip**: remove only this transaction from the review queue.

### 10. Optional: create an hourly trigger

After confirming manual imports work, run:

```text
createHourlyImportTrigger
```

This creates an hourly Apps Script trigger for `importGmailTransactions`.

To remove it later, run:

```text
deleteBudgetTriggers
```

## Configuration

Most setup options live near the top of `Code.gs` inside `CONFIG`.

| Setting | Default | What it does |
| --- | --- | --- |
| `SEARCH_DAYS_BACK` | `7` | How many days back Gmail import should search. |
| `MAX_THREADS_PER_RUN` | `50` | Maximum Gmail threads checked per import run. |
| `OPENWEBUI_API_URL` | placeholder | Optional OpenWebUI chat completions endpoint. |
| `OPENWEBUI_API_KEY` | placeholder | Optional OpenWebUI API key. Do not commit a real key. |
| `OPENWEBUI_MODEL` | `AUTO` | Auto-selects a usable model, or you can set an exact model ID. |
| `PUSH_TO_GOOGLE_FORM` | `false` | If true, approved expenses are submitted to a Google Form. |
| `GOOGLE_FORM_URL` | placeholder | Google Form edit URL, only needed if form push is enabled. |
| `SEND_REVIEW_EMAILS` | `false` | If true, sends email when transactions need review. |
| `NOTIFICATION_EMAIL` | blank | Email address for review notifications. |
| `WEB_APP_URL` | blank | Deployed review app URL included in notification emails. |
| `DEFAULT_RULE_CONFIDENCE` | `0.95` | Confidence assigned to learned merchant rules. |
| `DEFAULT_AUTO_APPROVE_UNCERTAINTY` | `0.10` | Auto-approval requires roughly 90% confidence or higher. |

## Optional OpenWebUI AI categorization

AI categorization is optional. If you do not want AI categorization, leave the OpenWebUI placeholders alone or turn off `USE_AI_CATEGORIZER` in the Review Settings sheet.

To enable OpenWebUI:

1. Set `OPENWEBUI_API_URL` to your OpenWebUI chat completions endpoint.
2. Set `OPENWEBUI_API_KEY` to your API key.
3. Leave `OPENWEBUI_MODEL` as `AUTO`, or set it to an exact model ID from your OpenWebUI `/api/models` response.
4. Run `testOpenWebUiCategorizerFromAppsScript`.

Never commit your real API key to GitHub.

## Optional Google Form submission

Google Form submission is off by default.

To enable it:

1. Create a Google Form.
2. Make sure it has these exact item titles:
   - `What is Expense (Name)`
   - `Amount ($)`
   - `Expense Category`
3. Make sure the `Expense Category` multiple choice options match the values in `CONFIG.FORM_CATEGORY_MAP`.
4. Paste the Google Form edit URL into `CONFIG.GOOGLE_FORM_URL`.
5. Set `PUSH_TO_GOOGLE_FORM` to `true`.
6. Run `testGoogleFormExpenseSubmit`.

If the form category choices do not match, Google Forms may reject the response.

## Merchant Rules

Merchant Rules let the app auto-categorize repeated merchants.

Columns:

| Column | Meaning |
| --- | --- |
| Pattern | Regex pattern matched against the merchant name. |
| Category | Budget category to apply. |
| Confidence | 0.00 to 1.00 confidence score. |
| Auto Approve | If checked, confident matches can go straight to Expenses. |
| Learning Allowed | If checked, the review app can learn this type of rule. |
| Notes | Human explanation. |
| Enabled | If unchecked, the rule is ignored. |

Broad merchants such as Amazon, Walmart, Target, Costco, PayPal, Venmo, Apple, and Google are treated as ambiguous by default because they can represent many different budget categories.

## Ignored Merchants

Ignored Merchants are merchants you never want imported into the budget.

Use this for things like transfers, statement items, known reimbursements, or charges that are intentionally outside the budget.

The pattern field supports regex. For example, to ignore only a specific subscription and not every merchant with the same broad name, use a precise pattern such as:

```text
^EXAMPLE SUBSCRIPTION MONTHLY\b
```

## Troubleshooting

### Setup says no active spreadsheet was found

The script was probably created as a standalone Apps Script project. Open the Google Sheet first, then go to:

```text
Extensions → Apps Script
```

Paste the code there.

### The review app is blank or says it cannot find `ReviewApp`

Make sure the HTML file in Apps Script is named exactly:

```text
ReviewApp
```

The filename in the Apps Script sidebar should appear as `ReviewApp.html`.

### The web app still shows old code

Apps Script web apps use deployment versions. After editing code, create a new deployment version or edit the existing deployment and select the latest version.

### No emails are imported

Check these items:

- Your credit card transaction alerts are actually arriving in Gmail.
- The emails match the supported Chase or American Express alert formats.
- The emails are newer than `SEARCH_DAYS_BACK`.
- The emails do not already have the `Budget/Processed` label.
- `MAX_THREADS_PER_RUN` is high enough for your inbox volume.

### Transactions go to review instead of auto-approving

That usually means no confident merchant rule was found, the merchant is ambiguous, or AI confidence did not meet the threshold. Approve with **Approve & Learn** to create a rule for next time.

### Google Form submission fails

Turn `PUSH_TO_GOOGLE_FORM` back to `false` until the form is fixed. Then confirm:

- The form URL is the edit URL.
- The form item titles match exactly.
- The category options match `CONFIG.FORM_CATEGORY_MAP`.

### AI categorization fails

Check:

- `OPENWEBUI_API_URL` is the chat completions endpoint.
- `OPENWEBUI_API_KEY` is valid.
- Your OpenWebUI instance is reachable from Google Apps Script.
- `OPENWEBUI_MODEL` is either `AUTO` or an exact model ID.

Run:

```text
testOpenWebUiCategorizerFromAppsScript
```

## Security checklist before publishing your fork

Before making your repository public, verify that these values do not contain private information:

- `OPENWEBUI_API_URL`
- `OPENWEBUI_API_KEY`
- `GOOGLE_FORM_URL`
- `NOTIFICATION_EMAIL`
- `WEB_APP_URL`
- Any spreadsheet URLs
- Any sample email names, account endings, or merchant history you do not want public

## Limitations

- This is not a bank sync tool. It depends on transaction alert emails.
- Parser support is currently focused on Chase and American Express purchase alerts.
- Merchant rules and AI can categorize incorrectly. Review uncertain transactions.
- Apps Script quotas may limit very large inboxes or very frequent imports.
- Google Form integration requires exact form field titles and valid category choices.

## License

This project is licensed under the Apache License 2.0. See `LICENSE` for details.
