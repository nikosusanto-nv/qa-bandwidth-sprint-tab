# QA sprint tab script

Google Apps Script for the **QA Bandwith Calculation** sheet. It copies the template tab, names the copy after the active sprint on Jira board 220 (for example `XB Sprint 191`), and lists the sprint's QA tickets from **B21** down.

- **Column B:** `XXB-1234 <title>`, linked to the Jira ticket.
- **Column C:** the ticket's QA Estimate (QA story point).
- **Which tickets:** anything that isn't a sub-task and either has `[QA]` in its title or a QA Estimate value.
- **More than 10 tickets:** extra rows are inserted so the Total row still adds up every ticket.

## One-time setup

1. Open the sheet, then go to **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with this folder's `Code.gs`, then save.
3. Reload the sheet. A **QA Tools** menu appears.
4. Create a Jira Personal Access Token: in Jira, click your avatar → **Profile** → **Personal Access Tokens** → **Create token**.
5. In the sheet, go to **QA Tools → Set Jira token** and paste the token. It's saved for your Google account only.

The first time you run a menu item, Google asks you to authorize the script. It needs access to the spreadsheet and to external URLs (Jira).

## Each sprint

Choose **QA Tools → Create tab for active sprint**. The new tab is added as the first tab. If a tab with that sprint's name already exists, the script stops without changing anything.

## Settings

Change these in `CONFIG` at the top of `Code.gs`: board ID, template tab gid (`1704848751`), start row, and the QA Estimate field ID.
