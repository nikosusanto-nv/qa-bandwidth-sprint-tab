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

## After the QA ticket discussion

Move each ticket under the QA who will test it in the **BREAKDOWN** section. Row 7 holds each QA's name with their Jira email, for example `Niko niko.susanto@ninjavan.co`. A ticket can appear under more than one QA.

Then, with that sprint's tab open, choose **QA Tools → Assign QA tickets in Jira**. The script lists every change it will make and only touches Jira after you click **Yes**:

- **Title has `[QA]`, one QA:** the ticket itself is assigned to that QA, and so is its `[QA]` sub-task (see below).
- **Every ticket:** each listed QA gets their own `[QA]` sub-task. A sub-task already assigned to them is kept. Otherwise an open `[QA]` sub-task that no listed QA holds is reused, and if none is left, a new `[QA] Execute testing` sub-task is created and assigned.

Tickets that are already assigned correctly are skipped, so it's safe to run again after changing the sheet.

## Settings

Change these in `CONFIG` at the top of `Code.gs`: board ID, template tab gid (`1704848751`), start row, the QA Estimate field ID, the new sub-task's title, and `QA_USERS` (Jira usernames used when row 7 has no email).
