# Setup Guide: Experimentation Repo Automations

This adds three things to the Audience Development repository. First, a daily email reminds each marketing lead when one of their tests is 3 weeks or less from its Live Date. Second, rows whose Live Date has passed but whose Status is still Planning, Design or In Development get highlighted, and the lead gets one email asking them to confirm. The script never changes the Status itself. Third, a new "Launched Last 7 Days" tab always lists tests with a Live Date in the past week.

Lead emails come straight from the name chips in the Marketing Lead column, so you don't need a separate email column. When a cell has two people, both get the email. GAIN is skipped for now.

Nothing goes to the marketing leads until the last step, so it's safe to install this directly on the real sheet. Total time is about 25 minutes.

## Install (about 10 minutes)

1. Open the repository sheet and go to **Extensions → Apps Script**.
2. Your sheet already has a script file with the team's recorded macros (FilterRefresh, RoadmapRefresh and Refresh). **Leave that file as it is.** In the left sidebar, click the **+** next to **Files**, choose **Script**, and name the new file `Repo Automations`. Delete the starter `function myFunction() {}` that appears in the new file, paste in everything from `Experiment_Repo_Automations.gs`, and click the save icon. Your macros and the new script will run side by side without affecting each other.
3. In the left sidebar, click the **+** next to **Services**, choose **Google Sheets API**, and click **Add**. This lets the script read the email behind each name chip.
4. Go back to the sheet and reload the page. A **Repo Automations** menu will appear next to Help after a few seconds.

## Test (about 15 minutes)

5. Click **Repo Automations → 1. Check lead emails.** The first time, Google asks for permission: click Continue, pick your account, then Allow. If you see "Google hasn't verified this app," click Advanced, then "Go to Repo Automations." That warning is normal for scripts built inside a company. If your company blocks it entirely, IT will need to approve the script. Once approved, click the menu item again. You should see each lead with their email, and GAIN listed as skipped.
6. Click **2. Set up highlighting + launch tab.** This adds the highlight rule, the "Launched Last 7 Days" tab, and an "Automation Log" tab that records every email sent.
7. Click **3. Preview today's emails.** This shows exactly who would get what today. Nothing is sent.
8. Click **4. Send today's emails to me (test).** You'll receive the real emails, marked [TEST], so you can see what the leads will see. Leads are not emailed.

## Turn it on (1 minute)

9. Click **Turn ON daily emails.** The check runs every morning between 8 and 9am in the sheet's time zone, and emails are sent from your Google account. Only one person should turn this on.

## What to expect

Based on the current data, the first run sends Sabrina one email covering three upcoming tests (AUD-172, AUD-155, AUD-146) and five tests whose Live Date has passed without a status change (AUD-123, AUD-150, AUD-152, AUD-153, AUD-154). After that, leads only hear from it when something new happens: a test enters the 3-week window, a Live Date passes while the test is still in Planning, Design or In Development, or a Live Date changes. Each lead gets at most one combined email per day.

Highlighted rows clear on their own once someone updates the Status or moves the Live Date. Tests with "NA," "Q2" or a blank Live Date are ignored until a real date is entered.

## Changing settings

Everything adjustable is in the `CONFIG` block at the top of the script: the 3-week window, which statuses get reminders or highlights, the highlight color, an optional CC address, the time the emails go out, and the skip list. When you have a contact for GAIN, remove `'GAIN'` from `SKIP_LEADS` and make sure the GAIN cells have a name chip (or add GAIN to a Leads tab, described below). To pause all emails without turning off the schedule, change `PAUSED: false` to `PAUSED: true`. Anyone with edit access can do this, which is useful if you're out and something needs to stop. Click save after any change. If you change the highlight statuses or color, run step 6 again.

## Troubleshooting

**A lead shows "NO EMAIL FOUND."** The name was typed as plain text rather than inserted as a name chip. Click the cell, type @, and pick the person. Alternatively, add a tab named "Leads" with names in column A and emails in column B (headers in row 1), and the script will use it as a backup.

**Every lead shows "NO EMAIL FOUND" and there's a note saying the name chips couldn't be read.** This is most likely caused by the `/** @OnlyCurrentDoc */` line at the top of your macros file. That line limits the whole project to the current sheet, and it can stop the script from reading name chips. Delete just that one line, save, and run "Check lead emails" again. Google will ask you to approve permissions once more. Your macros keep working as before.

**The menu doesn't appear.** Reload the sheet and confirm the script was saved.

**You want to pause the emails.** Click **Repo Automations → Turn OFF daily emails.** This must be done by the person who turned them on.

**A column gets renamed.** The script finds columns by name, so update the matching name in `CONFIG.COLUMNS`. Moving columns around is fine and needs no changes.
