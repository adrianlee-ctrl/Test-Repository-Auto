# Setup Guide: Experimentation Repo Automations

Once a week, on Tuesday morning, the script checks the Audience Development tab and sends two kinds of email.

**Lead emails.** Each marketing lead gets one email, and only in weeks when they have something on the list:

- **"Launching within 3 weeks."** This is sent once per test, in the first Tuesday email where the Live Date is 21 days away or less. It's sent again only if the Live Date changes.
- **"Live Date has passed. Did these launch?"** This appears every week while a test's Live Date has passed and its Status is still Planning, Design or In Development. A test drops off as soon as someone updates its Status or Live Date. The script never changes Status itself, and these rows are also highlighted in the sheet.

If Jira is connected, both lists also show each test's Jira ticket status.

**Calendar summary.** One email to your manager (and you), laid out like the Experimentation Calendar slide: Recently Ended, Recently Launched, Upcoming, and Delays/Major Changes, with fiscal weeks. Each cell can be copied into the slide.

The script also keeps a **"Launched Last 7 Days" tab**, which updates on its own whenever the sheet is opened or edited.

Nothing goes to anyone but you until you turn on the weekly emails in the last step.

## If you already installed the earlier version

Open **Extensions → Apps Script**, click your `Repo Automations` file, select all of its code, and replace it with the new `Experiment_Repo_Automations.gs`. Leave your macros file alone. Save, reload the sheet, then follow steps 6 onward. When you reach the end, click **Turn ON weekly emails**. That also removes the old daily schedule, as long as you're the person who turned it on.

## Install (about 10 minutes)

1. Open the repository sheet and go to **Extensions → Apps Script**.
2. Your sheet already has a script file with the team's recorded macros (FilterRefresh, RoadmapRefresh and Refresh). **Leave that file as it is.** In the left sidebar, click the **+** next to **Files**, choose **Script**, and name it `Repo Automations`. Delete the starter `function myFunction() {}`, paste in everything from `Experiment_Repo_Automations.gs`, and click save.
3. In the left sidebar, click the **+** next to **Services**, choose **Google Sheets API**, and click **Add**. This lets the script read the email behind each name chip and the links behind file chips.
4. Still in the editor, find `SUMMARY_RECIPIENTS: [],` near the top and add your manager's email, for example `SUMMARY_RECIPIENTS: ['manager@dowjones.com'],`. You'll get a copy too. Save.
5. Go back to the sheet and reload the page. A **Repo Automations** menu will appear next to Help.

## Check and set up (about 10 minutes)

6. Click **Repo Automations → 1. Check lead emails.** The first time, Google asks for permission: click Continue, choose your account, then Allow. If you see "Google hasn't verified this app," click Advanced, then "Go to Repo Automations." Then click the menu item again. Each lead should show an email address, and GAIN should show as skipped.
7. Click **2. Set up highlighting + launch tab.** It's safe to run again later. It reuses existing tabs instead of adding new ones.

## Connect Jira (about 15 minutes, optional)

This uses Atlassian's free add-on. The script only reads from Jira and never changes anything there.

8. In the sheet, go to **Extensions → Add-ons → Get add-ons**, search for **Jira Cloud for Sheets** (by Atlassian), and install it. If you can't find it or it's blocked, your company's Google admin has to allow Marketplace add-ons first.
9. Go to **Extensions → Jira Cloud for Sheets → Open…**, click **Connect**, choose your atlassian.net site, and click **Accept**.
10. In the same panel, open **Settings** and turn on **Enable =JIRA() function for users of this sheet**. Only the file's **owner** can do this, and Jira lookups will then run through the owner's Jira account. If you aren't the owner, ask them to do this step.
11. Click **Repo Automations → 3. Set up Jira tab.** This does two things:
    - **Adds a "Jira Ticket" column** at the far right of the Audience Development tab, pre-filled with any CIRC tickets already mentioned in Experiment Plan.
    - **Creates a "Jira Data" tab** that the add-on fills in.
12. Wait about a minute, then click **4. Check Jira data.** It lists each test's ticket and its Jira status. If the Jira Data tab shows an error right after you turned the function on, run step 11 again. Atlassian notes this as a one-time glitch.

From now on, put each test's ticket key (for example `CIRC-65857`) in the Jira Ticket column, and separate multiple tickets with commas. One mistyped key can stop the whole Jira lookup; "Check Jira data" will point you to it. The emails still go out without the Jira column if Jira has a problem.

## Test (about 5 minutes)

13. Click **5. Preview this week's emails.** It shows who would get what, plus what goes in each section of the calendar summary. Nothing is sent.
14. Click **6. Send this week's emails to me (test).** You'll receive every lead email and the calendar summary, marked [TEST]. Nobody else is emailed, and nothing is recorded as sent.

## Turn it on (1 minute)

15. Click **Turn ON weekly emails.** Emails go out every Tuesday between 8 and 9am in the sheet's time zone, sent from your Google account. If Jira is set up, the Jira tab refreshes at 6am the same morning. Only one person should turn this on.

## Copying the summary into the slide

The calendar summary uses the slide's four colored sections, with the team on the left and the bold, linked experiment name followed by a drafted sentence on the right. To fill the slide, select the text in one cell of the email, from the bold name to the end, and paste it into the matching cell of the slide table. The bold and the link carry over. If you want the slide's own font and color instead, paste with Ctrl+Shift+V (Cmd+Shift+V on Mac) and re-bold the name. Copying the whole table at once may create a new table on the slide instead of filling the existing one, so going cell by cell is the reliable route.

The sentences are drafts. They come from each test's Description, with side notes like "*Pending…" or "6/17: …" removed and cut to two sentences. Anything in [brackets] needs input, such as the reason for a delay or a result that hasn't been entered. The summary only covers the Audience Development tab, so other teams' rows still come from their own sources.

## What to expect

On a typical Tuesday, most leads get nothing. A lead hears from the script only when one of their tests enters the 3-week window or is past its Live Date without an update. Your manager gets the calendar summary every Tuesday, even when some sections are empty. Everything sent is recorded in the Automation Log tab.

## Changing settings

Everything adjustable is in the `CONFIG` block at the top of the script:

- **Send day and time:** `SEND_DAY`, `SEND_HOUR`
- **Reminder window:** `REMINDER_DAYS_BEFORE` (21 days)
- **Statuses to watch:** `REMINDER_STATUSES`, `LAUNCH_CHECK_STATUSES`
- **Summary recipients:** `SUMMARY_RECIPIENTS`
- **Skip list:** `SKIP_LEADS` (GAIN for now)
- **Jira:** `JIRA_PROJECT_KEYS`, `JIRA_DONE_STATUSES`
- **Pause:** set `PAUSED: true` to stop all emails without touching the schedule. Anyone with edit access can do this.
- **Fiscal weeks:** `FISCAL_YEAR_START` is the Monday that starts FW1. It's set to 6/29/2026, which makes 10/19–10/25 FW17 and matches your FW16 review deck. Update it when FY28 starts.

Click save after any change. If you change the send day or time, click **Turn ON weekly emails** again.

## Troubleshooting

**A lead shows "NO EMAIL FOUND."** The name was typed as plain text instead of inserted as a name chip. Click the cell, type @, and pick the person. As a backup, you can add a tab named "Leads" with names in column A and emails in column B.

**Every lead shows "NO EMAIL FOUND" with a note that the chips couldn't be read.** Delete the `/** @OnlyCurrentDoc */` line at the top of your macros file, save, and run the check again. Your macros keep working.

**The Jira column is missing from Tuesday's email.** Click "4. Check Jira data." It explains whether the =JIRA() function is off, a ticket key is mistyped, or the tab hasn't loaded. If the data looks fine when you open the sheet but was missing from the email, let me know. That would mean the add-on only refreshes while someone has the sheet open, and the fix would be to connect to Jira directly instead.

**You see extra tabs after running setup.** Delete the duplicates. The current version recognizes its tabs even if their capitalization or spacing is slightly different.

**A column gets renamed.** Update the matching name in `CONFIG.COLUMNS`. Moving columns around is fine and needs no changes.

**You want to stop the emails.** Click **Turn OFF weekly emails**. This must be done by the person who turned them on. Anyone else can set `PAUSED: true`.
