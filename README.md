# BNI Elli tools

Small, single-file web pages for the chapter, served by GitHub Pages. The Passport page keeps its data in a Google Sheet.

| Page | Address | What it is for |
|---|---|---|
| Action Plan | https://cerithw.github.io/bni-elli/ | Monthly member action plan with the points and Traffic Light status |
| Passport to Success | https://cerithw.github.io/bni-elli/passport/ | New member mentoring: each member's own passport, mentor crib sheets for all eleven sessions, the Mentor Coordinator's tracker, and the reference rules |

## Passport to Success

Built from the BNI South Wales *Passport to Success 2026-27* pack (ten mentor crib sheets, ten slide decks, the mentor tracking sheet, the effectiveness review checklist and the General Policies), de-duplicated onto one page and connected to the chapter's Google Sheet.

- **My Passport** – the new member signs in (pick your name, enter your password) and sees their own tracker: the eleven sessions plus MSP, each with who to book it with, what it covers, what to do before and after, a *booked for* date and a stamp when signed off. Shows week number, the 84-day target date and pace.
- **Sessions** – the ten Passport 121s with chapter members, then **Session 11, Meeting with Chapter Ambassador**: the final sign-off, covering the regional team, networking across chapters, regional events and the member's contact sphere, plus the last check (ten sessions signed, MSP done, Connect profile at 100%). It is delivered by the Chapter Ambassador (Barrie H), or the Director Consultant if the Ambassador isn't available.
- **Mentors** – crib sheets for every session, open to everyone (`…/passport/#mentor-7` links straight to one). A mentor who signs in also sees *their* new members: which sessions they owe, a date picker to book each one, **Signed off today**, and a ready-made booking message. Sessions where they are only the alternative mentor are tucked under a separate heading.
- **Coordinator tracker** – every new member's progress and mentor pairings; booked 121s with any missed dates flagged; **Copy weekly summary** and **Email summary now**; a **Welcome message** per member (link, their name as it appears in the list, their password); an Archive for members whose certificate has been presented, with how many completed within 90 days; a **Programme health** panel showing which sessions lag; a printable **Certificate**; and the QR code for the induction pack.
- **Reference** – the rules that recur across the sessions.

### The Google Sheet behind it

The page reads and writes a Google Sheet through the script in [`passport/sheet-script/Code.gs`](passport/sheet-script/Code.gs). Tabs:

- **Members** – one row per new member: name, password, classification, induction date, then for each session the sign-off date (`S1…S11`, `MSP`), the mentor (`Mentor S1…`) and the booked date (`Booked S1…`). Dates are plain text like `2026-10-22`; `yes` means done with no date recorded. Signed-off cells turn green and booked cells amber.
- **Archive** – members whose certificate has been presented, with the completion date.
- **Settings** – chapter name; coordinator password; *Session N mentors* (usual first, alternative second); one *Mentor password: Name* row per mentor; the page link; and the weekly email (on/off and who receives it).
- **Dashboard** – live formulas over Members: week, target date, sessions done, status.

Every **Wednesday at about 6pm** the script emails the weekly summary (each member's progress and next session, and each mentor's outstanding sessions) to the address in Settings. **Email summary now** on the Coordinator tab sends the same thing on demand.

### Setting up, or updating, the sheet script

First time: create a blank Google Sheet with *Elli* in its name, open **Extensions → Apps Script**, paste in the whole of `Code.gs`, save, choose **setup** in the function dropdown and **Run** (approve the permissions: it needs the sheet, email and a timer), then **Deploy → New deployment → Web app**, *Execute as* **Me**, *Who has access* **Anyone**, and put the web app URL into `apiUrl` in the `CHAPTER` block of `passport/index.html`.

When `Code.gs` changes (the current version is 3, which adds Session 11): paste the new code over the old, save, run **setup** again (it only adds what is missing — new columns, Settings rows, the Dashboard, the timer — and never touches existing members), then **Deploy → Manage deployments → ✎ → Version: New version → Deploy**. The URL stays the same.

Afterwards:

- Passwords: members' in column B of **Members**; the coordinator's and each mentor's in **Settings**. Change any of them whenever you like.
- To change who delivers a session, edit the *Session N mentors* rows; run **setup** again if a new mentor name needs a password row.
- If sign-in says it cannot reach the store, the usual cause is the deployment's *Who has access* not being **Anyone**.

### Chapter settings in the page

The `CHAPTER` block near the top of the script in `passport/index.html` holds the web app address, the chapter name, the mentor roster and the members used when the page runs without a sheet (its on-device fallback). Online, the sheet's **Settings** tab overrides the name and roster.

## Editing

Each page is one `index.html` with its styles and script inside. Edit the file, commit to `main`, and GitHub Pages publishes it within a minute or two.
