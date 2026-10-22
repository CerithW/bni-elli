# BNI Elli tools

Small, single-file web pages for the chapter, served by GitHub Pages.

| Page | Address | What it is for |
|---|---|---|
| Action Plan | https://cerithw.github.io/bni-elli/ | Monthly member action plan with the points and Traffic Light status |
| Passport to Success | https://cerithw.github.io/bni-elli/passport/ | New member mentoring: each member's own passport, mentor crib sheets for all ten sessions, the Mentor Coordinator's tracker, and the reference rules |

## Passport to Success

Built from the BNI South Wales *Passport to Success 2026-27* pack (ten mentor crib sheets, ten slide decks, the mentor tracking sheet, the effectiveness review checklist and the General Policies), de-duplicated onto one page.

- **My Passport** – the new member's own tracker: the ten sessions plus MSP, each with who to book it with, what it covers, what to do before and after, and a stamp when signed off. Shows week number, the 84-day target date and pace.
- **Mentor crib sheets** – one screen per session, open to everyone. A direct link to a session is `…/passport/#mentor-7` (change the number).
- **Coordinator tracker** – every new member's progress, mentor pairings, a copyable weekly summary for Announcements, and an Archive for members whose certificate has been presented (with the "completed within 90 days" figure).
- **Reference** – the rules that recur across the sessions.

### Two ways it can run

1. **On-device** (the default until the sheet below is connected). Everything is saved in the browser of the device it is opened on. The Coordinator tracker has Export / Import for handover.
2. **Online, with a Google Sheet behind it.** Members pick their name, enter the password you gave them, and see only their own passport; what they stamp is saved to the sheet. The Coordinator signs in with a separate password and sees everyone. The sheet itself stays the master record and you can edit it directly.

### Connecting the Google Sheet (one-time, about five minutes)

1. Create a blank Google Sheet (go to **sheets.new**) and name it, e.g. *Passport to Success – BNI Elli*.
2. In the sheet, open **Extensions → Apps Script**. Delete the sample code, paste in the whole of [`passport/sheet-script/Code.gs`](passport/sheet-script/Code.gs), and click the save icon.
3. In the toolbar, choose the function **setup** from the dropdown and click **Run**. Google will ask you to review permissions: choose your account, click *Advanced → Go to … (unsafe)* (it is your own script), then *Allow*. Running it creates three tabs — **Members**, **Archive**, **Settings** — with the five current members, a password for each, the mentor roster and a coordinator password.
4. Click **Deploy → New deployment**. Click the gear next to *Select type* and choose **Web app**. Set *Execute as* to **Me** and *Who has access* to **Anyone**. Click **Deploy**, then **Copy** the Web app URL (it ends in `/exec`).
5. Paste that URL into `apiUrl` in the `CHAPTER` block near the top of `passport/index.html`, between the quotes — or send it to Claude to do. Commit to `main` and the page switches to online mode within a couple of minutes.

Afterwards:

- Passwords are in the **Members** tab (column B). Change any of them whenever you like; the coordinator password is in **Settings**.
- Add members in the page or straight in the sheet (name, password, induction date as `2026-10-22`). Dates in the sheet are plain text so they are never mangled.
- To change who delivers a session, edit the *Session N mentors* rows in **Settings**: usual mentor first, alternative second.
- If you change the script later, use **Deploy → Manage deployments → edit → New version** so the same URL serves the new code.

If sign-in says it cannot reach the store, the usual cause is step 4: *Who has access* must be **Anyone**.

### Chapter settings in the page

The `CHAPTER` block near the top of the script in `passport/index.html` holds the chapter name, the mentor roster, whether Launch Pad and the Regional Team sign-off are tracked, and the members loaded on first open in on-device mode. Online, the sheet's **Settings** tab overrides the name and roster.

## Editing

Each page is one `index.html` with its styles and script inside. Edit the file, commit to `main`, and GitHub Pages publishes it within a minute or two.
