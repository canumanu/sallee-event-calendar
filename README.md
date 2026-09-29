# Sallee Event Calendar

The racing and sales calendar for Sallee's social media team. The page is hosted on GitHub Pages and people sign in with their Sallee Microsoft 365 account. The events live in a SharePoint list on the **Social Media & Marketing** site. The site's own permissions decide who can edit and who can only view, and Power Automate sends a reminder one week before each event.

```
GitHub Pages (this repo)  ──sign-in──▶  Microsoft 365
        │                                   │
        └──── Microsoft Graph ────▶  SharePoint: Social Media & Marketing site
                                       ├─ Event Calendar            (the events)
                                       └─ Event Reminder Recipients (who gets reminders, and how)
                                                  │
                             Power Automate (daily) ──▶ email · Teams · text
```

## What's in this folder

| File | What it is |
|---|---|
| `index.html` | The calendar app |
| `config.js` | Your tenant, site and list settings. This is the only file you edit. |
| `data.js` | Sign-in and the SharePoint read/write code |
| `events.json` | The 2026 events, imported into SharePoint on first run |
| `logos/` | Venue logos (internal use only) |
| `vendor/msal-browser.min.js` | Microsoft's sign-in library (v3.28.1), stored here so the page doesn't depend on a third-party CDN |

If you open the page before `config.js` is filled in, it runs in **demo mode** with the sample events and saves nothing. That's handy for checking the page before you connect it.

---

## Step 1 — App registration (Entra ID)

You can reuse the dispatch board's app registration or create a new one.

1. Go to **Entra admin center → App registrations →** (the app) **→ Authentication → Add a platform → Single-page application**.
2. Add both redirect URIs, using your GitHub Pages address:
   - `https://<github-user>.github.io/<repo>/`
   - `https://<github-user>.github.io/<repo>/index.html`
3. Go to **API permissions → Add a permission** and add these **delegated** permissions:
   - Microsoft Graph: `User.Read`, `Sites.ReadWrite.All`
   - Microsoft Graph: `Sites.Manage.All`. This is only needed for the one-time setup in step 5, which creates the lists. You can remove it afterwards.
   - SharePoint: `AllSites.Read`. This is optional. It lets the page tell view-only people up front that they can't edit. Without it, they only find out when a save fails.
4. Click **Grant admin consent for Sallee**. If the dispatch board uses this same app, this also clears the consent it has been waiting on.
5. Copy the **Application (client) ID** and **Directory (tenant) ID**.

## Step 2 — Fill in `config.js`

```js
clientId: "<Application (client) ID>",
tenantId: "<Directory (tenant) ID>",
siteHost: "<tenant>.sharepoint.com",
sitePath: "/sites/<social-media-site>",   // copy it from the site's URL
```

## Step 3 — Publish on GitHub Pages

1. Create a repo, for example `sallee-event-calendar`, and push this folder to it.
2. Go to **Settings → Pages → Deploy from branch → `main` / root**.
3. Open `https://<github-user>.github.io/<repo>/`.

Pages on a public repo means anyone can read the code and `events.json`. The code holds no secrets: a client ID isn't a password, and all the data sits behind Microsoft sign-in. The event dates in `events.json` are the published racing and sales schedules. If you'd rather not leave them in the repo, delete `events.json` after step 5. A private repo with Pages needs a paid GitHub plan.

## Step 4 — Who can do what

The page follows the SharePoint site's permissions (**Site settings → Site permissions**):

| SharePoint role | In the calendar |
|---|---|
| **Owners** | Everything, including the one-time setup |
| **Members** | Add, edit and delete events; check off shots; write notes; plan next year |
| **Visitors** | View only |
| Not on the site | Can't open the data. The page tells them to ask you for access. |

Add the social media agent as a **Member**. Add the rest of the crew as Members or Visitors, depending on whether they should edit.

## Step 5 — First run (site owner)

1. Open the page and sign in.
2. The page shows **One-time setup**. Click **Create lists and import events**. Microsoft asks once to approve the extra permission.
3. The page creates two lists and imports the 54 events from `events.json`:
   - **Event Calendar**: Title, Venue, Type, Start, End, Location, Marquee, Footage status, Dates TBC, Notes, plus two app-data columns (shot list and copied-from).
   - **Event Reminder Recipients**: Name, Email, Mobile, Email / Teams / Text reminders (yes/no), Active.

If you edit events directly in the SharePoint list, keep **Start** and **End** as `YYYY-MM-DD` (for example `2026-10-30`). They're stored as text so the page and the reminder flow read them the same way in every time zone.

## Step 6 — Add reminder recipients

Once you're signed in as a Member, the toolbar has a **Reminder recipients** button that opens the list. Add one row per person:

| Name | Email | Mobile | Email reminders | Teams reminders | Text reminders | Active |
|---|---|---|---|---|---|---|
| (social media agent) | name@… | +15555550123 | Yes | Yes | Yes | Yes |

---

## Step 7 — Reminder flow (Power Automate)

This sends one message per event, **7 days before it starts**. Build it at **make.powerautomate.com → Create → Scheduled cloud flow**.

**Trigger: Recurrence**. Every 1 day, time zone *(UTC-05:00) Eastern Time*, at hour `7`, minute `0`.

**1. Compose**, renamed `Target date`:
```
formatDateTime(addDays(convertTimeZone(utcNow(),'UTC','Eastern Standard Time'),7),'yyyy-MM-dd')
```

**2. SharePoint → Get items**, renamed `Events next week`
- Site address: the Social Media & Marketing site
- List name: `Event Calendar`
- Filter query: `EventStart eq '@{outputs('Target_date')}'`

**3. SharePoint → Get items**, renamed `Recipients`
- List name: `Event Reminder Recipients`
- Filter query: `Active eq 1`

**4. Apply to each**, over `value` from *Events next week*

Inside it, add **Compose**, renamed `Message`:
```
concat('1 week out: ', items('Apply_to_each')?['Title'], ' — ', items('Apply_to_each')?['EventLocation'], ', ', items('Apply_to_each')?['EventStart'], if(equals(items('Apply_to_each')?['EventEnd'], items('Apply_to_each')?['EventStart']), '', concat(' to ', items('Apply_to_each')?['EventEnd'])), '. Plan the shoot: https://<github-user>.github.io/<repo>/')
```

Also inside it, add **Apply to each 2**, over `value` from *Recipients*, with three conditions:

| If | Action | Settings |
|---|---|---|
| `ByEmail` is equal to `true` | **Office 365 Outlook → Send an email (V2)** | To: `Email` · Subject: `Event reminder: ` + Title · Body: output of *Message* |
| `ByTeams` is equal to `true` | **Microsoft Teams → Post message in a chat or channel** | Post as: *Flow bot* · Post in: *Chat with Flow bot* · Recipient: `Email` · Message: *Message* |
| `ByText` is equal to `true` | **Twilio → Send Text Message (SMS)** | From: your Twilio number · To: `Mobile` · Text: *Message* |

To post once to a team channel instead of messaging each person in Teams, move the Teams action out of *Apply to each 2* and set **Post in: Channel**.

### About text messages

Microsoft 365 can't send SMS on its own, so there are two ways to do it:

1. **Twilio (real SMS).** This needs a **Power Automate Premium** license for whoever owns the flow, because Twilio is a premium connector. It also needs a Twilio account with a phone number, and US business texting requires A2P 10DLC registration through Twilio. Twilio charges per text. Check current pricing on both.
2. **Teams on the phone (no extra cost).** If everyone has the Teams mobile app, the Teams reminder already arrives as a push notification. For most crews this does the job of a text.

A good way to start is with email and Teams, then add Twilio if people still want real texts. Leave the **Text reminders** column empty until then.

**Test it:** temporarily change the *Target date* expression to `…addDays(…, 1)…` or pick a date that has an event, then run the flow manually.

---

## Every year

- On the calendar, click **Plan next year**. It copies the season forward with **Dates TBC** tags.
- As the new schedules come out, fix each event's dates and tick **Dates confirmed**.
- To add a new venue logo, save it as `logos/<code>.jpg` and add a line to `VENUES` in `index.html`. Event-only logos (like the Breeders' Cup) go in `EVENT_LOGOS` with the year.

Logos are other companies' trademarks, used here only to label events on an internal tool. Don't use them in public posts without the venue's permission.
