# Availability Picker — Deployment guide for IT / Microsoft 365 admins

This is a small Outlook add-in: a task pane that reads the signed-in user's
**own** calendar via Microsoft Graph (`Calendars.Read`, delegated — read-only,
per-user), shows their open time slots on a mini calendar, and lets them
insert a list of times into the email they're composing. It never sends
anything on its own.

Users can optionally add colleagues (searched from the org directory via
`People.Read`) to check group availability — this uses Graph's `getSchedule`
API, the same one behind Outlook's Scheduling Assistant, which only ever
returns free/busy/tentative/out-of-office *status* for other users, never
event details or subjects. `Calendars.Read` (already required above) is
sufficient for this; no extra calendar permission is needed for checking
colleagues' schedules.

Source code: this folder. Currently hosted at:
`https://ag-guy-dev.github.io/calendar-availability-picker/`

## What needs to happen

1. Register an Azure AD (Entra ID) app in your own tenant and grant admin
   consent for `Calendars.Read` and `People.Read`.
2. Point the add-in's config at that app.
3. Push the add-in to users via the Microsoft 365 admin center.

## Step 1 — Register the app in your tenant (recommended path)

An app registration already exists for this add-in, but it was created under
the original requester's **personal** Microsoft account, not your
organization's tenant. That's fine for one person testing it, but for an
org-wide rollout you'll want your own tenant to own the registration (so IT
retains control regardless of who requested this, and consent/audit trails
live in your directory). To do that:

1. Go to **[entra.microsoft.com](https://entra.microsoft.com)** (or
   portal.azure.com → Microsoft Entra ID) in your admin account.
2. **App registrations → New registration.**
3. Name it (e.g. `Availability Picker`).
4. Supported account types: **"Accounts in this organizational directory
   only"** — this is an internal tool, no need for multi-tenant or personal
   accounts.
5. Redirect URI: platform = **Single-page application (SPA)**, value =
   `https://ag-guy-dev.github.io/calendar-availability-picker/taskpane.html`
   (or your own hosted URL, see "Hosting" below).
6. Register, then copy the **Application (client) ID** from the Overview page.
7. **API permissions → Add a permission → Microsoft Graph → Delegated
   permissions →** add `Calendars.Read` and `People.Read`.
8. Click **Grant admin consent for [your org]** — this is the step that lets
   every user skip the individual consent prompt.

*(Faster alternative: instead of a new registration, you could grant
tenant-wide admin consent to the existing app via
`https://login.microsoftonline.com/{your-tenant-id}/adminconsent?client_id={existing-client-id}`.
That works technically, but the registration stays owned by an individual's
personal directory — most admins will prefer Step 1 above for that reason.)*

## Step 2 — Update the add-in's config

In `config.js`, set:

```js
clientId: "THE_CLIENT_ID_FROM_STEP_1",
authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
```

Using your actual tenant ID (instead of the generic `organizations` endpoint
currently in the file) restricts sign-in to only your org's accounts.

Then redeploy `config.js` to wherever `taskpane.html` is hosted (see
[README.md](README.md) for how the current GitHub Pages hosting works, if
you're keeping it).

## Hosting

Right now everything is hosted on a personal GitHub Pages site
(`ag-guy-dev.github.io`). That works fine functionally, but you may prefer to
move hosting under org-controlled infrastructure (Azure Static Web Apps in
your own subscription, or an internal web server) for continuity/ownership
reasons — anyone can keep maintaining a personal GitHub repo, but IT usually
wants this under its own control for a tool deployed to the whole org. If you
do move it, update the Redirect URI in Step 1 and the URLs in `manifest.xml`
to match the new domain.

## Step 3 — Fill in `manifest.xml`

- `<ProviderName>` — currently a placeholder (`REPLACE_WITH_YOUR_ORG_NAME`);
  set it to your organization's name.
- Confirm all the URLs under `<IconUrl>`, `<SourceLocation>`, `<SupportUrl>`,
  and inside `<Resources>` point to wherever you're hosting these files.

## Step 4 — Deploy org-wide via the Microsoft 365 admin center

1. Go to **[admin.microsoft.com](https://admin.microsoft.com)** → **Settings
   → Integrated apps**.
2. **Upload custom apps** → upload your finished `manifest.xml`.
3. Choose who gets it: everyone, specific groups, or specific users (good for
   a pilot before rolling out to everyone).
4. Deploy. Microsoft 365 pushes it automatically to those users' Outlook
   (web, Windows, and Mac) — no manual sideloading needed on their end, and
   this deployment step itself also handles/confirms the permission consent
   for the assigned users.

## Note for whoever originally set this up personally

Once the org-wide deployment is live, remove your own manual sideloaded
copies (Outlook on the web: Settings → Get Add-ins → My add-ins → remove; and
on Mac, delete the file from
`~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/`) so you
don't end up with two copies of the button.
