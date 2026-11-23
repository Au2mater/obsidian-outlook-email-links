# Outlook Email Links for Obsidian

Drag an email from Outlook on the web, or the Outlook app installed from it
(PWA), onto a note. A link to that email is inserted where you drop it:

```markdown
[📧 RE: Budget for Q3](https://outlook.office365.com/mail/deeplink/read/AAkALg…)
```

Clicking the link opens the email in Outlook. Drag several selected emails at
once and each gets its own line.

- **No sign-in and no network.** Outlook puts the email's subject and ID on
  the drag; the plugin turns those into a link. It never contacts Microsoft
  and copies nothing of the email but what the link template asks for.
- **Opens in the Outlook app** on Windows, if you installed Outlook as an
  app from Edge (see below).
- **Links should survive moving the email.** The ID Outlook hands over
  looks like an Exchange *immutable* ID, which Microsoft documents as staying
  the same when an email moves to another folder in the same mailbox.
- Drops that are not Outlook emails are left to Obsidian as usual.

This plugin is not made by or affiliated with Microsoft. Outlook is a
trademark of Microsoft.

## Opening links in the Outlook app

On Windows, clicking an Outlook link opens it in the Outlook app you
installed from Edge (the PWA), in a window of its own. Without that app, and
on other systems, links open in your browser.

- The first click looks for the app, which takes a second or two. The plugin
  asks Windows, with PowerShell, which Outlook apps Edge has installed, and
  remembers the one it finds in `data.json`. **Find app** in the settings
  looks again, for example after you reinstall the app.
- It opens the app with Edge's `msedge_proxy.exe`, the way the app's own
  shortcuts do, and moves the link to the app's address, such as
  `outlook.office365.com`. Links made with another Outlook address, such as
  `outlook.office.com`, also open in the app.
- This covers every work or school Outlook mail link in your notes
  (addresses with `/mail/`), not only the ones this plugin made. Calendar
  links and personal Outlook (`outlook.live.com`) links open in the browser.
  Turn **Open links in the Outlook app** off to send them all to the browser.
- Outlook installed as an app from Chrome isn't detected.

## Settings

| Setting | Default | Notes |
| --- | --- | --- |
| Link template | `[📧 {{subject}}]({{url}})` | `{{subject}}`, `{{url}}`, `{{id}}`, `{{messageId}}`, `{{mailbox}}` |
| Email URL | `https://outlook.office365.com/mail/deeplink/read/{{id}}` | `{{id}}` is the URL-encoded item ID, `{{rawId}}` the ID as is; also `{{mailbox}}`, `{{messageId}}` |
| Between emails | new line | Put between links when several emails are dropped. |
| Open links in the Outlook app | on | Windows only. |

`{{messageId}}` is the email's Internet message ID (`<…@…>`), which is the
same in every copy of the email, and `{{mailbox}}` is the mailbox it was
dragged from.

## Limits

- Only Outlook on the web and the apps built on it (the PWA, new Outlook for
  Windows) put the email's ID on a drag. Classic Outlook for Windows drags a
  file instead, which this plugin ignores.
- An email moved to another mailbox, such as an online archive, gets a new
  ID, and its old link stops working.
- A link opens the email only for someone signed in to that mailbox.

## Install

- **With BRAT:** add `Au2mater/obsidian-outlook-email-links`.
- **Manually:** download `main.js` and `manifest.json` from the latest
  release into `<vault>/.obsidian/plugins/outlook-email-links/`, then enable
  the plugin under **Settings → Community plugins**.

## Development

Plain JavaScript with no dependencies. The source is `src/plugin.js`;
`npm run build` checks it and writes `main.js`, the file Obsidian loads.

```bash
./install.sh "/path/to/vault/.obsidian"   # builds, then copies
npm test
```

To release, bump `version` in `manifest.json` and `package.json`, add it to
`versions.json` and commit, then push a tag such as `0.1.0`; the workflow
builds, tests and publishes the release.
