# Outlook Email Links for Obsidian

Drag an email from Outlook on the web, or the Outlook app installed from it
(PWA), onto a note. A link to that email is inserted where you drop it:

```markdown
[📧 RE: Budget for Q3](https://outlook.office365.com/mail/deeplink/read/AAkALg…)
```

Clicking the link opens the email in Outlook. Drag several selected emails at
once and each gets its own line.

Drag a folder from Outlook's folder pane and you get a link that opens that
folder:

```markdown
[📁 01 Projects](https://outlook.office365.com/mail/AAMkAD…)
```

- **No sign-in and no network.** Outlook puts the email's subject and ID on
  the drag; the plugin turns those into a link. It never contacts Microsoft
  and copies nothing of the email but what the link template asks for.
- **Opens in the Outlook app** on Windows, if you installed Outlook as an
  app from Edge (see below).
- **Links survive moving the email** to another folder. The ID Outlook hands
  over is an Exchange *immutable* ID, which stays the same while the email
  stays in the same mailbox.
- Drops that are not Outlook emails or folders are left to Obsidian as usual.

This plugin is not made by or affiliated with Microsoft. Outlook is a
trademark of Microsoft.

## Disclosures

- **Network use:** none. The plugin makes no network requests.
- **Programs it runs (Windows only):** to open links in the Outlook app, the
  plugin runs two programs with Node's `child_process`, which is why
  Obsidian's plugin review warns about shell execution:
  - `powershell.exe`, to ask Windows which Outlook apps Edge has installed
    (`Get-AppxPackage`, reading each Outlook app's `AppxManifest.xml`). It
    runs on the first Outlook link you click, and again only when you select
    **Find app**, when opening the app fails, or, if no app was found, on
    the first click after Obsidian restarts.
  - Edge's `msedge_proxy.exe`, to open a link in the app.

  Both run without a shell, and the link is passed as a single argument, so
  a link can't run commands. Nothing runs on other systems, and nothing runs
  if you turn off **Open links in the Outlook app** before your first click.
- **Files:** the plugin reads and writes no files itself beyond its own
  settings in `data.json`. The app it found (app ID, browser profile, start
  address and which `msedge_proxy.exe` started it) is kept with Obsidian's
  local storage on this computer, outside the vault's files, so it doesn't
  sync to other computers.
- **No telemetry, ads, payments or account.**

## Opening links in the Outlook app

On Windows, clicking an Outlook link opens it in the Outlook app you
installed from Edge (the PWA), in a window of its own. Without that app, and
on other systems, links open in your browser.

- The first click looks for the app, which takes a second or two. The plugin
  asks Windows, with PowerShell, which Outlook apps Edge has installed, and
  remembers the one it finds on this computer. **Find app** in the settings
  looks again, for example after you reinstall the app.
- It opens the app with Edge's `msedge_proxy.exe`, the way the app's own
  shortcuts do. Only `msedge_proxy.exe` in Edge's own install folders is
  ever run. It moves the link to the app's address, such as
  `outlook.office365.com`. Links made with another Outlook address, such as
  `outlook.office.com`, or copied from Outlook's address bar
  (`outlook.cloud.microsoft`), also open in the app.
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
| Folder link template | `[📁 {{name}}]({{url}})` | `{{name}}`, `{{url}}`, `{{id}}`, `{{mailbox}}` |
| Folder URL | `https://outlook.office365.com/mail/{{id}}` | `{{id}}` is the folder ID URL-encoded twice (see below), `{{rawId}}` the ID as is |
| Between emails | new line | Put between links when several emails are dropped. |
| Open links in the Outlook app | on | Windows only. |

`{{messageId}}` is the email's Internet message ID (`<…@…>`), which is the
same in every copy of the email, and `{{mailbox}}` is the mailbox it was
dragged from.

## Limits

- Folder links rely on how Outlook's addresses forward. `outlook.office365.com`
  forwards to `outlook.cloud.microsoft` and decodes the address once on the
  way, so the folder ID is encoded twice to arrive the way Outlook expects.
  If Microsoft changes that forwarding, folder links open the inbox instead,
  and **Folder URL** needs changing; email links are not affected.
- Tested with Outlook on the web installed as an app from Edge. New Outlook
  for Windows is built on the same code but is untested. Classic Outlook for
  Windows drags a file instead, which this plugin ignores.
- Folder links are tested for folders in your own mailbox, not in shared
  mailboxes.
- An email moved to another mailbox, such as an online archive, gets a new
  ID, and its old link stops working.
- A link opens the email only for someone signed in to that mailbox. Links
  don't name the account, so with several Microsoft accounts in Edge, an
  email from another account than the app's may not open.

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
`versions.json` and commit, then push a tag such as `0.3.0`; the workflow
builds, tests and publishes the release.
