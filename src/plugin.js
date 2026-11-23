// Outlook Email Links: drag an email from Outlook on the web (or its PWA)
// onto a note, and a Markdown link to that email is inserted where it lands.
// Clicking the link opens the email in Outlook.
//
// Outlook puts a JSON description of the dragged messages on the drag, under
// a custom type such as "maillistitempart". It holds each message's subject,
// its Exchange immutable item ID (which survives moving the email to another
// folder), its Internet message ID and the mailbox. The plugin reads only
// that; it never signs in, makes no network request and copies no email
// content into the vault beyond what the link template asks for.
//
// On Windows, clicking an Outlook link opens it in the Outlook app installed
// from Edge (the PWA) when there is one, instead of a browser tab. Edge
// registers such apps as Windows packages; the plugin finds Outlook's once,
// with PowerShell, on the first click, and launches it at the link through
// Edge's msedge_proxy.exe, the way the app's own shortcuts do.
//
// Plain JavaScript with no dependencies: `npm run build` checks this file and
// writes it unchanged to main.js, the file Obsidian loads (see README.md).

const { Notice, Platform, Plugin, PluginSettingTab, Setting } = require("obsidian");

const DEFAULT_SETTINGS = {
  // {{id}} is the URL-encoded item ID; {{rawId}} is the item ID as is.
  urlTemplate: "https://outlook.office365.com/mail/deeplink/read/{{id}}",
  linkTemplate: "[📧 {{subject}}]({{url}})",
  separator: "\n",
  // A dragged folder; {{id}} is the folder ID encoded twice (see folderPathId).
  folderUrlTemplate: "https://outlook.office365.com/mail/{{id}}",
  folderLinkTemplate: "[📁 {{name}}]({{url}})",
  openInApp: true,
  // The Outlook app found on this computer: { appId, profile, startUrl, proxy }.
  app: null,
};

// Work and school Outlook's addresses. Only mail links on these are opened in
// the app; calendar links and personal Outlook (outlook.live.com) are not.
const OUTLOOK_HOSTS = ["outlook.office365.com", "outlook.office.com", "outlook.cloud.microsoft"];

// Prints, for each Outlook app Edge installed, the line Edge keeps in the
// app's package manifest: "parameters?--app-id=…;profile-directory?…;start-url?…".
const DETECT_SCRIPT = [
  "$ProgressPreference = 'SilentlyContinue'",
  'Get-AppxPackage | Where-Object { $_.Name -like "outlook.*" } | ForEach-Object {',
  '  $m = Join-Path $_.InstallLocation "AppxManifest.xml"',
  "  if (Test-Path $m) {",
  "    $x = Get-Content -Raw $m",
  "    if ($x -match 'HostId=\"PWA\"' -and $x -match 'Description=\"(parameters\\?[^\"]*)\"') { $Matches[1] }",
  "  }",
  "}",
].join("\n");

// Outlook names its drag types after what is dragged: "maillistitempart" for
// emails from the message list and "mailfoldernode" for a folder. This also
// covers related ones it may use.
const OUTLOOK_TYPE = /^(multi)?maillist|^mailfolder/i;

function isOutlookType(type) {
  return typeof type === "string" && OUTLOOK_TYPE.test(type);
}

// Reads the dragged emails, or folder, from a DataTransfer, as
// { kind: "email" | "folder", … }. Returns [] when the drag is not from
// Outlook, or null when it is from Outlook but names no email or folder.
function readOutlookDrag(dataTransfer) {
  if (!dataTransfer) return [];
  const types = Array.from(dataTransfer.types || []);
  let sawOutlook = false;
  for (const type of types) {
    const outlookType = isOutlookType(type);
    if (!outlookType && type !== "text/plain") continue;
    let raw;
    try {
      raw = dataTransfer.getData(type);
    } catch (e) {
      continue;
    }
    const payload = parseJson(raw);
    if (!payload || typeof payload !== "object") continue;
    if (outlookType || Array.isArray(payload.itemIds) || Array.isArray(payload.subjects)) sawOutlook = true;
    const items = folderFromPayload(payload) || emailsFromPayload(payload);
    if (items.length) return items;
  }
  return sawOutlook ? null : [];
}

function parseJson(raw) {
  if (typeof raw !== "string" || !raw.trim().startsWith("{")) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function folderFromPayload(payload) {
  const id = typeof payload.folderId === "string" ? payload.folderId.trim() : "";
  if (!id || Array.isArray(payload.itemIds)) return null;
  const name = typeof payload.displayName === "string" ? payload.displayName.replace(/\s+/g, " ").trim() : "";
  return [{ kind: "folder", id, name: name || "(Folder)", mailbox: mailboxAddress(payload.mailboxInfo) }];
}

function mailboxAddress(info) {
  return info && typeof info.mailboxSmtpAddress === "string" ? info.mailboxSmtpAddress : "";
}

function emailsFromPayload(payload) {
  const ids = Array.isArray(payload.itemIds) ? payload.itemIds : [];
  const subjects = Array.isArray(payload.subjects) ? payload.subjects : [];
  const messageIds = Array.isArray(payload.nodeIds) ? payload.nodeIds : [];
  const mailboxes = Array.isArray(payload.mailboxInfos) ? payload.mailboxInfos : [];
  const fallbackMailbox = payload.itemData && payload.itemData.mailboxInfo;
  const emails = [];
  ids.forEach((id, i) => {
    if (typeof id !== "string" || !id.trim()) return;
    const mailbox = mailboxes[i] || fallbackMailbox || {};
    const messageId = typeof messageIds[i] === "string" && messageIds[i].startsWith("<") ? messageIds[i] : "";
    emails.push({
      kind: "email",
      id: id.trim(),
      subject: cleanSubject(subjects[i]),
      messageId,
      mailbox: mailboxAddress(mailbox),
    });
  });
  return emails;
}

function cleanSubject(subject) {
  const text = typeof subject === "string" ? subject.replace(/\s+/g, " ").trim() : "";
  return text || "(No subject)";
}

// Escapes what would break the text of a Markdown link.
function escapeLinkText(text) {
  return text.replace(/([\\\[\]])/g, "\\$1");
}

function fill(template, values) {
  return template.replace(/\{\{(\w+)\}\}/g, (match, key) => (key in values ? values[key] : match));
}

function emailUrl(email, settings) {
  return fill(settings.urlTemplate || DEFAULT_SETTINGS.urlTemplate, {
    id: encodeURIComponent(email.id),
    rawId: email.id,
    mailbox: encodeURIComponent(email.mailbox),
    messageId: encodeURIComponent(email.messageId),
  });
}

function formatEmail(email, settings) {
  const url = emailUrl(email, settings);
  return fill(settings.linkTemplate || DEFAULT_SETTINGS.linkTemplate, {
    subject: escapeLinkText(email.subject),
    url,
    id: email.id,
    messageId: email.messageId,
    mailbox: email.mailbox,
  });
}

// Outlook opens a folder at /mail/<encoded folder ID> on outlook.cloud.microsoft
// ("=" written %3D). The Outlook app lives on outlook.office365.com, which
// forwards there but decodes the path once on the way, and Outlook shows the
// inbox for a folder ID with a bare "=". Encoding the ID twice survives that
// forwarding. Tested in November 2026; if Microsoft changes the forwarding,
// change Folder URL in the settings.
function folderPathId(id) {
  return encodeURIComponent(encodeURIComponent(id));
}

function formatFolder(folder, settings) {
  const url = fill(settings.folderUrlTemplate || DEFAULT_SETTINGS.folderUrlTemplate, {
    id: folderPathId(folder.id),
    rawId: folder.id,
    mailbox: encodeURIComponent(folder.mailbox),
  });
  return fill(settings.folderLinkTemplate || DEFAULT_SETTINGS.folderLinkTemplate, {
    name: escapeLinkText(folder.name),
    url,
    id: folder.id,
    mailbox: folder.mailbox,
  });
}

function formatItems(items, settings) {
  const separator = settings.separator === undefined ? DEFAULT_SETTINGS.separator : settings.separator;
  return items.map((item) => (item.kind === "folder" ? formatFolder(item, settings) : formatEmail(item, settings))).join(separator);
}

// Where in the editor a drop at (x, y) lands, falling back to the cursor.
function dropPosition(editor, evt) {
  const view = editor.cm;
  if (view && typeof view.posAtCoords === "function") {
    const offset = view.posAtCoords({ x: evt.clientX, y: evt.clientY });
    if (offset !== null && offset !== undefined) return editor.offsetToPos(offset);
  }
  return editor.getCursor();
}

// Reads the apps from DETECT_SCRIPT's output. Edge's own field separators are
// ";" and "?", and "&amp;" may appear because the line comes from XML.
function parseDetectedApps(output) {
  const apps = [];
  for (const line of String(output || "").split(/\r?\n/)) {
    const fields = {};
    for (const part of line.trim().replace(/&amp;/g, "&").split(";")) {
      const at = part.indexOf("?");
      if (at > 0) fields[part.slice(0, at)] = part.slice(at + 1);
    }
    const appId = /--app-id=([a-p]{32})\b/.exec(fields.parameters || "");
    const startUrl = safeOutlookUrl(fields["start-url"]);
    if (appId && startUrl) apps.push({ appId: appId[1], profile: fields["profile-directory"] || "Default", startUrl: startUrl.href });
  }
  return apps;
}

// An https mail URL on a work Outlook host, or null.
function safeOutlookUrl(value) {
  let url;
  try {
    url = new URL(value);
  } catch (e) {
    return null;
  }
  return url.protocol === "https:" && OUTLOOK_HOSTS.includes(url.hostname.toLowerCase()) && url.pathname.startsWith("/mail/") ? url : null;
}

// The link moved to the app's own host, so it opens inside the app's window
// rather than in one with a "this page is outside the app" bar.
function urlForApp(link, app) {
  const url = safeOutlookUrl(link);
  if (!url) return null;
  const start = new URL(app.startUrl);
  url.protocol = start.protocol;
  url.host = start.host;
  return url.href;
}

function launchArgs(app, url) {
  return [`--profile-directory=${app.profile}`, `--app-id=${app.appId}`, `--app-launch-url-for-shortcuts-menu-item=${url}`];
}

// Opens Outlook links in the installed Outlook app on Windows.
class AppLauncher {
  constructor(plugin, system) {
    this.plugin = plugin;
    this.system = system;
    this.missing = false; // no app found this session; asked again on request
  }

  // Opens the link in the app and returns true, or returns false to let the
  // browser open it. Finding the app the first time takes a second or two.
  async open(link) {
    if (!safeOutlookUrl(link)) return false;
    const app = await this.app();
    const url = app && urlForApp(link, app);
    if (!url) return false;
    // Edge's launcher is in one of a few known places; the first that starts
    // is remembered. No file is read to find it.
    for (const proxy of app.proxy ? [app.proxy] : this.system.edgeProxies()) {
      try {
        await this.system.execFile(proxy, launchArgs(app, url));
        if (app.proxy !== proxy) {
          app.proxy = proxy;
          await this.plugin.saveSettings();
        }
        return true;
      } catch (e) {
        if (e && e.code === "ENOENT") continue;
        console.warn("Outlook Email Links: could not open the Outlook app", e);
        break;
      }
    }
    this.plugin.settings.app = null;
    await this.plugin.saveSettings();
    return false;
  }

  async app() {
    const saved = this.plugin.settings.app;
    if (saved && saved.appId) return saved;
    if (this.missing) return null;
    return this.detect();
  }

  async detect() {
    let app = null;
    try {
      [app = null] = parseDetectedApps(await this.system.powershell(DETECT_SCRIPT));
    } catch (e) {
      console.warn("Outlook Email Links: could not look for the Outlook app", e);
    }
    this.missing = !app;
    this.plugin.settings.app = app;
    await this.plugin.saveSettings();
    return app;
  }
}

// What the launcher needs from Windows, kept apart so tests can stand in.
// It runs two programs and nothing else: PowerShell, once, to find the app,
// and Edge's msedge_proxy.exe to open a link in it. No shell is involved, so
// a link is passed as one argument and cannot run commands.
function windowsSystem() {
  const childProcess = require("child_process");
  const run = (file, args) =>
    new Promise((resolve, reject) => {
      childProcess.execFile(file, args, { windowsHide: true, timeout: 20000 }, (error, stdout) => (error ? reject(error) : resolve(stdout)));
    });
  return {
    edgeProxies: () =>
      [process.env["ProgramFiles(x86)"], process.env.ProgramFiles, process.env.LOCALAPPDATA]
        .filter(Boolean)
        .map((root) => `${root}\\Microsoft\\Edge\\Application\\msedge_proxy.exe`),
    // Encoded, so quotes in the script survive Windows' command-line rules.
    powershell: (script) =>
      run("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", Buffer.from(script, "utf16le").toString("base64")]),
    execFile: (file, args) => run(file, args),
  };
}

class OutlookEmailLinksPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    this.addSettingTab(new OutlookEmailLinksSettingTab(this.app, this));

    // Let an Outlook drag drop on the editor at all: without this the drop
    // event may never fire, since the drag carries no text or files.
    this.registerDomEvent(document, "dragover", (evt) => {
      const dt = evt.dataTransfer;
      if (!dt || !Array.from(dt.types || []).some(isOutlookType)) return;
      const target = evt.target;
      if (!target || typeof target.closest !== "function" || !target.closest(".cm-editor")) return;
      evt.preventDefault();
    });

    this.registerEvent(this.app.workspace.on("editor-drop", (evt, editor) => this.handleDrop(evt, editor)));

    if (Platform.isDesktopApp && Platform.isWin) this.catchOutlookLinks(windowsSystem());
  }

  // Obsidian opens every external link, in any view, through window.open.
  // Outlook links go to the app instead when it is installed; anything else,
  // and Outlook links when it is not, go on to window.open as before.
  catchOutlookLinks(system) {
    this.launcher = new AppLauncher(this, system);
    const original = window.open;
    const plugin = this;
    window.open = function (url, ...rest) {
      const link = typeof url === "string" ? url : url && url.href;
      if (!plugin.settings.openInApp || !safeOutlookUrl(link)) return original.call(this, url, ...rest);
      plugin.launcher.open(link).then((opened) => {
        if (!opened) original.call(window, url, ...rest);
      });
      return null;
    };
    this.register(() => {
      window.open = original;
    });
  }

  handleDrop(evt, editor) {
    if (evt.defaultPrevented) return;
    const items = readOutlookDrag(evt.dataTransfer);
    if (items === null) {
      evt.preventDefault();
      new Notice("Outlook didn't say which email or folder this is. Drag a single message rather than a conversation.");
      return;
    }
    if (!items.length) return;
    evt.preventDefault();
    const text = formatItems(items, this.settings);
    const pos = dropPosition(editor, evt);
    editor.replaceRange(text, pos);
    const end = editor.offsetToPos(editor.posToOffset(pos) + text.length);
    editor.setCursor(end);
    editor.focus();
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }
}

class OutlookEmailLinksSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    this.text(
      "Link template",
      "What a dropped email becomes. Placeholders: {{subject}}, {{url}}, {{id}}, {{messageId}}, {{mailbox}}.",
      "linkTemplate"
    );
    this.text(
      "Email URL",
      "The address a link opens. {{id}} is the email's item ID. Use https://outlook.cloud.microsoft/… if that is where your Outlook lives.",
      "urlTemplate"
    );
    this.text("Folder link template", "What a dropped folder becomes. Placeholders: {{name}}, {{url}}, {{id}}, {{mailbox}}.", "folderLinkTemplate");
    this.text("Folder URL", "The address a folder link opens. {{id}} is the folder's ID.", "folderUrlTemplate");
    this.text("Between emails", "What goes between links when you drag several emails at once. \\n is a new line.", "separator");
    if (this.plugin.launcher) this.appSetting();
  }

  appSetting() {
    const app = this.plugin.settings.app;
    new Setting(this.containerEl)
      .setName("Open links in the Outlook app")
      .setDesc(
        app
          ? `Outlook links open in the Outlook app installed from Edge (${new URL(app.startUrl).host}).`
          : "Outlook links open in the Outlook app installed from Edge, when there is one. It is looked for on the first click."
      )
      .addButton((button) =>
        button.setButtonText("Find app").onClick(async () => {
          const found = await this.plugin.launcher.detect();
          new Notice(found ? "Found the Outlook app." : "No Outlook app installed from Edge was found; links open in the browser.");
          this.display();
        })
      )
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.openInApp).onChange(async (value) => {
          this.plugin.settings.openInApp = value;
          await this.plugin.saveSettings();
        })
      );
  }

  text(name, desc, key) {
    new Setting(this.containerEl)
      .setName(name)
      .setDesc(desc)
      .addText((text) =>
        text
          .setPlaceholder(DEFAULT_SETTINGS[key].replace(/\n/g, "\\n"))
          .setValue(this.plugin.settings[key].replace(/\n/g, "\\n"))
          .onChange(async (value) => {
            this.plugin.settings[key] = key === "separator" ? value.replace(/\\n/g, "\n") : value.trim() || DEFAULT_SETTINGS[key];
            await this.plugin.saveSettings();
          })
      );
  }
}

module.exports = OutlookEmailLinksPlugin;
