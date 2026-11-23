// Outlook Email Links: drag an email or a folder from Outlook on the web (or
// its PWA) onto a note, and a Markdown link to it is inserted where it lands.
// Clicking the link opens the email or folder in Outlook.
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
// with PowerShell, on the first click, remembers it on this device, and
// launches it at the link through Edge's msedge_proxy.exe, the way the app's
// own shortcuts do.
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
};

// Where the Outlook app found on this computer is remembered:
// { appId, profile, startUrl, proxy }. It belongs to the device, not the
// vault, so it is kept in the app's local storage rather than in data.json,
// which syncs with the vault to other computers.
const APP_STORAGE_KEY = "outlook-email-links:app";

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
    if (!isOutlookType(type)) continue;
    sawOutlook = true;
    let raw;
    try {
      raw = dataTransfer.getData(type);
    } catch (e) {
      continue;
    }
    const payload = parseJson(raw);
    if (!payload || typeof payload !== "object") continue;
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
  return template.replace(/\{\{(\w+)\}\}/g, (match, key) => (Object.prototype.hasOwnProperty.call(values, key) ? values[key] : match));
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
// rather than in one with a "this page is outside the app" bar. A link on
// outlook.cloud.microsoft, such as one copied from Outlook's address bar, is
// encoded once more, because the app's host forwards back there and decodes
// the path once on the way (see folderPathId).
function urlForApp(link, app) {
  const url = safeOutlookUrl(link);
  if (!url) return null;
  const start = new URL(app.startUrl);
  if (url.hostname.toLowerCase() === "outlook.cloud.microsoft" && start.hostname.toLowerCase() !== "outlook.cloud.microsoft") {
    url.pathname = url.pathname.replace(/%/g, "%25");
  }
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
    // is remembered. No file is read to find it, and only those places are
    // ever run.
    const places = this.system.edgeProxies();
    for (const proxy of places.includes(app.proxy) ? [app.proxy] : places) {
      try {
        await this.system.execFile(proxy, launchArgs(app, url));
        if (app.proxy !== proxy) this.remember(Object.assign({}, app, { proxy }));
        return true;
      } catch (e) {
        if (e && e.code === "ENOENT") continue;
        console.warn("Outlook Email Links: could not open the Outlook app", e);
        break;
      }
    }
    this.remember(null);
    return false;
  }

  // The app found earlier on this computer, else a new search, once a session.
  async app() {
    const saved = this.saved();
    if (saved) return saved;
    if (this.missing) return null;
    return this.detect();
  }

  saved() {
    return validApp(this.system.load(APP_STORAGE_KEY));
  }

  async detect() {
    let app = null;
    try {
      [app = null] = parseDetectedApps(await this.system.powershell(DETECT_SCRIPT));
    } catch (e) {
      console.warn("Outlook Email Links: could not look for the Outlook app", e);
    }
    this.missing = !app;
    this.remember(app);
    return app;
  }

  remember(app) {
    this.system.save(APP_STORAGE_KEY, app);
  }
}

// A remembered app, checked the way a freshly found one is, or null.
function validApp(app) {
  if (!app || typeof app !== "object") return null;
  if (typeof app.appId !== "string" || !/^[a-p]{32}$/.test(app.appId)) return null;
  if (typeof app.profile !== "string" || !/^[\w .-]{1,64}$/.test(app.profile)) return null;
  if (!safeOutlookUrl(app.startUrl)) return null;
  return { appId: app.appId, profile: app.profile, startUrl: app.startUrl, proxy: typeof app.proxy === "string" ? app.proxy : "" };
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
    load(key) {
      try {
        return JSON.parse(window.localStorage.getItem(key) || "null");
      } catch (e) {
        return null;
      }
    },
    save(key, value) {
      try {
        if (value) window.localStorage.setItem(key, JSON.stringify(value));
        else window.localStorage.removeItem(key);
      } catch (e) {
        // Without storage the app is looked for again next session.
      }
    },
  };
}

class OutlookEmailLinksPlugin extends Plugin {
  async onload() {
    const data = (await this.loadData()) || {};
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
    // Before 0.3.0 the app found was kept in data.json; it now stays on the
    // device. Older versions also saved every default.
    delete this.settings.app;
    if ("app" in data) await this.saveSettings();
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
    let active = true;
    const patched = function (url, ...rest) {
      const link = typeof url === "string" ? url : url && url.href;
      if (!active || !plugin.settings.openInApp || !safeOutlookUrl(link)) return original.call(this, url, ...rest);
      const browser = () => original.call(window, url, ...rest);
      plugin.launcher.open(link).then(
        (opened) => opened || browser(),
        (e) => {
          console.warn("Outlook Email Links: could not open the Outlook app", e);
          browser();
        }
      );
      return null;
    };
    window.open = patched;
    // Put window.open back. If another plugin has wrapped it since, leave
    // theirs in place and let this wrapper pass every link straight through.
    this.register(() => {
      active = false;
      if (window.open === patched) window.open = original;
    });
  }

  handleDrop(evt, editor) {
    if (evt.defaultPrevented) return;
    const items = readOutlookDrag(evt.dataTransfer);
    if (items === null) {
      evt.preventDefault();
      new Notice("Outlook didn't say which email or folder this is, so no link was added.");
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

  // Saves only what differs from the defaults, so a later improvement to a
  // default reaches everyone who hasn't changed that setting.
  async saveSettings() {
    const changed = {};
    for (const key of Object.keys(DEFAULT_SETTINGS)) {
      if (this.settings[key] !== DEFAULT_SETTINGS[key]) changed[key] = this.settings[key];
    }
    await this.saveData(changed);
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
      "The address an email link opens. {{id}} is the email's item ID.",
      "urlTemplate"
    );
    this.text("Folder link template", "What a dropped folder becomes. Placeholders: {{name}}, {{url}}, {{id}}, {{mailbox}}.", "folderLinkTemplate");
    this.text("Folder URL", "The address a folder link opens. {{id}} is the folder's ID.", "folderUrlTemplate");
    this.text("Between emails", "What goes between links when you drag several emails at once. \\n is a new line.", "separator");
    if (this.plugin.launcher) this.appSetting();
  }

  appSetting() {
    const app = this.plugin.launcher.saved();
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
          .setValue(String(this.plugin.settings[key] ?? DEFAULT_SETTINGS[key]).replace(/\n/g, "\\n"))
          .onChange(async (value) => {
            this.plugin.settings[key] = key === "separator" ? value.replace(/\\n/g, "\n") : value.trim() || DEFAULT_SETTINGS[key];
            await this.plugin.saveSettings();
          })
      );
  }
}

module.exports = OutlookEmailLinksPlugin;
