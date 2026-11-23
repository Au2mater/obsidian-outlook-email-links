// Tests for reading Outlook drags and inserting email links.
// Run with: node test/run.js
const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");

const stubPath = path.join(__dirname, "obsidian-stub.js");
const realResolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  if (request === "obsidian") return stubPath;
  return realResolve.call(this, request, ...rest);
};
const mainPath = path.join(__dirname, "..", "src", "plugin.js");
const source =
  fs.readFileSync(mainPath, "utf8") +
  "\nmodule.exports.__internals = { readOutlookDrag, formatItems, escapeLinkText, isOutlookType, DEFAULT_SETTINGS, parseDetectedApps, urlForApp, launchArgs, AppLauncher, DETECT_SCRIPT };\n";
const mod = new Module(mainPath, module);
mod.filename = mainPath;
mod.paths = Module._nodeModulePaths(path.dirname(mainPath));
mod._compile(source, mainPath);
const OutlookEmailLinksPlugin = mod.exports;
const { readOutlookDrag, formatItems, escapeLinkText, isOutlookType, DEFAULT_SETTINGS, parseDetectedApps, urlForApp, launchArgs, AppLauncher, DETECT_SCRIPT } = OutlookEmailLinksPlugin.__internals;

global.document = {};

// The shape Outlook on the web put on a drag of one message (values made up).
const mailbox = {
  sourceId: "x@y",
  type: "UserMailbox",
  mailboxSmtpAddress: "me@example.com",
  userIdentity: "me@example.com",
  mailboxRank: "Coprincipal",
  mailboxProvider: "Office365",
};
const onePayload = {
  tableViewId: "folderId:AAMk=;lVT:0;vF:All",
  itemType: "maillistitempart",
  subjects: ["RE: Budget [draft]"],
  itemIds: ["AAkALgAAAAAAHYQDEapmEc2byACqAC/EWg0Ad+Pn=="],
  nodeIds: ["<ABC123@EXAMPLE.PROD.OUTLOOK.COM>"],
  sizes: [114122],
  mailboxInfos: [mailbox],
  haveIRM: [false],
  itemData: { mailboxInfo: mailbox },
};

function transfer(data) {
  return {
    types: Object.keys(data),
    getData: (type) => (type in data ? data[type] : ""),
  };
}

function fakeEditor(text) {
  const editor = {
    text,
    cursor: { line: 0, ch: 0 },
    focused: false,
    posToOffset: (pos) => {
      const lines = editor.text.split("\n");
      let offset = 0;
      for (let i = 0; i < pos.line; i++) offset += lines[i].length + 1;
      return offset + pos.ch;
    },
    offsetToPos: (offset) => {
      const before = editor.text.slice(0, offset).split("\n");
      return { line: before.length - 1, ch: before[before.length - 1].length };
    },
    replaceRange(insert, pos) {
      const at = editor.posToOffset(pos);
      editor.text = editor.text.slice(0, at) + insert + editor.text.slice(at);
    },
    getCursor: () => editor.cursor,
    setCursor(pos) {
      editor.cursor = pos;
    },
    focus() {
      editor.focused = true;
    },
  };
  return editor;
}

function dropEvent(data, at) {
  return {
    dataTransfer: transfer(data),
    defaultPrevented: false,
    clientX: 1,
    clientY: 1,
    preventDefault() {
      this.defaultPrevented = true;
    },
    at,
  };
}

async function loadedPlugin(data) {
  global.__data = data || null;
  global.__notices = [];
  const plugin = new OutlookEmailLinksPlugin();
  await plugin.onload();
  return plugin;
}

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test("reads a dragged Outlook message", () => {
  const emails = readOutlookDrag(transfer({ maillistitempart: JSON.stringify(onePayload) }));
  assert.deepStrictEqual(emails, [
    {
      kind: "email",
      id: "AAkALgAAAAAAHYQDEapmEc2byACqAC/EWg0Ad+Pn==",
      subject: "RE: Budget [draft]",
      messageId: "<ABC123@EXAMPLE.PROD.OUTLOOK.COM>",
      mailbox: "me@example.com",
    },
  ]);
});

test("reads several dragged messages in order", () => {
  const payload = Object.assign({}, onePayload, {
    subjects: ["One", "Two"],
    itemIds: ["id1", "id2"],
    nodeIds: ["<a@x>", "<b@x>"],
    mailboxInfos: [mailbox, mailbox],
  });
  const emails = readOutlookDrag(transfer({ maillistitempart: JSON.stringify(payload) }));
  assert.deepStrictEqual(emails.map((e) => [e.id, e.subject, e.messageId]), [
    ["id1", "One", "<a@x>"],
    ["id2", "Two", "<b@x>"],
  ]);
});

test("ignores drags that are not from Outlook", () => {
  assert.deepStrictEqual(readOutlookDrag(transfer({ "text/plain": "hello" })), []);
  assert.deepStrictEqual(readOutlookDrag(transfer({ "text/plain": '{"a":1}' })), []);
  // JSON text that looks like Outlook's is still just text.
  assert.deepStrictEqual(readOutlookDrag(transfer({ "text/plain": JSON.stringify(onePayload) })), []);
  assert.deepStrictEqual(readOutlookDrag(transfer({ "text/plain": JSON.stringify(folderPayload) })), []);
  assert.deepStrictEqual(readOutlookDrag(transfer({ Files: "" })), []);
  assert.deepStrictEqual(readOutlookDrag(null), []);
});

test("an Outlook drag without item IDs is reported, not ignored", () => {
  const payload = { itemType: "maillistrow", subjects: ["Thread"], conversationIds: ["c1"] };
  assert.strictEqual(readOutlookDrag(transfer({ maillistrow: JSON.stringify(payload) })), null);
});

// The shape Outlook on the web put on a drag of a folder (values made up).
const folderPayload = {
  itemType: "mailfoldernode",
  folderId: "AAMkADQ5ZTAx+dq5KamS/AACXuXBmAAA=",
  displayName: "2606 Projects [old]",
  treeType: "primaryFolderTree",
  mailboxInfo: mailbox,
  parentFolderId: "AAMkParent=",
  itemData: { treeType: "primaryFolderTree", mailboxInfo: mailbox, draggedFolderId: "AAMkADQ5ZTAx+dq5KamS/AACXuXBmAAA=" },
};

test("reads a dragged Outlook folder", () => {
  assert.deepStrictEqual(readOutlookDrag(transfer({ mailfoldernode: JSON.stringify(folderPayload) })), [
    { kind: "folder", id: "AAMkADQ5ZTAx+dq5KamS/AACXuXBmAAA=", name: "2606 Projects [old]", mailbox: "me@example.com" },
  ]);
});

test("formats a folder link with an encoded ID and escaped name", () => {
  const items = readOutlookDrag(transfer({ mailfoldernode: JSON.stringify(folderPayload) }));
  assert.strictEqual(
    formatItems(items, DEFAULT_SETTINGS),
    "[📁 2606 Projects \\[old\\]](https://outlook.office365.com/mail/AAMkADQ5ZTAx%252Bdq5KamS%252FAACXuXBmAAA%253D)"
  );
});

test("a folder link keeps its double encoding when sent to the app", () => {
  const link = "https://outlook.office365.com/mail/AAMk%252Bdq5%253D";
  assert.strictEqual(urlForApp(link, foundApp), link);
});

test("a link copied from outlook.cloud.microsoft is encoded once more for the app", () => {
  assert.strictEqual(
    urlForApp("https://outlook.cloud.microsoft/mail/AAMk%2Bdq5%3D", foundApp),
    "https://outlook.office365.com/mail/AAMk%252Bdq5%253D"
  );
  const cloudApp = Object.assign({}, foundApp, { startUrl: "https://outlook.cloud.microsoft/mail/" });
  assert.strictEqual(urlForApp("https://outlook.cloud.microsoft/mail/AAMk%2Bdq5%3D", cloudApp), "https://outlook.cloud.microsoft/mail/AAMk%2Bdq5%3D");
});

test("dropping a folder inserts its link", async () => {
  const plugin = await loadedPlugin({ folderLinkTemplate: "{{name}} <{{url}}>" });
  const editor = fakeEditor("");
  const evt = dropEvent({ mailfoldernode: JSON.stringify(folderPayload) });
  plugin.events["editor-drop"](evt, editor);
  assert.ok(evt.defaultPrevented);
  assert.strictEqual(editor.text, "2606 Projects \\[old\\] <https://outlook.office365.com/mail/AAMkADQ5ZTAx%252Bdq5KamS%252FAACXuXBmAAA%253D>");
});

test("knows Outlook's drag type names", () => {
  assert.ok(isOutlookType("mailfoldernode"));
  assert.ok(isOutlookType("maillistitempart"));
  assert.ok(isOutlookType("multimaillistmessagerows"));
  assert.ok(!isOutlookType("text/plain"));
  assert.ok(!isOutlookType("Files"));
});

test("formats a link with an encoded ID and escaped subject", () => {
  const [email] = readOutlookDrag(transfer({ maillistitempart: JSON.stringify(onePayload) }));
  assert.strictEqual(
    formatItems([email], DEFAULT_SETTINGS),
    "[📧 RE: Budget \\[draft\\]](https://outlook.office365.com/mail/deeplink/read/AAkALgAAAAAAHYQDEapmEc2byACqAC%2FEWg0Ad%2BPn%3D%3D)"
  );
});

test("templates can use every placeholder", () => {
  const email = { id: "a/b", subject: "Hi", messageId: "<m@x>", mailbox: "me@example.com" };
  const settings = {
    urlTemplate: "https://outlook.cloud.microsoft/mail/{{mailbox}}/id/{{id}}",
    linkTemplate: "{{subject}} {{url}} {{id}} {{messageId}} {{mailbox}} {{unknown}}",
    separator: "\n",
  };
  assert.strictEqual(
    formatItems([email], settings),
    "Hi https://outlook.cloud.microsoft/mail/me%40example.com/id/a%2Fb a/b <m@x> me@example.com {{unknown}}"
  );
});

test("a missing or blank subject reads (No subject), and whitespace collapses", () => {
  const payload = Object.assign({}, onePayload, { subjects: ["  Line one\n line two "] });
  assert.strictEqual(readOutlookDrag(transfer({ maillistitempart: JSON.stringify(payload) }))[0].subject, "Line one line two");
  const blank = Object.assign({}, onePayload, { subjects: [] });
  assert.strictEqual(readOutlookDrag(transfer({ maillistitempart: JSON.stringify(blank) }))[0].subject, "(No subject)");
});

test("placeholders only come from the values, not from Object's prototype", () => {
  const settings = Object.assign({}, DEFAULT_SETTINGS, { linkTemplate: "{{constructor}} {{toString}} {{subject}}" });
  assert.strictEqual(formatItems([{ kind: "email", id: "1", subject: "Hi", messageId: "", mailbox: "" }], settings), "{{constructor}} {{toString}} Hi");
});

test("escapes brackets and backslashes in link text", () => {
  assert.strictEqual(escapeLinkText("a [b] \\c"), "a \\[b\\] \\\\c");
});

test("dropping an email inserts its link where it lands", async () => {
  const plugin = await loadedPlugin();
  const editor = fakeEditor("Before after");
  editor.cm = { posAtCoords: () => 7 };
  const evt = dropEvent({ maillistitempart: JSON.stringify(onePayload) });
  plugin.events["editor-drop"](evt, editor);
  assert.ok(evt.defaultPrevented);
  assert.ok(editor.text.startsWith("Before [📧 RE: Budget"));
  assert.ok(editor.text.endsWith("%3D%3D)after"));
  assert.strictEqual(editor.posToOffset(editor.cursor), editor.text.length - "after".length);
  assert.ok(editor.focused);
});

test("without a drop position the link goes in at the cursor", async () => {
  const plugin = await loadedPlugin({ linkTemplate: "<{{subject}}>" });
  const editor = fakeEditor("ab\ncd");
  editor.cursor = { line: 1, ch: 1 };
  plugin.events["editor-drop"](dropEvent({ maillistitempart: JSON.stringify(onePayload) }), editor);
  assert.strictEqual(editor.text, "ab\nc<RE: Budget \\[draft\\]>d");
});

test("several emails go in on their own lines", async () => {
  const plugin = await loadedPlugin({ linkTemplate: "- {{subject}}" });
  const payload = Object.assign({}, onePayload, { subjects: ["One", "Two"], itemIds: ["1", "2"] });
  const editor = fakeEditor("");
  plugin.events["editor-drop"](dropEvent({ maillistitempart: JSON.stringify(payload) }), editor);
  assert.strictEqual(editor.text, "- One\n- Two");
});

test("other drops are left to Obsidian", async () => {
  const plugin = await loadedPlugin();
  const editor = fakeEditor("x");
  const evt = dropEvent({ "text/plain": "hello" });
  plugin.events["editor-drop"](evt, editor);
  assert.ok(!evt.defaultPrevented);
  assert.strictEqual(editor.text, "x");
  const handled = dropEvent({ maillistitempart: JSON.stringify(onePayload) });
  handled.defaultPrevented = true;
  plugin.events["editor-drop"](handled, editor);
  assert.strictEqual(editor.text, "x");
});

test("an Outlook drop naming no email shows a notice and inserts nothing", async () => {
  const plugin = await loadedPlugin();
  const editor = fakeEditor("x");
  const evt = dropEvent({ maillistrow: JSON.stringify({ subjects: ["Thread"] }) });
  plugin.events["editor-drop"](evt, editor);
  assert.ok(evt.defaultPrevented);
  assert.strictEqual(editor.text, "x");
  assert.strictEqual(global.__notices.length, 1);
});

test("dragover over the editor accepts only Outlook drags", async () => {
  const plugin = await loadedPlugin();
  const { fn } = plugin.domEvents.find((d) => d.name === "dragover");
  const inEditor = { closest: (sel) => (sel === ".cm-editor" ? {} : null) };
  const outside = { closest: () => null };
  const over = (types, target) => {
    const evt = { target, dataTransfer: { types, dropEffect: "none" }, prevented: false, preventDefault() { this.prevented = true; } };
    fn(evt);
    return evt;
  };
  const yes = over(["maillistitempart"], inEditor);
  assert.ok(yes.prevented);
  assert.ok(!over(["text/plain"], inEditor).prevented);
  assert.ok(!over(["maillistitempart"], outside).prevented);
});

test("loading makes no network request and reads no secrets", async () => {
  const plugin = await loadedPlugin();
  assert.ok(!/requestUrl|fetch\(|secretStorage/.test(fs.readFileSync(mainPath, "utf8")));
  assert.ok(plugin.events["editor-drop"]);
});

// What Edge writes in the package manifest of an Outlook app it installed.
const detectedLine =
  "parameters?--app-id=pkooggnaalmfkidjmlhoelhdllpphaga --ip-edge-aumid=Microsoft.MicrosoftEdge.Stable_8wekyb3d8bbwe!MSEDGE;profile-directory?Default;start-url?https://outlook.office365.com/mail/;handlers?protocol_handlers";
const foundApp = { appId: "pkooggnaalmfkidjmlhoelhdllpphaga", profile: "Default", startUrl: "https://outlook.office365.com/mail/" };
const PROXY = "C:\\Edge\\msedge_proxy.exe";

function fakeSystem(options = {}) {
  const calls = { powershell: 0, launches: [] };
  const storage = Object.assign({}, options.storage);
  return {
    calls,
    storage,
    load: (key) => (key in storage ? JSON.parse(storage[key]) : null),
    save(key, value) {
      if (value) storage[key] = JSON.stringify(value);
      else delete storage[key];
    },
    edgeProxies: () => (options.noEdge ? [] : ["C:\\Missing\\msedge_proxy.exe", PROXY]),
    async powershell(script) {
      calls.powershell++;
      assert.strictEqual(script, DETECT_SCRIPT);
      return options.output === undefined ? "#< CLIXML\r\n" + detectedLine + "\r\n<Objs/>" : options.output;
    },
    async execFile(file, args) {
      if (options.launchFails) throw new Error("Access denied");
      if (file !== PROXY) throw Object.assign(new Error("not found"), { code: "ENOENT" });
      calls.launches.push([file, args]);
      return "";
    },
  };
}

test("reads the Outlook app from PowerShell's output", () => {
  assert.deepStrictEqual(parseDetectedApps("#< CLIXML\r\n" + detectedLine + "\r\n<Objs/>"), [foundApp]);
  assert.deepStrictEqual(parseDetectedApps(""), []);
  assert.deepStrictEqual(parseDetectedApps(detectedLine.replace("outlook.office365.com", "evil.example")), []);
  assert.deepStrictEqual(parseDetectedApps("parameters?--app-id=../x;start-url?https://outlook.office365.com/"), []);
});

test("moves a link to the app's host and builds the launch arguments", () => {
  const link = "https://outlook.office.com/mail/deeplink/read/AAk%2FA%3D?x=1";
  assert.strictEqual(urlForApp(link, foundApp), "https://outlook.office365.com/mail/deeplink/read/AAk%2FA%3D?x=1");
  assert.strictEqual(urlForApp("https://example.com/mail/", foundApp), null);
  assert.strictEqual(urlForApp("http://outlook.office.com/mail/", foundApp), null);
  assert.deepStrictEqual(launchArgs(foundApp, "https://outlook.office365.com/mail/"), [
    "--profile-directory=Default",
    "--app-id=pkooggnaalmfkidjmlhoelhdllpphaga",
    "--app-launch-url-for-shortcuts-menu-item=https://outlook.office365.com/mail/",
  ]);
});

test("the launcher finds the app once, saves it and opens links in it", async () => {
  const plugin = await loadedPlugin();
  const system = fakeSystem();
  const launcher = new AppLauncher(plugin, system);
  assert.strictEqual(await launcher.open("https://outlook.office.com/mail/deeplink/read/A"), true);
  assert.strictEqual(await launcher.open("https://outlook.cloud.microsoft/mail/deeplink/read/B"), true);
  assert.strictEqual(system.calls.powershell, 1);
  assert.deepStrictEqual(JSON.parse(system.storage["outlook-email-links:app"]), Object.assign({}, foundApp, { proxy: PROXY }));
  assert.ok(!global.__saved || !("app" in global.__saved), "the app is not written to data.json");
  assert.deepStrictEqual(system.calls.launches.map(([file, args]) => [file, args[2]]), [
    [PROXY, "--app-launch-url-for-shortcuts-menu-item=https://outlook.office365.com/mail/deeplink/read/A"],
    [PROXY, "--app-launch-url-for-shortcuts-menu-item=https://outlook.office365.com/mail/deeplink/read/B"],
  ]);
});

test("without the app, links go to the browser and PowerShell runs once", async () => {
  const plugin = await loadedPlugin();
  const system = fakeSystem({ output: "" });
  const launcher = new AppLauncher(plugin, system);
  assert.strictEqual(await launcher.open("https://outlook.office.com/mail/x"), false);
  assert.strictEqual(await launcher.open("https://outlook.office.com/mail/y"), false);
  assert.strictEqual(system.calls.powershell, 1);
  const noEdge = fakeSystem({ noEdge: true });
  assert.strictEqual(await new AppLauncher(plugin, noEdge).open("https://outlook.office.com/mail/x"), false);
  assert.strictEqual(noEdge.calls.launches.length, 0);
});

const remembered = (app) => ({ "outlook-email-links:app": JSON.stringify(app) });

test("a failed launch forgets the app and lets the browser open the link", async () => {
  const plugin = await loadedPlugin();
  const system = fakeSystem({ launchFails: true, storage: remembered(Object.assign({}, foundApp, { proxy: PROXY })) });
  assert.strictEqual(await new AppLauncher(plugin, system).open("https://outlook.office.com/mail/x"), false);
  assert.strictEqual(system.calls.powershell, 0);
  assert.ok(!("outlook-email-links:app" in system.storage));
});

test("a remembered program outside Edge's folders is never run", async () => {
  const plugin = await loadedPlugin();
  const system = fakeSystem({ storage: remembered(Object.assign({}, foundApp, { proxy: "C:\\evil.exe" })) });
  assert.strictEqual(await new AppLauncher(plugin, system).open("https://outlook.office.com/mail/x"), true);
  assert.deepStrictEqual(system.calls.launches.map(([file]) => file), [PROXY]);
  assert.strictEqual(system.calls.powershell, 0);
});

test("a malformed remembered app is ignored and the app looked for again", async () => {
  const plugin = await loadedPlugin();
  for (const bad of [{ appId: "x", profile: "Default", startUrl: foundApp.startUrl }, Object.assign({}, foundApp, { startUrl: "https://evil.example/mail/" }), Object.assign({}, foundApp, { profile: 'a" --x' })]) {
    const system = fakeSystem({ storage: remembered(bad) });
    assert.strictEqual(await new AppLauncher(plugin, system).open("https://outlook.office.com/mail/x"), true);
    assert.strictEqual(system.calls.powershell, 1);
    assert.strictEqual(system.calls.launches[0][1][1], "--app-id=pkooggnaalmfkidjmlhoelhdllpphaga");
  }
});

test("an app kept in data.json by an older version is dropped", async () => {
  global.__saved = undefined;
  const plugin = await loadedPlugin({ app: Object.assign({}, foundApp, { proxy: "C:\\evil.exe" }), urlTemplate: DEFAULT_SETTINGS.urlTemplate, linkTemplate: "- {{subject}}" });
  assert.ok(!("app" in plugin.settings));
  assert.deepStrictEqual(global.__saved, { linkTemplate: "- {{subject}}" }, "only changed settings are kept");
});

test("on Windows, window.open sends Outlook links to the app and others on", async () => {
  const opened = [];
  global.window = { open: (url) => opened.push(url) };
  const stub = require(stubPath);
  stub.Platform.isWin = true;
  try {
    const plugin = await loadedPlugin();
    plugin.launcher.system = fakeSystem();
    window.open("https://example.com/");
    window.open("https://outlook.office.com/mail/deeplink/read/A");
    await new Promise((r) => setImmediate(r));
    assert.deepStrictEqual(opened, ["https://example.com/"]);
    assert.strictEqual(plugin.launcher.system.calls.launches.length, 1);

    plugin.launcher.system = fakeSystem({ output: "" });
    window.open("https://outlook.office.com/mail/deeplink/read/B");
    await new Promise((r) => setTimeout(r, 5));
    assert.deepStrictEqual(opened, ["https://example.com/", "https://outlook.office.com/mail/deeplink/read/B"]);

    for (const other of ["https://outlook.office365.com/owa/?ItemID=X", "https://outlook.office.com/calendar/item/X", "https://outlook.live.com/mail/0/id/X"]) {
      window.open(other);
      assert.strictEqual(opened[opened.length - 1], other);
    }
    opened.length = 2;

    plugin.settings.openInApp = false;
    window.open("https://outlook.office.com/mail/C");
    assert.strictEqual(opened.length, 3);
    plugin.cleanups.forEach((fn) => fn());
    assert.strictEqual(window.open.name, "open");

    // Another plugin's later wrapper is left in place on unload.
    const again = await loadedPlugin();
    const wrapped = window.open;
    const theirs = (url) => wrapped(url);
    window.open = theirs;
    again.cleanups.forEach((fn) => fn());
    assert.strictEqual(window.open, theirs);
    // …and the unloaded wrapper underneath passes links straight through.
    opened.length = 0;
    window.open("https://outlook.office.com/mail/D");
    assert.deepStrictEqual(opened, ["https://outlook.office.com/mail/D"]);

    // An unexpected error while opening the app still opens the browser.
    window.open = (url) => opened.push(url);
    const failing = await loadedPlugin();
    failing.launcher.open = async () => {
      throw new Error("boom");
    };
    opened.length = 0;
    window.open("https://outlook.office.com/mail/E");
    await new Promise((r) => setImmediate(r));
    assert.deepStrictEqual(opened, ["https://outlook.office.com/mail/E"]);
  } finally {
    stub.Platform.isWin = false;
  }
});

(async () => {
  let failed = 0;
  for (const [name, fn] of tests) {
    try {
      await fn();
      console.log(`ok   ${name}`);
    } catch (e) {
      failed++;
      console.log(`FAIL ${name}\n     ${e.stack.split("\n").slice(0, 3).join("\n     ")}`);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed`);
  process.exit(failed ? 1 : 0);
})();
