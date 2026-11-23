// Builds the plugin: checks src/plugin.js parses and writes it, unchanged, to
// main.js, the file Obsidian loads. The plugin is plain JavaScript with no
// dependencies, so the build is a checked copy and its output is reproducible
// byte for byte. The source is not called main.js because build checks remove
// every file named like a release asset before building.
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const root = path.join(__dirname, "..");
const source = fs.readFileSync(path.join(root, "src", "plugin.js"));
new vm.Script(source.toString("utf8"), { filename: "src/plugin.js" }); // throws on a syntax error
fs.writeFileSync(path.join(root, "main.js"), source);
console.log(`Built main.js (${source.length} bytes)`);
