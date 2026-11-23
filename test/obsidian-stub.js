// A stand-in for the "obsidian" module, enough to load the plugin in Node.
class Plugin {
  constructor() {
    this.events = {};
    this.domEvents = [];
    this.app = {
      workspace: {
        on: (name, fn) => {
          this.events[name] = fn;
          return { name };
        },
      },
    };
  }
  registerEvent() {}
  register(fn) {
    (this.cleanups = this.cleanups || []).push(fn);
  }
  registerDomEvent(target, name, fn) {
    this.domEvents.push({ target, name, fn });
  }
  addSettingTab() {}
  async loadData() {
    return global.__data || null;
  }
  async saveData(data) {
    global.__saved = data;
  }
}

module.exports = {
  Notice: class {
    constructor(message) {
      (global.__notices = global.__notices || []).push(message);
    }
  },
  Platform: { isDesktopApp: true, isWin: false },
  Plugin,
  PluginSettingTab: class {},
  Setting: class {},
};
