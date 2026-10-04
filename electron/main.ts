import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { autoUpdater } from "electron-updater";
import fs from "node:fs/promises";
import path from "node:path";

let win: BrowserWindow;
const rulesFile = () => path.join(app.getPath("userData"), "rules.json");

async function ensureRules() {
  try {
    await fs.access(rulesFile());
  } catch {
    const bundled = await fs.readFile(
      path.join(app.getAppPath(), "rules", "default-rules.json"),
      "utf8",
    );
    await fs.writeFile(rulesFile(), bundled, "utf8");
  }
}

function createWindow() {
  win = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  if (process.argv.includes("--dev")) win.loadURL("http://127.0.0.1:5173");
  else win.loadFile(path.join(app.getAppPath(), "dist", "index.html"));
}

app.whenReady().then(async () => {
  await ensureRules();
  ipcMain.handle("rules:load", () => fs.readFile(rulesFile(), "utf8"));
  ipcMain.handle("rules:save", (_e, content: string) =>
    fs.writeFile(rulesFile(), content, "utf8"),
  );
  ipcMain.handle("dialog:save", async (_e, defaultName: string) =>
    dialog.showSaveDialog(win, {
      defaultPath: defaultName,
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    }),
  );
  ipcMain.handle(
    "file:write",
    async (_e, filePath: string, bytes: Uint8Array) => {
      await fs.writeFile(filePath, Buffer.from(bytes));
      return true;
    },
  );
  ipcMain.handle("app:version", () => app.getVersion());
  createWindow();
});
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
