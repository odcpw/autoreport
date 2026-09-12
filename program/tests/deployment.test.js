const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const { spawn, spawnSync } = require("node:child_process");
const { ROOT, loadBrowserScripts } = require("./helpers");

const repo = path.dirname(ROOT);
const windows = process.platform === "win32";
const powershell = windows ? "powershell.exe" : process.env.AUTOBERICHT_POWERSHELL;

const run = (command, args, options) => new Promise((resolve, reject) => {
  const child = spawn(command, args, options);
  let output = "";
  child.stdout?.on("data", (chunk) => { output += chunk; });
  child.stderr?.on("data", (chunk) => { output += chunk; });
  child.on("error", reject);
  child.on("close", (code) => code === 0 ? resolve(output) : reject(new Error(`${command} exited ${code}: ${output}`)));
});

test("keep-only-sync installation, launcher and bundled assets work from a folder with spaces", {
  skip: !powershell && "Windows PowerShell or AUTOBERICHT_POWERSHELL required",
  timeout: 120000,
}, async (t) => {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "AutoBericht install test "));
  const installation = path.join(temporary, "Program files");
  const project = path.join(temporary, "Client project");
  fs.mkdirSync(installation);
  fs.mkdirSync(project);
  fs.writeFileSync(path.join(project, "project_sidecar.json"), '{"private":"unchanged"}');
  fs.copyFileSync(path.join(repo, "sync-autobericht.ps1"), path.join(installation, "sync-autobericht.ps1"));

  // Build the same top-level archive shape as GitHub, from the current tracked
  // files. This also tests uncommitted edits locally, without downloading main.
  const tracked = spawnSync("git", ["ls-files", "-z"], { cwd: repo, encoding: "utf8" });
  assert.equal(tracked.status, 0, tracked.stderr);
  const zip = loadBrowserScripts(["mini/shared/word-docx-zip.js"]).AutoBerichtWordDocxZip;
  const archive = zip.buildZipStore(tracked.stdout.split("\0").filter(Boolean).map((name) => ({
    name: `autoreport-main/${name}`,
    data: new Uint8Array(fs.readFileSync(path.join(repo, name))),
  })));
  const downloadServer = http.createServer((request, response) => {
    response.writeHead(200, { "Content-Type": "application/zip" });
    response.end(archive);
  });
  await new Promise((resolve) => downloadServer.listen(0, "127.0.0.1", resolve));
  let launcher;
  t.after(async () => {
    if (launcher && launcher.exitCode === null) {
      if (windows) await run("taskkill", ["/PID", String(launcher.pid), "/T", "/F"], {}).catch(() => {});
      else launcher.kill("SIGTERM");
      await new Promise((resolve) => {
        if (launcher.exitCode !== null) resolve();
        else launcher.once("close", resolve);
      });
    }
    downloadServer.closeAllConnections();
    await new Promise((resolve) => downloadServer.close(resolve));
    fs.rmSync(temporary, { recursive: true, force: true });
  });
  const env = { ...process.env, TEMP: temporary };
  await run(powershell, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
    path.join(installation, "sync-autobericht.ps1"), "-RepoArchiveUrl",
    `http://127.0.0.1:${downloadServer.address().port}/main.zip`,
  ], { cwd: project, env });

  for (const filename of ["start-autobericht.cmd", "sync-autobericht.ps1", "program/start-autobericht.ps1",
    "program/project-setup.json", "skill/autobericht-skill.zip", "skill/AutoBericht_portable.md", "guides/README.md"]) {
    assert.ok(fs.existsSync(path.join(installation, filename)), `Missing installed ${filename}`);
  }
  for (const name of ["AutoBericht", "dist", "project-template"]) {
    assert.equal(fs.existsSync(path.join(installation, name)), false, `Obsolete folder ${name}`);
  }
  assert.deepEqual(fs.readdirSync(project), ["project_sidecar.json"]);
  assert.equal(fs.readFileSync(path.join(project, "project_sidecar.json"), "utf8"), '{"private":"unchanged"}');

  const portServer = http.createServer();
  await new Promise((resolve) => portServer.listen(0, "127.0.0.1", resolve));
  const port = portServer.address().port;
  await new Promise((resolve) => portServer.close(resolve));
  launcher = windows
    ? spawn("cmd.exe", ["/d", "/s", "/c", `""${path.join(installation, "start-autobericht.cmd")}" -NoOpen -Port ${port}"`],
      { cwd: project, env, windowsVerbatimArguments: true })
    : spawn(powershell, ["-NoProfile", "-File", path.join(installation, "program/start-autobericht.ps1"), "-NoOpen", "-Port", String(port)],
      { cwd: project, env });
  let startup = "";
  launcher.stdout.on("data", (chunk) => { startup += chunk; });
  launcher.stderr.on("data", (chunk) => { startup += chunk; });
  let launchError;
  launcher.on("error", (error) => { launchError = error; });
  const base = `http://127.0.0.1:${port}/`;
  let ready = false;
  for (let attempt = 0; attempt < 100; attempt += 1) {
    try { ready = (await fetch(base, { signal: AbortSignal.timeout(500) })).ok; } catch {}
    if (ready || launchError || launcher.exitCode !== null) break;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  assert.ok(ready, `Launcher failed: ${launchError || startup}`);

  const check = async (relative, filename) => {
    const response = await fetch(new URL(relative, base));
    assert.equal(response.status, 200, `HTTP ${relative}`);
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), fs.readFileSync(filename), relative);
  };
  for (const page of ["mini/index.html", "mini/photosorter.html", "mini/librarymaker.html"]) {
    await check(page, path.join(ROOT, page));
    const html = fs.readFileSync(path.join(ROOT, page), "utf8");
    for (const match of html.matchAll(/(?:src|href)="([^"#]+\.(?:js|css))"/g)) {
      const url = new URL(match[1], new URL(page, base));
      await check(url.href, path.join(ROOT, decodeURIComponent(url.pathname)));
    }
  }
  await check("project-setup.json", path.join(ROOT, "project-setup.json"));
  for (const entry of JSON.parse(fs.readFileSync(path.join(ROOT, "project-setup.json"))).files) {
    await check(entry.source, path.join(repo, entry.source));
  }
  for (const locale of ["de", "fr", "it"]) {
    await check(`data/seed/knowledge_base_${locale}.json`, path.join(ROOT, `data/seed/knowledge_base_${locale}.json`));
    await check(`data/checklists/checklists_${locale}.json`, path.join(ROOT, `data/checklists/checklists_${locale}.json`));
  }
  await check("data/weights.json", path.join(ROOT, "data/weights.json"));
  for (const relative of ["skill/AutoBericht_portable.md", ".git/config", "sync-autobericht.ps1"]) {
    assert.equal((await fetch(new URL(relative, base))).status, 404, relative);
  }
});
