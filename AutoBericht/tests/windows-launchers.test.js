const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { ROOT } = require("./helpers");

const repoRoot = path.resolve(ROOT, "..");

const readRepoFile = (...relativeParts) => fs.readFileSync(path.join(repoRoot, ...relativeParts), "utf8");

const normalize = (value) => value.replace(/\r\n/g, "\n");

const expectProcessLocalBypass = (content, ps1RelativePath) => {
  assert.match(
    content,
    new RegExp(`powershell\\.exe\\s+-NoProfile\\s+-ExecutionPolicy\\s+Bypass\\s+-File\\s+\"%PS1%\"\\s+%\\*[\\s\\S]*set\\s+\"EXIT_CODE=%ERRORLEVEL%\"[\\s\\S]*exit\\s+/b\\s+%EXIT_CODE%`, "i"),
  );
  assert.equal(content.includes(`set "PS1=%SCRIPT_DIR%${ps1RelativePath}"`), true);
};

test("Windows cmd launchers use process-local bypass, quote variables, and propagate PowerShell failures", () => {
  const startCmd = normalize(readRepoFile("start-autobericht.cmd"));
  const syncCmd = normalize(readRepoFile("sync-autobericht.cmd"));

  expectProcessLocalBypass(startCmd, "AutoBericht\\start-autobericht.ps1");
  expectProcessLocalBypass(syncCmd, "sync-autobericht.ps1");

  for (const content of [startCmd, syncCmd]) {
    assert.match(content, /setlocal/i);
    assert.match(content, /if not exist "%PS1%"/i);
    assert.doesNotMatch(content, /Set-ExecutionPolicy/i);
  }
});

test("sync script unblocks only the installed app-owned PowerShell scripts after copy", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));

  const copyIndex = script.indexOf("Copy-Item");
  const unblockIndex = script.indexOf("foreach ($relativeScript");
  assert.notEqual(copyIndex, -1, "Copy-Item call missing");
  assert.notEqual(unblockIndex, -1, "Unblock helper usage missing");
  assert.ok(unblockIndex > copyIndex, "scripts must be unblocked after copy");

  const unblockArrayMatch = script.match(/foreach\s*\(\$relativeScript\s+in\s+@\(([\s\S]*?)\)\s*\)\s*\{/);
  assert.ok(unblockArrayMatch, "unblock target array missing");
  const unblockedTargets = Array.from(unblockArrayMatch[1].matchAll(/'([^']+\.ps1)'/g)).map((match) => match[1]);
  assert.deepEqual(unblockedTargets, [
    "sync-autobericht.ps1",
    "AutoBericht\\start-autobericht.ps1",
    "AutoBericht\\tools\\serve-autobericht.ps1",
  ]);
  const unblockCommands = Array.from(script.matchAll(/Unblock-File\b[^\n]*/g)).map((match) => match[0]);
  assert.equal(unblockCommands.length, 1);
  assert.match(unblockCommands[0], /Unblock-File -LiteralPath \$scriptPath -ErrorAction Stop/);
  assert.doesNotMatch(unblockCommands[0], /[*?]/);
  assert.doesNotMatch(unblockCommands[0], /-Recurse\b/i);
  assert.equal(script.includes("Expected installed script missing after sync"), true);
  assert.doesNotMatch(script, /Set-ExecutionPolicy/i);
});

test("sync success guidance points users at the startup cmd launcher", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));
  assert.match(script, /Sync complete\.\s+Start AutoBericht with start-autobericht\.cmd\./);
  assert.doesNotMatch(script, /localhost:5501/i);
});
