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

test("sync script clears and verifies the downloaded ZIP before extraction", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));

  const downloadIndex = script.indexOf("Invoke-WebRequest");
  const zipVerifyIndex = script.indexOf('Clear-AndVerifyMarkOfTheWeb -Path $zipPath');
  const expandIndex = script.indexOf("Expand-Archive");
  const extractPayloadVerifyIndex = script.indexOf("Clear-AndVerifyExtractedPayload -Path $sourceInner");
  const copyIndex = script.indexOf("Copy-Item");
  assert.notEqual(downloadIndex, -1, "archive download missing");
  assert.notEqual(zipVerifyIndex, -1, "ZIP marker verification missing");
  assert.notEqual(expandIndex, -1, "archive expansion missing");
  assert.notEqual(extractPayloadVerifyIndex, -1, "extracted payload verification missing");
  assert.notEqual(copyIndex, -1, "copy step missing");
  assert.ok(downloadIndex < zipVerifyIndex, "ZIP verification must happen after download");
  assert.ok(zipVerifyIndex < expandIndex, "ZIP verification must happen before extraction");
  assert.ok(expandIndex < extractPayloadVerifyIndex, "extracted payload verification must happen after extraction");
  assert.ok(extractPayloadVerifyIndex < copyIndex, "extracted payload verification must happen before copy");
});

test("sync script verifies exact launcher surfaces and fails closed on lingering Zone.Identifier streams", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));

  const copyIndex = script.indexOf("Copy-Item");
  const unblockIndex = script.indexOf("foreach ($relativePath");
  assert.notEqual(copyIndex, -1, "Copy-Item call missing");
  assert.notEqual(unblockIndex, -1, "Unblock helper usage missing");
  assert.ok(unblockIndex > copyIndex, "scripts must be unblocked after copy");

  const unblockArrayMatch = script.match(/foreach\s*\(\$relativePath\s+in\s+@\(([\s\S]*?)\)\s*\)\s*\{/);
  assert.ok(unblockArrayMatch, "unblock target array missing");
  const unblockedTargets = Array.from(unblockArrayMatch[1].matchAll(/'([^']+\.(?:cmd|ps1))'/g)).map((match) => match[1]);
  assert.deepEqual(unblockedTargets, [
    "start-autobericht.cmd",
    "sync-autobericht.cmd",
    "sync-autobericht.ps1",
    "AutoBericht\\start-autobericht.ps1",
    "AutoBericht\\tools\\serve-autobericht.ps1",
  ]);

  assert.match(script, /function Assert-FileExists\(/);
  assert.match(script, /Expected file missing during sync verification: \$Path/);
  assert.match(script, /\$ErrorActionPreference = 'Stop'/);
  assert.match(script, /Get-Item -LiteralPath \$Path -Stream \* -ErrorAction Stop/);
  assert.match(script, /\$streams\.Stream -contains 'Zone\.Identifier'/);
  assert.match(script, /Mark-of-the-Web still present after Unblock-File: \$Path/);
  assert.match(script, /Clear-AndVerifyMarkOfTheWeb -Path \(Join-Path \$ResolvedTarget \$relativePath\)/);
  assert.match(script, /function Clear-AndVerifyExtractedPayload\(/);
  assert.match(script, /Expected extracted payload missing during sync verification: \$Path/);
  assert.match(script, /Get-ChildItem -LiteralPath \$Path -Recurse -File/);
  assert.match(script, /Get-Item -LiteralPath \$_.FullName -Stream \* -ErrorAction Stop/);
  assert.match(script, /Extracted payload still contains Zone\.Identifier streams:/);

  const unblockCommands = Array.from(script.matchAll(/^\s*Unblock-File\b[^\n]*/gm)).map((match) => match[0]);
  assert.equal(unblockCommands.length, 1);
  assert.match(unblockCommands[0], /Unblock-File -LiteralPath \$Path -ErrorAction Stop/);
  assert.doesNotMatch(unblockCommands[0], /[*?]/);
  assert.doesNotMatch(unblockCommands[0], /-Recurse\b/i);
  assert.doesNotMatch(script, /Unblock-File\s+-LiteralPath\s+\$ResolvedTarget/i);
  assert.doesNotMatch(script, /Get-ChildItem\s+-LiteralPath\s+\$ResolvedTarget\s+-Recurse\s+-File/i);
  assert.match(script, /catch\s*\{[\s\S]*Write-Error \$_ -ErrorAction Continue[\s\S]*exit 1[\s\S]*\}/);
  assert.doesNotMatch(script, /Set-ExecutionPolicy/i);
});

test("sync success output reports verified marker-free launch files and startup cmd guidance", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));
  assert.match(script, /Verified launcher files are marker-free\./);
  assert.match(script, /Sync complete\.\s+Start AutoBericht with start-autobericht\.cmd\./);
  assert.doesNotMatch(script, /localhost:5501/i);
});

test("sync launcher list is exact and ordered for all five post-copy launch surfaces", () => {
  const script = normalize(readRepoFile("sync-autobericht.ps1"));
  const unblockArrayMatch = script.match(/foreach\s*\(\$relativePath\s+in\s+@\(([\s\S]*?)\)\s*\)\s*\{/);
  assert.ok(unblockArrayMatch, "launcher verification array missing");
  const verifiedTargets = Array.from(unblockArrayMatch[1].matchAll(/'([^']+\.(?:cmd|ps1))'/g)).map((match) => match[1]);
  assert.deepEqual(verifiedTargets, [
    "start-autobericht.cmd",
    "sync-autobericht.cmd",
    "sync-autobericht.ps1",
    "AutoBericht\\start-autobericht.ps1",
    "AutoBericht\\tools\\serve-autobericht.ps1",
  ]);
});
