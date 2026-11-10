# SPOReplaceLink

SPOReplaceLink is a Windows tool (.NET 10) that fixes broken links in Excel files (xlsx, xlsm, xlsb, xls) stored in SharePoint Online after a migration. It replaces old addresses in hyperlinks, external link sources and formulas according to prefix rules, works with SharePoint through Microsoft Graph with app-only access, keeps the load on the tenant within limits, and records every file in a resumable run. It comes as a desktop app and a CLI for scheduling and automation.

🇺🇦 [Українська версія](README.uk.md)

**User guide — [English](docs/USER_GUIDE.en.md) · [Українською](docs/USER_GUIDE.md)**: scenarios from the first setup to the production run, a reference of all fields, statuses, reports, the CLI. The same guide is built into the app: the "Help" section (section tree, search) and the "?" buttons on screens and profile tabs.
Requirements — [docs/FEATURES.md](docs/FEATURES.md), plan — [docs/IMPLEMENTATION_PLAN.md](docs/IMPLEMENTATION_PLAN.md) (design notes, available in Ukrainian only).

The MVP (stages 1–5) is implemented:

- the rule engine and OOXML processing without Excel;
- working with SharePoint through Microsoft Graph: listing (delta), download, writing with `If-Match`, upload session;
- the state database, three levels of retries, deferred files, the circuit breaker;
- load accounting (RU budget, app-level throttling, work windows, block detection) and the app pool;
- the desktop app: run dashboard, files, errors, retries, settings of profiles, rule templates and connections;
- a CLI for scheduling and automation;
- xls and xlsb residuals through Excel in separate processes;
- pilot tools: a pilot report (RU cost per file, time forecast, SharePoint behavior on write) and a sample of files for manual checking.

A pilot on a real tenant is still ahead — the order and criteria are in the [user guide, scenario S6](docs/USER_GUIDE.en.md#s6-pilot-on-a-single-library) ([Ukrainian: сценарій С6](docs/USER_GUIDE.md#с6-пілот-на-одній-бібліотеці)).

## Languages

The app UI, CLI messages, reports and the built-in help are available in English and Ukrainian.

- **App:** switch with the language button (🌐 icon) in the top bar. The choice is saved in the data folder (`language.txt`).
- **CLI:** `--lang en|uk`, or the `SPO_RL_LANG` environment variable; otherwise the system UI language is used.
- **Default:** the system language — Ukrainian if the Windows UI is Ukrainian, otherwise English.

## Build and tests

```bash
dotnet build -c Release
```

```bash
dotnet test
```

- App: `src/SPOReplaceLink.Desktop/bin/Release/net10.0-windows10.0.17763.0/SPOReplaceLink.exe` (requires the WebView2 Runtime — Windows 10/11 have it).
- CLI: `src/SPOReplaceLink.Cli/bin/Release/net10.0/spo-replace-link.exe`

## Release and installer

```bash
powershell -ExecutionPolicy Bypass -File installer\build.ps1
```

The script runs the tests, publishes the app and the CLI for win-x64 together with .NET (.NET is not needed on the target machine) and puts into `out\release`:

- `SPOReplaceLink-<version>-x64.msi` — the installer: `Program Files\SPOReplaceLink`, a Start menu shortcut, upgrade over the old version. Data in `%LOCALAPPDATA%\SPOReplaceLink` stays after an upgrade and an uninstall.
- `SPOReplaceLink-<version>-x64-portable.zip` — the same folder without installation.
- `SHA256SUMS.txt` — checksums.

The version is taken from `Directory.Build.props` (`<Version>`), or it can be set with a parameter: `-Version 1.0.1`. The installer is built with WiX 5 ([installer/](installer/)); WiX is downloaded as a NuGet package, there is no need to install it separately.

## How it works

**SPOReplaceLink.exe** is a single desktop app that controls everything: settings, starting runs, the dashboard, errors and retries. The engine runs inside the app.

- **Data.** The state database (`state.db`, SQLite), the log (`logs/host-*.jsonl`) and temporary files are in the data folder: `--data`, by default `%LOCALAPPDATA%\SPOReplaceLink`. One instance works per data folder; starting it again opens the already running window.
- **Closing the window during a run.** The app asks what to do:
  - minimize to the tray — the run continues, the window opens with a double click on the icon next to the clock;
  - stop the run and exit — the current files are finished, the rest stays in the queue.
- **Resuming.** After the next start, a stopped run is resumed with the "Resume run" button. A run interrupted by a crash or a reboot is resumed by the app itself, from the same place.
- **CLI.** While the app is open, it listens on a local API at `http://localhost:5080`, and if the port is taken (another user of the machine) — on a free port. It writes the address and the access key to `api.json` in the data folder; the CLI of the same user takes them from there and starts scheduled runs through the API. The API rejects requests without the key, so several users of one machine work independently.

```bash
SPOReplaceLink.exe --data D:\SPOReplaceLink
```

`--selftest <file>` — start, check that the window and sections open, write the result to the file and exit (to check an installation without looking at the screen).

## xls and xlsb through Excel

xlsx, xlsm and xlsb links in `.rels` are processed without Excel. These go through Excel:
- xls files;
- xlsb files whose binary records still contain old addresses (for example, a `HYPERLINK` formula).

How it works:
- **Separate processes.** Excel works in separate `SPOReplaceLink.ExcelWorker.exe` processes, only with a local copy of the file. It is hidden, without dialogs, without updating links, with macros disabled.
- **Replacement without `ChangeLink`.** Excel saves a copy as xlsm, the links in it are replaced by the same engine as for xlsx, and Excel saves the result back in the original format. Excel's `ChangeLink` is not used: it immediately re-reads the new source, and if it is not reachable from this machine (SharePoint), the cached values become `#REF!`.
- **Check in Excel.** The written file is opened in Excel once more. If Edit Links, hyperlinks or formulas still contain old addresses, the file gets the `:residual` mark.
- **Process pool.** The number of Excel processes (2 by default; 0 — without Excel) and the timeout are set on the "Performance" tab of the profile.
- **Hangs.** If Excel does not respond longer than the timeout, the process is killed, and the file gets one more attempt in a new process. If that fails too — `failed:excel`. Other files continue to be processed.
- **After a run** no Excel processes remain. Processes left after the app crashed are terminated on the next start. This applies only to processes the app started itself; the user's Excel is not touched.
- **Without Excel** (not installed or 0 processes) files get `needs_excel`, and they can be processed later with a retry run.
- **Links to a local drive.** Excel would save a link to a file on the same drive as the temp folder (the data folder, on C: by default) without the drive letter. Such a file is not written: it gets `failed:validation` with an explanation. To process it, put the data folder on another drive (`--data D:\…`) or use `local --temp D:\…`. This does not apply to links to SharePoint, network paths and other drives.

## Interface

- **Dashboard** — the current or the last run:
  - counters: found, changed, unchanged, replacements by place, errors by class, deferred, with residuals, needing attention;
  - progress and estimated time to completion, speed;
  - load on SharePoint per app and tenant (RU, 429/503, pauses);
  - zoomable charts (mouse selection, last 1 / 6 / 24 h), breakdowns by format and target;
  - buttons: pause, resume, stop, retry, verify.
- **Runs** — the list and a new run from a profile. The profile is checked before the start; apply requires confirmation with an estimate of the number of files.
- **Run card:**
  - "Files" — filters, search, selection, retry or ignore of the selected ones, export of the current view to CSV/XLSX (via "Save as");
  - "Errors and attention" — groups by status and message with a recommendation and group actions;
  - "Run settings" — a snapshot of the profile.
- **Profiles:**
  - tabs "General and rules", "Targets" (a site by URL or search, a library from a list, a folder tree), "Filters" (with a count by the last listing), "Replacement places", "Performance" (presets, work windows), "Retries", "Connections" (a pool with a warning);
  - cloning, JSON import and export.
- **Rule templates** — a table with checking while typing (duplicates, conflicts, cycles, repeated replacements), link check, import from CSV/XLSX or from a ReplaceMagic list, export to CSV.
- **Connections** — tenant, Client ID, a certificate from the Windows store, the license range with limits, a check of access to the profile's targets.

You can look at the app without SharePoint on demo data. The `tools/SPOReplaceLink.DemoSeed` utility runs the real engine on a Graph emulator and opens the app with the same emulator — you can start new runs in it:

```bash
dotnet run --project tools/SPOReplaceLink.DemoSeed -- D:\Temp\spo-demo --open
```

- `--large 200000` adds a synthetic run of 200 thousand files.
- `--port` sets the local API port for the CLI (5097 by default).
- The emulated SharePoint is saved to `fake-sharepoint.json` in the data folder when the app closes. Starting again with the same folder and `--open` opens everything as it was, so a stopped run can be resumed.
- If Excel is on the machine, the demo has a real xls (`Fin/Звіт старий.xls`) that is processed through Excel.

## SharePoint setup (once)

SPOReplaceLink signs in to SharePoint as an Entra app (app-only), without a user: an app in Entra + a certificate. There is no password or secret in the settings — the certificate's private key is only in the Windows store on the VM, only the public part is uploaded to Entra.

### 1. Certificate

There is no need to buy a certificate: a self-signed one is fine for an app's sign-in to Entra. It is created on the machine where SPOReplaceLink will run, **under the account that will run it** (otherwise the app will not see the key). If the company has an internal certificate authority and policy requires it, the certificate can be issued there; the requirements are the same: RSA 2048+, with a private key, in the `My` store.

**Option A — the app is run by a regular user** (`certificateStore: "CurrentUser"`). PowerShell under that account, no admin rights needed:

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=spo-replace-link-1" -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 -KeySpec Signature -KeyExportPolicy NonExportable -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
```

**Option B — the certificate is shared by several accounts of the machine** (`certificateStore: "LocalMachine"`). PowerShell as administrator, `Cert:\LocalMachine\My` instead of `Cert:\CurrentUser\My`. Then give the needed accounts access to the key: `certlm.msc` → Personal → Certificates → the certificate → All Tasks → Manage Private Keys → add the account with the **Read** permission.

Next, in the same PowerShell window — the thumbprint for the connection and the public part for Entra:

```powershell
$cert.Thumbprint
```

```powershell
Export-Certificate -Cert $cert -FilePath .\spo-replace-link-1.cer
```

The `.cer` contains only the public key, it can be sent to the Entra admin by email. `NonExportable` means the private key cannot be exported from the VM — for a new VM a new certificate is created and added to the app. To find the thumbprint later: `Get-ChildItem Cert:\CurrentUser\My` (or `LocalMachine\My`).

The certificate is valid for 2 years. Before it expires, create a new one, add it to the app next to the old one, change the thumbprint in the connection, then delete the old one. The app refuses an expired certificate with the `failed:auth` error.

### 2. App in Entra

An Entra admin is needed (Application Administrator or Global Administrator).

1. [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**: a name, for example `spo-replace-link-1`; **Accounts in this organizational directory only**; no Redirect URI is needed.
2. On the **Overview** page, note the **Application (client) ID** → `clientId` and the **Directory (tenant) ID** → `tenantId` (or the `…onmicrosoft.com` domain).
3. **Certificates & secrets** → **Certificates** → **Upload certificate** → the `.cer` file from step 1. The thumbprint shown in Entra must match `$cert.Thumbprint`.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → **Sites.Selected** → **Add**. The `User.Read` (Delegated) permission added by default can be removed.
5. **Grant admin consent for …** — the status must turn green "Granted".

Sites.Selected by itself gives no access to any site — only to those granted explicitly in the next step.

### 3. Access to sites

On each site collection that holds target libraries, grant the app the `write` role (`read` is not enough to apply changes). This is done by a SharePoint admin or a Global Administrator with the `Sites.FullControl.All` permission — through Graph Explorer, PnP.PowerShell or Microsoft Graph PowerShell; step by step, with checking and revoking access — [user guide, 1.5, step 3](docs/USER_GUIDE.en.md#step-3-site-access) ([Ukrainian](docs/USER_GUIDE.md#крок-3-доступ-до-сайтів)). The shortest is PnP:

```powershell
Grant-PnPEntraIDAppSitePermission -AppId <clientId spo-replace-link-1> -DisplayName spo-replace-link-1 -Site https://contoso.sharepoint.com/sites/Fin -Permissions Write
```

### 4. Connection

Copy [samples/connection.example.json](samples/connection.example.json), fill in `tenantId`, `clientId`, `certificateThumbprint`, `certificateStore` (`CurrentUser` or `LocalMachine`, as in step 1) and `licenses`. For several apps in the pool — one file per app, with different `name`s. Import and check: `test` gets a token and tries to read each target of the profile.

```bash
spo-replace-link connections import samples/connection.example.json
```

```bash
spo-replace-link connections test spo-replace-link-1 --profile finance
```

**If `test` fails:**
- "Certificate … not found in the store" — the app is run under a different account than the one the certificate was created under, or the wrong `certificateStore`;
- "without a private key" — the selected store has only an imported `.cer`; choose the store where the certificate was created, or create a new one (a `.cer` is not imported into the store);
- "No access to the private key" — in option B the app's account was not granted Read on the key;
- `AADSTS700027` / `AADSTS700016` — the certificate was not uploaded to the app, or a wrong `clientId`/`tenantId`;
- there is a token, but `403` on a target — admin consent for Sites.Selected was not given, or `write` was not granted on that site (step 3).

## Profile

JSON with rules, targets, filters, replacement places, performance and retry parameters. Example — [samples/finance.profile.json](samples/finance.profile.json). Profiles, connections and rule templates are stored in the app (edited in its window; from the CLI — `profiles|connections|templates import/get/list/delete`); `get -o file.json` — export.

**Rules:**
- `rules` — `old` → `new` pairs. The order does not matter: the rule with the longest old prefix applies, only on a path segment boundary (`/Document` does not touch `/Documents`), ignoring case and the encoding.
- A prefix can be a full address (`https://…`, `\\server\…`, `C:\…`) or a path from the root (`/sites/…`, then `baseHost` is required).
- `templates` — names of rule templates saved in the app; their rules go before `rules`.
- The rule set is checked at start: duplicates, conflicts, cycles, and also rules whose result another rule would change on a repeated run.

**What and where to process:**
- `targets` — site, library, optional folder. `oldLocation` — the old address of this library (for relative links). If not set, the old address is derived from the rules when it is unambiguous.
- `filters` — `extensions`, `include`/`exclude` masks (`*`, `**`, `?`), `maxFileSizeMb`, `modifiedAfter`/`modifiedBefore`.
- `places` — `hyperlinks`, `linkSources`, `formulas`, `otherExternal`. The other places are not implemented yet.

**Connections and load:**
- `connections` — connection names. More than one work only with `pool.enabled: true`. Microsoft does not recommend several apps for one task, so the load of the whole pool is limited by `pool.tenantCeilingPercent` (50% by default) of the tenant budget.
- `performance` — parallelism (`downloads`, `processors`, `uploads`), timeouts, `budgetPercent` (the share of the app's RU limit), work windows `windows` with their own budget, `onlyInWindows`, `timeZone`.
- `retry` — request and file attempts, pauses between attempts, waves for locked files, the circuit breaker, the time until the "app is blocked" conclusion.

RU limits are determined from the connection's `licenses`: the tenant's license range, according to the Microsoft table. They can be adjusted manually if needed.

## Runs

```bash
spo-replace-link profiles import samples/finance.profile.json
```

```bash
spo-replace-link profiles validate finance
```

**inventory** — listing, statistics and a histogram of prefixes of all links (`prefixes.csv`), to complete the rules:

```bash
spo-replace-link inventory finance --wait
```

**scan** — dry-run: what will be replaced:

```bash
spo-replace-link scan finance --wait
```

**apply** — replacement and writing. Only changed files are written, with `If-Match`; the version number before writing is recorded:

```bash
spo-replace-link apply finance --wait
```

**verify** — re-download the files written by run 12 and search for residuals:

```bash
spo-replace-link verify 12 --wait
```

**retry** — retry the failed files of run 12 without a new listing of the library. The selection can be narrowed with `--status`, `--class`, `--ext`, `--path-mask`, `--file-id`, `--list file.csv`:

```bash
spo-replace-link retry 12 --class transient locked --wait
```

**Reports** — `files`, `links`, `residuals`, `attention` (what needs attention and what to do about it), `prefixes`, `sample` (a sample of changed files for manual checking), `summary`, `pilot` (pilot report: RU per file, 429/503, time forecast, whether SharePoint returns the version number and whether it changes files on save):

```bash
spo-replace-link export 12 -o reports\run-12
```

**Control:**
- `status`, `runs`, `files <run> --status failed`;
- `pause`, `resume`, `stop`, `start` (resume a stopped one);
- `ignore <file> --comment "…"` — the file will not be taken into the next retries;
- `delete-run <run> [--yes]` — delete a run that is not executing, with all its results (nothing changes in SharePoint).

Ctrl+C during `--wait` stops the run gracefully: the current uploads finish, the rest of the files stay in the queue until `start`.

**File statuses:**
- `done`, `unchanged`, `planned`, `verified`, `changes_pending` (verify found that there is still something to change);
- `needs_excel`, `ignored`;
- `skipped:filtered|too_large|encrypted|deleted`;
- `failed:transient|conflict|locked|excel|auth|file|validation|internal`;
- suffixes `:residual` and `:needs_excel`.

Intermediate statuses: `pending`, `in_progress`, `deferred` (locked, waiting for the next wave).

**CLI exit codes:**

| Code | Meaning |
|---|---|
| 0 | all good |
| 1 | there are files that need attention |
| 2 | an error in the profile, the parameters or the app's response (or the app is not running) |
| 3 | the run was stopped or interrupted before completion |

## Local processing without SharePoint

Check the rules, see the result for a link, process a local folder (dry-run by default; `--write --out` writes changed copies, the input files are not changed). `--base-url` — the old address of the folder from which relative links are resolved. `--excel N` — the number of Excel processes for xls and xlsb residuals (by default — from the profile; `0` — without Excel). `--temp` — the temp folder (see "Links to a local drive"). A profile for `local` must not refer to the app's rule templates.

```bash
spo-replace-link rules validate -c samples/finance.profile.json
```

```bash
spo-replace-link rules test -c samples/finance.profile.json "https://contoso.sharepoint.com/Document/Analytics/a.xlsx"
```

```bash
spo-replace-link local D:\Copies\Fin -c samples/finance.profile.json --base-url https://contoso.sharepoint.com/Document/Fin --out D:\Fixed\Fin --write
```

## Check in Excel

`tests/excel` — scripts for a machine with Excel:
- `New-ExcelFixtures.ps1` creates real xlsx/xlsm/xlsb/xls with links to the old location;
- `Test-ExcelOutput.ps1` opens the processed files in Excel and checks Edit Links, hyperlinks, formulas, and that the external link reads data from the new source.
- The test files refer to C:, so process the xls and xlsb among them with a temp folder on another drive: `local … --temp D:\Temp\spo`.

## License

MIT — see [LICENSE](LICENSE).
