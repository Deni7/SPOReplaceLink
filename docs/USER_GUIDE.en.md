# SPOReplaceLink — user guide

MVP version (stages 1–5), 26.10.2026.

SPOReplaceLink fixes broken links in Excel files (xlsx, xlsm, xlsb, xls) stored in SharePoint Online libraries after a migration. Links to old addresses (for example, `https://contoso.sharepoint.com/Document/Fin/…`) are replaced with new ones (`https://contoso.sharepoint.com/sites/finance/Fin/…`) according to rules you define.

The guide has three parts:
- **Part 1** — what you need to know before you start: concepts, installation, launch.
- **Part 2** — step-by-step scenarios, from the initial setup to the production run and handling errors.
- **Part 3** — reference: every field of every screen (what it is, why it exists, what it affects, what it results in), statuses, reports, CLI commands.

---

## Contents

- [Part 1. Before you start](#part-1-before-you-start)
  - [1.1. How it works](#11-how-it-works)
  - [1.2. Key concepts](#12-key-concepts)
  - [1.3. Safety guarantees](#13-safety-guarantees)
  - [1.4. Requirements and installation](#14-requirements-and-installation)
  - [1.5. Setting up SharePoint access (one-time)](#15-setting-up-sharepoint-access-one-time)
  - [1.6. Launching the app](#16-launching-the-app)
  - [1.7. The app window](#17-the-app-window)
- [Part 2. Scenarios](#part-2-scenarios)
  - [S1. Initial setup of SharePoint access](#s1-first-time-sharepoint-access-setup)
  - [S2. Replacement rules (template)](#s2-replacement-rules-template)
  - [S3. A profile for a library](#s3-profile-for-a-library)
  - [S4. Inventory and completing the rules](#s4-inventory-and-completing-the-rules)
  - [S5. Dry run (scan) and reviewing the report](#s5-dry-run-scan-and-report-review)
  - [S6. Pilot on a single library](#s6-pilot-on-a-single-library)
  - [S7. Production run](#s7-production-run)
  - [S8. Errors and retries](#s8-errors-and-retries)
  - [S9. Residual old addresses](#s9-residual-old-addresses)
  - [S10. xls and xlsb files (Excel)](#s10-xls-and-xlsb-files-excel)
  - [S10a. Password-protected files](#s10a-password-protected-files)
  - [S11. Pause, stop, closing the window, reboot](#s11-pause-stop-closing-the-window-reboot)
  - [S12. Restoring a previous file version](#s12-restoring-a-previous-version-of-a-file)
  - [S13. Scheduled runs from the command line](#s13-scheduled-runs-from-the-command-line)
  - [S14. Processing a local folder without SharePoint](#s14-processing-a-local-folder-without-sharepoint)
  - [S15. Moving settings to another machine](#s15-moving-settings-to-another-machine)
- [Part 3. Reference](#part-3-reference)
  - [3.1. Connection](#31-connections)
  - [3.2. Rule template](#32-rule-template)
  - [3.3. Profile](#33-profile)
  - [3.4. New run](#34-new-run)
  - [3.5. Retry run](#35-retry-run)
  - [3.6. Dashboard and run card](#36-dashboard-and-run-card)
  - [3.7. File table, errors, file card](#37-files-table-errors-file-card)
  - [3.8. Statuses](#38-statuses)
  - [3.9. Error classes and what to do](#39-error-classes-and-what-to-do)
  - [3.10. Reports](#310-reports)
  - [3.11. Command line (CLI)](#311-command-line-cli)
  - [3.12. App files and folders](#312-app-files-and-folders)
  - [3.13. Limitations](#313-limitations)
  - [3.14. If something goes wrong](#314-troubleshooting)

---

# Part 1. Before you start

## 1.1. How it works

For each file the app does the following:

1. Gets the list of files in the library via Microsoft Graph.
2. Downloads the file to its machine.
3. Finds the links in it (hyperlinks, links to other workbooks, formulas) and replaces old addresses with new ones according to the rules.
4. Verifies the result: the file opens, the number of replacements matches, no old addresses remain. If verification fails, the file is not written.
5. Writes the file back to SharePoint, only if something in it has changed and nobody has modified the file in the meantime.

xlsx and xlsm are processed without Excel, as ZIP archives of XML. xls and some xlsb files are processed via Excel installed on the same machine, in separate hidden processes.

Excel does not open files from the network, and the app does not work through OneDrive sync or a browser. All work goes through Graph with a local copy.

## 1.2. Key concepts

| Concept | What it is |
|---|---|
| **Connection** | An app registered in Entra (Azure AD) and the certificate it uses to sign in to SharePoint. The app gets access only to the sites explicitly granted to it. |
| **Rule** | An "old prefix → new prefix" pair. For example, `https://contoso.sharepoint.com/Document` → `https://contoso.sharepoint.com/sites/finance`. A link that starts with the old prefix gets the new one; the rest of the address (folders, file name) stays as is. |
| **Rule template** | A named set of rules that can be attached to several profiles. |
| **Target** | Site + library + (optionally) folder where files need to be fixed. |
| **Profile** | Everything a run needs: rules and templates, targets, filters, replacement locations, speed, retries, connection. |
| **Run** | One execution of a profile in one of the modes. It has a number, a state and its own history; it resumes after being stopped. |
| **Mode** | `inventory` — listing, statistics and an address histogram; `scan` — dry-run processing without writing; `apply` — replacement and writing; `verify` — checking that no old addresses remain. inventory and scan both download and parse every file (otherwise links cannot be found), so they cost the same. |
| **Settings snapshot** | A copy of the profile, templates and connections as of the moment the run was created. A run works only with its snapshot: later changes to the profile do not affect it. |
| **Retry run** | A new run over selected files of a previous run (for example, only the failed ones). The library is not re-listed. |
| **Residual** | An old prefix left in a file after replacement: the replacement location is disabled or not supported, or there is no rule for the address. Always included in the report. |
| **RU (resource units)** | The units SharePoint uses to measure load. Each app has an RU limit per minute and per day. The app itself makes sure it does not exceed the configured share of the limit. |

## 1.3. Safety guarantees

- **Nothing is written by default.** Only `apply` mode writes to SharePoint, and it requires confirmation.
- **Only changed files are written.** Files where nothing was changed are not uploaded and do not get a new version.
- **Other people's changes are not overwritten.** If a file was modified after the app downloaded it, the write is rejected and the file is processed again with the new version.
- **Checked-out files** or files opened by others are deferred and retried later.
- **The version number is recorded before writing.** The previous version can be restored from the SharePoint version history (S12).
- **Re-running over already fixed files changes nothing** and writes nothing.
- **Every file gets a final status**, and every unfixed old address is included in the report.
- **Care for SharePoint.** The app stays below the load limit. If SharePoint asks to wait, the pause applies to the whole app. If there are many errors in a row, the run stops itself (circuit breaker).

## 1.4. Requirements and installation

**Machine:**
- Windows 10/11 or Windows Server with Desktop Experience;
- WebView2 Runtime (already present in Windows 10/11; the installer checks for it);
- .NET is not needed separately: it is bundled in the package;
- for xls and some xlsb files — Microsoft Excel installed (365, 2016+), activated for the account the app runs under;
- access to `login.microsoftonline.com` and `graph.microsoft.com`.

A dedicated VM with automatic sign-in of the account is recommended: large runs take hours.

**Installing with the installer** — `SPOReplaceLink-<version>-x64.msi`:

- installs the app to `C:\Program Files\SPOReplaceLink` (administrator rights required) and adds a shortcut to the Start menu;
- the same folder contains the app `SPOReplaceLink.exe`, the CLI `spo-replace-link.exe`, the Excel worker and this guide (`docs\`);
- a newer version installs over the old one, without uninstalling;
- data (connections, profiles, run state, logs) is stored separately, in `%LOCALAPPDATA%\SPOReplaceLink`. Upgrading and uninstalling the app do not touch it.

Silent installation (for example, via Intune or SCCM):

```bash
msiexec /i SPOReplaceLink-1.0.1-x64.msi /qn
```

**Without installation** — `SPOReplaceLink-<version>-x64-portable.zip`: extract to any folder (for example, `C:\Tools\SPOReplaceLink`) and run `SPOReplaceLink.exe`. Administrator rights are not required.

**Building a release from source** (on a developer machine, .NET 10 SDK required):

```bash
powershell -ExecutionPolicy Bypass -File installer\build.ps1
```

The script runs the tests, publishes the app and the CLI together with .NET, and builds the MSI and zip into `out\release`.

**Verifying the installation without looking at the screen:**

```bash
SPOReplaceLink.exe --selftest C:\Temp\selftest.txt
```

The file must contain `RESULT ok`.

**Access to SharePoint** is set up by an administrator once: a certificate, an app in Entra with the `Sites.Selected` permission, and `write` access granted to the required sites. Step by step — [1.5](#15-setting-up-sharepoint-access-one-time). Then — scenario [S1](#s1-first-time-sharepoint-access-setup).

## 1.5. Setting up SharePoint access (one-time)

SPOReplaceLink signs in to SharePoint as an Entra app (app-only), without a user: an app in Entra + a certificate. There is no password or secret in the settings — the certificate's private key resides only in the Windows certificate store on the VM; only the public part is uploaded to Entra.

Order: certificate on the machine → app in Entra → site access → connection in SPOReplaceLink.

### Step 1. Certificate

You do not need to buy a certificate: a self-signed one is fine for app sign-in to Entra. It is created on the machine where SPOReplaceLink will run, **under the account that will run it** (otherwise the app will not see the key), or in the machine store with access for that account (Option B). If your company has an internal certificate authority and policy requires it, you can issue the certificate there; the requirements are the same: RSA 2048+, with a private key, in the `My` store.

**Option A — the app is run by a regular user** (`CurrentUser` store). PowerShell under that account, no administrator rights required:

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=spo-replace-link-1" -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 -KeySpec Signature -KeyExportPolicy NonExportable -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
```

**Option B — the certificate is shared by several accounts on the machine** (`LocalMachine` store). PowerShell as administrator, `Cert:\LocalMachine\My` instead of `Cert:\CurrentUser\My`. Then give the required accounts access to the key: `certlm.msc` → Personal → Certificates → the certificate → All Tasks → Manage Private Keys → add the account with **Read** permission.

Next, in the same PowerShell window — the thumbprint for the connection and the public part for Entra:

```powershell
$cert.Thumbprint
```

```powershell
Export-Certificate -Cert $cert -FilePath .\spo-replace-link-1.cer
```

**Where the key is and where the `.cer` is.** The private key is created together with the certificate and stays only in the store where `New-SelfSignedCertificate` was run; SPOReplaceLink takes the certificate from exactly there. The `.cer` contains only the public key and is needed only for Entra (step 2) — it can be sent to the admin by email. **Do not import the `.cer` into the Windows store:** this creates a copy of the certificate without the private key, and the app will show it as "(no private key)".

`NonExportable` means the private key cannot be exported (neither to `.pfx` nor otherwise): the certificate cannot be moved to another VM or another account. For another account on this machine — Option B (`LocalMachine` and **Read** permission on the key); for a new VM — a new certificate added to the Entra app alongside the old one. To find the thumbprint and check the key later: `Get-ChildItem Cert:\CurrentUser\My | Format-List Subject, Thumbprint, HasPrivateKey` (or `LocalMachine\My`).

In the connection's certificate list ([3.1](#31-connections)) the markers are: "(no private key)" — this store holds only the public part; "(no access to the key)" — the key exists, but the account the app runs under has not been granted **Read** permission (Option B).

The certificate is valid for 2 years. Before it expires, create a new one, add it to the app alongside the old one, change the certificate in the connection, then delete the old one. The app refuses an expired certificate with the error `failed:auth`.

### Step 2. App in Entra

An Entra admin is required (Application Administrator or Global Administrator).

1. [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**: a name, for example `spo-replace-link-1`; **Accounts in this organizational directory only**; no Redirect URI needed.
2. On the **Overview** page, note the **Application (client) ID** (the connection's **Client ID** field) and the **Directory (tenant) ID** (the **Tenant** field; or the `…onmicrosoft.com` domain).
3. **Certificates & secrets** → **Certificates** → **Upload certificate** → the `.cer` file from step 1. The thumbprint shown in Entra must match `$cert.Thumbprint`.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → **Sites.Selected** → **Add**. The `User.Read` (Delegated) permission added by default can be removed.
5. **Grant admin consent for …** — the status must turn into a green "Granted".

Sites.Selected by itself does not grant access to any site — only to those explicitly granted in the next step.

### Step 3. Site access

After step 2 the app has the `Sites.Selected` permission but **does not yet have access to any site**: access is granted separately for each site. Until this is done, the connection test obtains a token but gets `403` for the target.

**What to grant.** The `write` role on each **site collection** containing target libraries (an address like `https://contoso.sharepoint.com/sites/Fin`). This is enough for all modes: reading files and versions, writing fixed files. The role applies to all libraries and subsites of the collection; it does not change library permissions (inheritance is not broken).

| Role | Suitable? |
|---|---|
| `read` | Inventory and scan only; applying changes will end with `403`. |
| `write` | **Required.** |
| `owner`, `fullcontrol` | Excessive: they grant control over the site and its permissions. |

**Who grants it.** A SharePoint admin or Global Administrator — an account that can consent to the `Sites.FullControl.All` permission (this is what is needed to grant site access). This permission is **not added** to the `spo-replace-link-1` app itself, nor is `Sites.ReadWrite.All`: with them, the per-site restriction no longer applies.

You will need the app's **Client ID** from step 2 and the site addresses. Below are three equivalent methods — choose whichever is most convenient for the admin.

#### Method 1. Graph Explorer (in the browser, nothing to install)

1. Open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) and sign in with the admin account.
2. **Modify permissions** (tab above the request) → find `Sites.FullControl.All` → **Consent**. If the button is unavailable, Global Administrator consent is required.
3. Get the site ID: method **GET**, the address below → **Run query** → copy the `id` field from the response (like `contoso.sharepoint.com,5a58…,9f2e…`).

```http
GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Fin
```

4. Grant access: method **POST**, the address with that `id`, in **Request body** — the JSON below (substitute the Client ID) → **Run query**. The response is `201 Created`.

```http
POST https://graph.microsoft.com/v1.0/sites/{site id}/permissions
Content-Type: application/json

{
  "roles": ["write"],
  "grantedToIdentities": [{ "application": { "id": "<app Client ID>", "displayName": "spo-replace-link-1" } }]
}
```

5. Verify: **GET** `https://graph.microsoft.com/v1.0/sites/{site id}/permissions` — the list contains an entry with `spo-replace-link-1` and the `write` role.

Repeat steps 3–5 for each site.

#### Method 2. PnP.PowerShell (convenient for many sites)

Requires PowerShell 7.4+ and the module `Install-Module PnP.PowerShell -Scope CurrentUser`. PnP needs its own app in Entra to sign in — it is registered once (a Global Administrator consents to the permissions), and its Client ID is used from then on:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant contoso.onmicrosoft.com
```

Grant access (sign in to any site of the tenant; the site for the permission is set by `-Site`):

```powershell
Connect-PnPOnline -Url https://contoso.sharepoint.com -Interactive -ClientId <Client ID of the PnP app>
```

```powershell
Grant-PnPEntraIDAppSitePermission -AppId <Client ID spo-replace-link-1> -DisplayName spo-replace-link-1 -Site https://contoso.sharepoint.com/sites/Fin -Permissions Write
```

Several sites at once:

```powershell
"https://contoso.sharepoint.com/sites/Fin", "https://contoso.sharepoint.com/sites/finance" | ForEach-Object { Grant-PnPEntraIDAppSitePermission -AppId <Client ID spo-replace-link-1> -DisplayName spo-replace-link-1 -Site $_ -Permissions Write }
```

Verify: `Get-PnPEntraIDAppSitePermission -Site https://contoso.sharepoint.com/sites/Fin`. Older PnP versions have the same commands named `…PnPAzureADAppSitePermission`.

#### Method 3. Microsoft Graph PowerShell

```powershell
Connect-MgGraph -Scopes Sites.FullControl.All
```

```powershell
$site = Get-MgSite -SiteId "contoso.sharepoint.com:/sites/Fin"
New-MgSitePermission -SiteId $site.Id -BodyParameter @{ roles = @("write"); grantedToIdentities = @(@{ application = @{ id = "<Client ID spo-replace-link-1>"; displayName = "spo-replace-link-1" } }) }
```

Verify: `Get-MgSitePermission -SiteId $site.Id`.

#### After granting

- **Verify in SPOReplaceLink:** **Connections → Test connection** with a profile that includes this site. Every target must be green. If access was just granted and the test still returns `403`, retry in a few minutes.
- **A new site in a profile** — grant access to it the same way first, otherwise the run for that target will end with an access error.
- **Site search** in the profile editor may not find sites the app has no access to — enter the full site address.
- **Revoke access** when the work is done: in Graph — `DELETE https://graph.microsoft.com/v1.0/sites/{site id}/permissions/{permission id}` (`permission id` — from the GET `…/permissions` response); in PnP — `Revoke-PnPEntraIDAppSitePermission -PermissionId <id> -Site …`; in Graph PowerShell — `Remove-MgSitePermission`. To disable the app for all sites entirely, revoke consent for `Sites.Selected` in Entra (step 2).

### Step 4. Connection in SPOReplaceLink

In the app window — scenario [S1](#s1-first-time-sharepoint-access-setup): **Connections → Create**, tenant, Client ID, certificate from the store, **Test connection**. For several apps in a pool — one connection per app, with different names.

From the command line — import JSON (`tenantId`, `clientId`, `certificateThumbprint`, `certificateStore` — `CurrentUser` or `LocalMachine`, as in step 1, and `licenses`) and test: `test` obtains a token and tries to read each target of the profile.

```bash
spo-replace-link connections import connection.json
```

```bash
spo-replace-link connections test spo-replace-link-1 --profile finance
```

**If the test fails:**
- "Certificate … not found in the store" — the app is running under a different account than the one the certificate was created under, or the wrong store is selected;
- "no private key" — the selected store contains only an imported `.cer`; select the store where the certificate was created, or create a new one (a `.cer` is not imported into the store);
- "No access to the private key" — in Option B the app's account was not granted Read on the key;
- `AADSTS700027` / `AADSTS700016` — the certificate was not uploaded to the app, or the Client ID / tenant is wrong;
- a token is obtained, but `403` for the target — admin consent for Sites.Selected was not given (step 2), or `write` was not granted on this site (step 3).

## 1.6. Launching the app

```bash
SPOReplaceLink.exe --data D:\SPOReplaceLink
```

| Parameter | What it is | Purpose and effect | Default |
|---|---|---|---|
| `--data <folder>` | Data folder: state database, logs, temporary files. | All settings and runs are stored here. A different folder is a separate, independent data set. **Recommended on drive D:** (not on the drive holding local files referenced by Excel workbooks — see [S10](#s10-xls-and-xlsb-files-excel)). | `%LOCALAPPDATA%\SPOReplaceLink` |
| `--urls <address>` | Local API address for the CLI. | Usually not needed. If the port is busy (for example, another user of this machine has already started the app), the API takes a free port and writes the address to `api.json` in the data folder — the same user's CLI finds it automatically. An explicit `--urls` is needed only for a fixed address; if it is busy, the app warns you and shows the actual one. | `http://localhost:5080` |
| `--selftest <file>` | Self-test: start, open the sections, write the result, exit. | Verifying an installation or upgrade. | — |

One instance runs per data folder: launching again opens the window that is already running.

**Several users on one machine.** Each user runs the app under their own account: they have their own data folder (`%LOCALAPPDATA%\SPOReplaceLink`), their own settings, runs and certificate. The first instance listens on `localhost:5080`, subsequent ones on free ports; each user's CLI finds its own instance via `api.json`. The API accepts only requests with its own instance's key (also stored in `api.json`), so another user of the machine cannot control someone else's runs. Do not put the data folder (`--data`) in a folder shared by several accounts: others could read the key from it.

**Autostart on a VM.** To have a run continue after a reboot, enable automatic sign-in for the account and add a Task Scheduler task "At log on" with the command `SPOReplaceLink.exe --data D:\SPOReplaceLink`. A run interrupted by a crash or forced process termination is resumed by the app automatically. A run stopped by a normal exit from the app is in the "stopped" state and is resumed with the **Resume run** button (or `spo-replace-link start <run>`). Before a planned VM reboot, stop the run, or check after the reboot that it has resumed.

## 1.7. The app window

The menu is on the left:
- **Dashboard** — the current run (or the last one, if nothing is running now);
- **Runs** — the list of runs and the "New run" button;
- **Settings:** **Profiles**, **Rule templates**, **Connections**.
- **Help** — this guide inside the app: sections as a tree, search, "previous / next" buttons.

Next to screen headings and on each profile tab there is a **?** button ("What these fields mean", "Help"). It opens, in a window, the guide section about exactly these fields; **Open in Help** shows the same section on the Help page. The **?** button in the top right corner opens Help from the beginning.

Each settings list (profiles, templates, connections) has:
- **Create**, **Import JSON**;
- for each row — export JSON, rename (references in profiles are updated), clone and delete (the app warns which profiles use the template or connection).

Closing the window during a run — see [S11](#s11-pause-stop-closing-the-window-reboot).

---

# Part 2. Scenarios

## S1. First-time SharePoint access setup

**Goal:** the app obtains a token and can see the required libraries.

**What you need from the administrator:** Tenant ID, the app's Client ID, a certificate in the Windows certificate store on this machine (under your account), and `write` access to the sites. How to set this up — [1.5](#15-setting-up-sharepoint-access-one-time).

1. **Connections → Create.**
2. Fill in:
   - **Name** — for example, `spo-replace-link-1`;
   - **Tenant** — the `…onmicrosoft.com` domain or a GUID;
   - **Client ID**.
3. **Authentication** — "Certificate". In the **Certificate from Windows store** list, select your certificate: the thumbprint and store are filled in automatically.
4. **Licenses in tenant** — select a range. It determines the load limits (table on the right).
5. **Save.**
6. **Test connection:** first without a profile — you should see "token: Token obtained". Once you have a profile ([S3](#s3-profile-for-a-library)), select it and test again — each target should show "Access granted".

**Result:** the connection is saved and tested. If the test fails — [3.14](#314-troubleshooting).

## S2. Replacement rules (template)

**Goal:** a set of "old location → new location" rules that you can check before any run.

1. **Rule templates → Create**, a name (for example, `finance`), a description.
2. Fill in the rules in one of two ways:
   - **manually** — **Add rule**: "Old prefix" and "New prefix";
   - **by import** — **Import rules**: a CSV/XLSX mapping table with two columns (old, new) or text from a ReplaceMagic list.
3. Watch the messages under the rows: duplicates, conflicts, cycles (A→B and B→A), and rules whose result another rule would change on a repeated run. Errors must be fixed; warnings should be reviewed.
4. **Test link:** paste a real old link from a file. The app shows:
   - the normalized address;
   - which rules matched (✔ — the one that will apply);
   - the result;
   - whether applying the rules again would change it.

   For relative links (`../Analytics/a.xlsx`), fill in "Old document folder".
5. **Save.** **CSV** — export the rules for sign-off by data owners.

**How rules work:**
- **order does not matter** — the rule with the longest old prefix wins;
- **only at a folder boundary** — `/Document` will not match `/Documents` or `/MyDocument`;
- **case and encoding do not matter** — `%20` and a space, encoded or unencoded Cyrillic, `\` and `/`;
- **only the prefix is replaced** — the rest of the address (folders, file, `#Sheet!A1`) is copied unchanged, in the same encoding style as the original;
- **a prefix can be** a full URL (`https://…`), a network path (`\\server\share`), a local path (`C:\Users\…`), or a site-root-relative path (`/Document/…`) — the latter requires `baseHost` in the profile.

**Result:** a template you can attach to a profile.

## S3. Profile for a library

**Goal:** everything needed for a run against a specific library.

1. **Profiles → Create**, a name (for example, `finance-pilot`).
2. **General and rules:**
   - `baseHost` — `https://contoso.sharepoint.com`;
   - select the template(s);
   - if needed, add the profile's own rules. They are validated together with the template rules.
3. **Connection:** select the connection from S1.
4. **Targets → Add target:**
   - paste the site URL → **Open**;
   - select the **Library**;
   - if needed, navigate into a folder → **Use this folder**;
   - fill in the library's **Old address** (where it was before the migration) if the files contain relative links.
5. **Filters:** keep all extensions. Exclude archives or system folders with masks (for example, `**/Archive/**`).
6. **Replacement locations:** leave everything enabled unless you have a reason not to.
7. **Performance:** for a pilot — the **Careful** preset; for a production run — **Off-hours**.
8. **Save.** The app validates the profile and marks errors on the tabs and next to the fields. A profile with errors is saved but cannot be started.
9. Go back to **Connections → Test** with this profile.

**Result:** a profile ready to run.

## S4. Inventory and completing the rules

**Goal:** see which addresses actually occur in the files and whether all of them are covered by rules. Nothing is written, but every file is downloaded and parsed — in time and load it is comparable to a scan. The inventory result also shows what will be replaced (`planned` statuses), so a separate scan is needed only after the rules change.

1. **Runs → New run**, profile, mode **inventory**, **Start**.
2. When it finishes: **Reports → Prefixes — XLSX**. The app asks for a depth — how many folder levels below the site, server or drive to keep in the prefix (1–10, default 2; in the CLI — `export … --prefix-depth N`). Each row is the beginning of an address up to that depth and how many links and files contain it. If different folders of the same library move to different places, increase the depth. Columns:
   - `replaced` — covered by rules;
   - `no_rule` — no rule;
   - `unresolved` — a relative link that could not be resolved.
3. Go through the rows with `no_rule` > 0:
   - **an old address that must be fixed** — add a rule to the template (S2);
   - **an address that must not be touched** (external site, new address) — do nothing;
   - **local paths `C:\Users\…`, old servers** — decide with the data owners whether and where to redirect them.
4. Rows with `unresolved` > 0 mean that the library's **Old address** (target) is not set or cannot be derived from the rules. Fill it in.

**Result:** the rules cover everything that needs fixing. The number of files in the profile is now shown on the "Targets" and "Filters" tabs (from the latest run of this profile that enumerated the library).

## S5. Dry run (scan) and report review

**Goal:** see exactly what will be replaced before anything is written.

1. **New run**, mode **scan**.
2. On the dashboard:
   - **Changed** (here — "will be changed");
   - **Link replacements** broken down by location;
   - **With residuals**;
   - **Errors**.
3. Review the reports:
   - **Links** — every replacement: file, file part, type, old and new value, rule number;
   - **Review sample** — 30 random files that will be changed, with a link to each file and sample replacements. Go through it manually: does each new address point where it should (the `checked` and `comment` columns are for your notes);
   - **Residuals** — what will remain unfixed ([S9](#s9-residual-old-addresses));
   - **Needs attention** — files with errors and what to do about them.
4. If something is wrong, fix the rules and run the scan again. Nothing is written, so you can repeat it as often as you like.

**Result:** you know what apply will do and have signed off on it.

## S6. Pilot on a single library

**Goal:** test the tool on real data and real SharePoint with one library before running it on all of them. A pilot is a mandatory step before a production run.

**Preparation:**
1. Choose one library with a variety of files (xlsx, xlsm, xls, xlsb, files with links between workbooks). Preferably one whose owner agrees to check the files.
2. Make sure **versioning** is enabled in the library (Library settings → Versioning settings) and the version limit is not too low: every written file gets one new version.
3. Profile: one target, the **Careful** preset. Ask the owners of key files not to edit them during the run, or run it outside working hours.

**Pilot steps:**
1. **inventory** → complete the rules ([S4](#s4-inventory-and-completing-the-rules)).
2. **scan** (if the rules changed; otherwise use the reports of the same inventory) → manual review of the sample ([S5](#s5-dry-run-scan-and-report-review)).
3. **apply.** The dialog shows the number of files; tick the confirmation checkbox. During the run, watch the **SharePoint load** panel: 429/503 responses and pauses should be rare or absent.
4. When it finishes — **Verify** on the run card. Verify re-downloads all written files and searches them for old addresses ([how it works](#the-verify-button)). Expected result: 0 files "Still need changes", 0 residuals (except agreed ones).
5. **Reports → Pilot report — TXT** for the apply run. It contains:
   - how many files were written, by format;
   - **RU per file** compared with the plan (~2 for an unchanged file, ~4 for a changed one) and the number of 429/503 responses;
   - throughput over active processing time (excluding waiting for windows, pauses and breaks between starts), seconds per file, the throughput limit set by the RU budget (taking windows and the daily limit into account), and the **time forecast for 10,000 files** and for the profile. For other libraries: their file count × seconds per file;
   - **whether SharePoint returns the version number before writing.** If 0 of N, version-based rollback and the "Save pre-change version" button will not work — pass the report to the developer;
   - **whether SharePoint modifies the file on save** (adds library column properties) — based on the size in the write response. This does not affect correctness, but you should be aware of it;
   - how many attempts ended in a conflict (412) or a lock (423, check-out), including those after which the file was eventually written;
   - errors, grouped, with recommendations;
   - Excel: operations, hangs, crashes.
6. **User check.** Give the owners a list of key files (the "Review sample" report or your own). Ask each of them to open the file in the browser and in Excel and check:
   - there is no file recovery warning;
   - hyperlinks open;
   - **Data → Edit Links** shows the new addresses, and **Update Values** pulls data from the new source;
   - formulas with `HYPERLINK` point to the new addresses;
   - for macro-enabled workbooks — the macros work.

**Pilot success criteria:**
- verify: 0 "Still need changes", only agreed residuals;
- users have confirmed the key files;
- no 429 responses, or only occasional ones, with no throttling block;
- no `failed:validation` or `failed:internal` (or they have been analyzed with the developer);
- RU per file and the time forecast are acceptable for the full volume.

**If something goes wrong:** collect the reports (`export`, [3.11](#311-command-line-cli)) and the log (`logs\host-*.jsonl` in the data folder) and pass them to the developer. An individual file can be restored to its previous version ([S12](#s12-restoring-a-previous-version-of-a-file)).

## S7. Production run

**Goal:** fix all libraries with minimal impact on users.

1. One profile per group of libraries. A single profile with several targets is possible, but batches (one or several libraries at a time) are easier to control.
2. Performance: the **Off-hours** preset. It works only in the windows Mon–Fri 20:00–07:00 and all weekend, with an 80% budget. Outside the windows the run pauses itself and resumes when the next window opens.
3. **apply** → **verify** → review errors → retries ([S8](#s8-errors-and-retries)).
4. After the last batch — a final **verify** for the profile (New run → verify), export the reports, and archive a copy of the data folder: it holds the history of all changes and the pre-write version numbers.

Enable an **app pool** (several connections in one run) only if one app is genuinely not fast enough. Microsoft does not recommend using several apps for one task: they share a single tenant budget with users and with other apps (backup, migration). Therefore, the combined load is capped (by default at 50% of the tenant budget).

## S8. Errors and retries

**Goal:** every file has a final status, and failed files are retried selectively, without a new pass over the library.

**What the app does on its own:**
- **network and SharePoint failures** (5xx, 429, dropped connection) — up to 6 attempts per request, then the file is processed again from scratch: up to 3 file attempts in total, with pauses of 1 and 5 min (the pause list is 1, 5, 15: the third value applies if you increase the number of attempts);
- **file changed during processing** (412) — processed again with the fresh version, within the same 3 file attempts; then `failed:conflict`;
- **file open or checked out** (423) — deferred; retried in waves every 10 min for 2 hours;
- **Excel hung or crashed** — one more attempt in a new Excel process.

**What you do — after the run finishes:**
1. **Run card → Errors & attention.** Errors are grouped by status and message, and each group shows what to do.
2. For a group:
   - ☰ — show the group's files;
   - ↻ — retry the group (a new run for these files);
   - ⊘ — ignore with a comment.
3. Or **Retry…** on the run card — a retry by filter (statuses, error classes, extensions, path mask). Or on the **Files** tab — select rows → **Retry selected**.
4. Each retry is a new run linked to its parent. Every file in it is downloaded again.

**Typical actions by class** (details — [3.9](#39-error-classes-and-what-to-do)):
- **transient, conflict** — retry;
- **locked** — ask the owner to close the file or check it in, retry later;
- **access** — check the app's permissions on the site;
- **file** — review manually (corrupted, password-protected, not Excel) or ignore;
- **self-check, internal** — pass to the developer.

**Circuit breaker.** The run stops itself (status "aborted" with the reason on the dashboard) if:
- 10 access errors occur in a row;
- among the last 50 processing results, more than half are transient, access or internal errors. Intermediate failed attempts are counted too (a file that had a transient error three times counts as three), so during mass failures the breaker trips sooner. Locks, conflicts, Excel, file and self-check errors are not counted by the breaker;
- SharePoint has responded only with 503 for 15 min (the app may be blocked).

Eliminate the cause and click **Resume run**. If it is a block, do not resume right away: repeated attempts only extend the block.

## S9. Residual old addresses

The `…:residual` status means the replacements were made, but old prefixes remain in the file. Causes:
- there is no rule for this address;
- the replacement location is disabled in the profile;
- the location is not supported (pivot caches, data connections, Power Query, cell text, VBA).

1. **Reports → Residuals:** file, file part (`xl/pivotCache/…`, `xl/connections.xml`, `xl/vbaProject.bin`, etc.), the match found, and context.
2. If a rule is missing — add it to the template.
3. **Retry…** with the statuses `done:residual`, `unchanged:residual` or `planned:residual` and the **Use current profile version** checkbox — the retry will use the new rules. CLI: `spo-replace-link retry <run> --status done:residual unchanged:residual --current-profile`.
4. Residuals in unsupported locations remain in the report: they are fixed manually or accepted as permissible.

## S10. xls and xlsb files (Excel)

- **xlsx, xlsm** and **xlsb** links in external links are processed without Excel.
- **xls**, as well as **xlsb** files whose binary records still contain old addresses (for example, a `HYPERLINK` formula), are processed through Excel:
  1. Excel saves a copy as xlsm;
  2. the links in the copy are replaced by the same engine as for xlsx;
  3. Excel saves the result in the original format;
  4. Excel opens the result again and checks that no old addresses remain.
- Excel runs hidden, without dialogs, without updating links, with macros disabled, and only on a local copy. Cached link values are preserved.
- **Number of Excel processes and timeout** — profile → Performance. If Excel does not respond for longer than the timeout, the process is force-terminated and the file gets one more attempt. If that also fails — `failed:excel`.
- **Without Excel** (not installed or 0 processes), such files get `needs_excel`. Process them later with a retry run on the `needs_excel` status on a machine with Excel.
- **Password-protected files** (password to open or to modify) — `skipped:encrypted`; Excel is not started for them. The app detects a password in xls files itself, before Excel, from the file contents. Details and retry — [S10a](#s10a-password-protected-files).
- **Excel dialogs.** If Excel still asks something (some versions show a password prompt even when none is needed), the handler immediately closes the window as "Cancel". If the window does not close within 5 s, that Excel is force-terminated. The file gets `skipped:encrypted` (a password was requested) or `failed:file` (another question), and the next file gets a new Excel process. The dashboard shows a "closed dialogs" counter.
- After the run, no `EXCEL.EXE` processes started by the app remain. The app does not touch your own Excel.
- **Explorer recent files.** Some Excel versions add every opened workbook to Windows "Recent items", and temporary copies `book.xls`, `book-out.xls` appear in Explorer ("Quick access → Recent files"). The app deletes shortcuts that point to its temp folder after each file, at the end of a run and at the start of the next one. Your own recent files and Windows settings are not changed. The app does not clear Excel's jump list on the taskbar.
- **Links to a local drive** in xls/xlsb. If a file links to another file on the same drive as the app's data folder (C: by default), Excel would save such a link without the drive letter. Therefore such a file is not written and gets `failed:validation` with an explanation. **Solution:** put the data folder on another drive (`--data D:\SPOReplaceLink`) and retry.

## S10a. Password-protected files

The app does not know passwords and does not try to guess them. A password-protected file is skipped, and the run continues.

1. **What happens.** The file gets `skipped:encrypted` with an explanation: password to open, password to modify, or "Excel asked …". An entry is written to the run log (the **Errors & attention** tab, file page → attempts) and to the app log. The file itself in SharePoint is not changed.
2. **Where to find them.** On the **Errors & attention** tab they form a separate group. On the **Files** tab — filter by the `skipped:encrypted` status. A list for the owners can be exported as a report.
3. **What to do.** The file owner removes the password (File → Info → Protect Workbook → Encrypt with Password, empty field; for a password to modify — Save As → Tools → General Options) or fixes the links themselves.
4. **Retry pass.** **Retry…** with the `skipped:encrypted` status (or the retry button next to the group on the **Errors & attention** tab). CLI: `spo-replace-link retry <run> --status skipped:encrypted`. Files that still have a password will get `skipped:encrypted` again and wait for the next retry.
5. A file that keeps its password can be **ignored** with a comment so it is not picked up by subsequent retries.

## S11. Pause, stop, closing the window, reboot

| Action | What happens | How to continue |
|---|---|---|
| **Pause** | No new files are picked up; current ones are finished. The pause lasts while the app is open: after a restart, the run resumes on its own. | **Resume** |
| **Stop** | Files already being written to SharePoint are written to completion; the remaining current files are interrupted and returned to the queue (nothing written), together with the others. Status "stopped". | **Resume run** on the run card |
| **Closing the window during a run** | The app asks: minimize to the tray (the run continues) or stop and exit. | Double-click the icon next to the clock; after exiting — start the app and **Resume run** |
| **Crash or forced process termination** | The run is interrupted midway. | After starting, the app resumes on its own from the same point. Files that were being processed are processed again; what was already written is not written again. |
| **Normal exit, including during a Windows reboot** | Same as "Stop" (if the app managed to shut down cleanly). | **Resume run** after starting |
| **Outside working windows** (the "Off-hours" preset) | The run waits for the next window. | Automatically |
| **Daily RU budget exhausted** | Paused until the budget recovers; files do not get errors. | Automatically |
| **Delete…** (card of a finished or stopped run) | After confirmation, the run disappears with all its results: file inventory, statuses, attempts, links, metrics. It is no longer shown in the list, in reports, or in the volume estimate for a new run ("From the inventory of run N…"). Nothing changes in SharePoint. Retries created from it remain, only without the link to the parent. A running or paused run cannot be deleted — **Stop** it first. If the run wrote files (apply), the change log and pre-write versions for rollback disappear with it ([S12](#s12-restoring-a-previous-version-of-a-file)): export the reports before deleting. CLI: `delete-run <run> [--yes]`. | — |

Tray icon menu: **Open**, **Data folder** (open in Explorer), **Exit**.

## S12. Restoring a previous version of a file

1. Find the file: run card → **Files** → search by path → click the row.
2. In the file card:
   - **Open in SharePoint** — the file in the browser;
   - **Save pre-change version (N)** — download a copy of the version that existed before the write (for comparison).
3. To restore, in SharePoint: the file's **Version history** → the version with the "Version before write" number from the card → **Restore**.

Bulk rollback (`rollback`) will be available in future versions. Pre-write version numbers are in the **Files** report (`version_before`).

## S13. Scheduled runs from the command line

The app must be open (it can be minimized to the tray): the CLI works through it.

```bash
spo-replace-link apply finance --wait --yes
```

```bash
spo-replace-link export 12 -o D:\Reports\run-12
```

For Task Scheduler, use the exit codes:
- `0` — all good;
- `1` — there are files that need attention;
- `2` — configuration error or the app is not running;
- `3` — the run was stopped.

Help — [3.11](#311-command-line-cli).

## S14. Processing a local folder without SharePoint

For testing rules on copies of files or for manual cases. The profile is a JSON file without templates (rules in `rules`).

```bash
spo-replace-link local D:\Copies\Fin -c finance.profile.json --base-url https://contoso.sharepoint.com/Document/Fin --out D:\Fixed\Fin --write
```

- **Without `--write`** — reports only (dry-run).
- **With `--write --out`** — modified copies are written to another folder; the source files are not changed.
- **`--base-url`** — the old address of the folder against which relative links are resolved.
- **`--excel N`** — number of Excel processes.
- **`--temp`** — temp folder (on another drive if files link to local paths).

## S15. Moving settings to another machine

1. On the old machine: export the JSON of each connection, template and profile (the ⤓ button in the list or `spo-replace-link profiles get finance -o finance.json`).
2. On the new machine: create your own certificate (the private key is not transferred) and add it to the app in Entra.
3. Import the JSON in the same order: connections → templates → profiles. In the connection, use the new thumbprint.
4. Test the connection with the profile.

Run history is stored in `state.db` in the data folder. To archive it, copy the entire data folder while the app is closed.

---

# Part 3. Reference

For each field, the name in the profile or connection JSON (used for import, export and the CLI) is given in parentheses.

## 3.1. Connections

**Connections → [name]**

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Name (`name`) | Connection name in the app. | Profiles refer to the connection by it. Renaming updates the references in profiles. | — |
| Tenant (`tenantId`) | Directory (tenant) ID from Entra, or the `…onmicrosoft.com` domain. | Where the app signs in. A wrong value causes an `AADSTS…` sign-in error. | — |
| Client ID (`clientId`) | Application (client) ID of the app in Entra. | Which app to sign in as. Must be a GUID. | — |
| Authentication (`authMode`) | `certificate` — the app signs in with a certificate, no user involved; `deviceCode` — a user signs in with a code. | Certificate is the primary method: works unattended and on a schedule. Device code is a fallback while the certificate is not set up yet; permissions are then limited to the user's own. | `certificate` |
| Certificate from the Windows store | List of certificates with a private key in `CurrentUser\My` and `LocalMachine\My`, with their validity period. | Selecting one fills in the thumbprint and store. Expired certificates are marked. | — |
| Thumbprint (`certificateThumbprint`) | 40 hexadecimal characters. | Which certificate to use. The private key stays in the Windows store; the settings contain no secrets. | — |
| Store (`certificateStore`) | `CurrentUser` or `LocalMachine`. | Where to look for the certificate. `CurrentUser` — the certificate of the account the app runs under. | `CurrentUser` |
| Licenses in tenant (`licenses`) | Range of the number of M365 licenses in the tenant. | Microsoft uses it to set SharePoint limits, and the app uses it to set its RU budget. The table under the field shows the limits. A range that is too high leads to 429s; too low — to slower processing. | up to 1,000 |
| RU/min (`appPerMinuteOverride`) | Manual value of the app's per-minute limit. | If Microsoft has stated a different limit or the pilot measured a different one. Empty — from the table. | empty |
| RU/day (`appPerDayOverride`) | Manual value of the daily limit. | Same, for the daily limit. | empty |
| Tenant, RU per 5 min (`tenantPer5MinOverride`) | Manual value of the whole-tenant limit. | For the app pool ceiling. | empty |
| Test connection | Button; optionally a profile. | Checks the settings and the certificate, acquires a token and, if a profile is selected, checks access to each of its targets. Result — a "token" line and a line per target; if the settings or the certificate are wrong — a single line with the reason. | — |

## 3.2. Rule template

**Rule templates → [name]**

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Name (`name`) | Template name. | Profiles attach the template by it. | — |
| Description (`description`) | Free text. | For people: where the rules came from, who approved them. | — |
| Host for rules without a host (`testBaseHost`) | Tenant address used only to test the template on its own, outside a profile. Stored in the template. | To test rules like `/Document/…`. Runs use the profile's `baseHost`. | Host of the first rule with a full address; if there is none — `https://base-host.invalid` (placeholder) |
| Old prefix (`rules[].old`) | Start of the old address. | Links starting with it (at a folder boundary, ignoring case and encoding) will be replaced. | — |
| New prefix (`rules[].new`) | What to replace the old prefix with. | The rest of the address is copied unchanged. | — |
| Rule name (`rules[].name`) | Optional label. | Shown in reports, so you know where a rule came from. | — |
| Import rules | CSV, TSV, TXT or XLSX (first sheet). | Two values per line — old and new prefix; a third (if present) — the rule name. Separator — tab, semicolon, comma or `\|`; lines like "old => new", "old -> new", "old → new" are also accepted. A header line (old/new, старий/новий…) is skipped. Existing rules can be replaced or extended. Unrecognized lines are shown as warnings. | — |
| CSV | Export rules. | For approval and transfer. | — |
| Test link → Link | Any link from a file. | Shows the normalized form, matching rules, the result, and whether a repeated run would change it again. Changes nothing. | — |
| Test link → Old document folder | Where the document was located before the migration. | For relative links: they are resolved from this folder. | — |

**Rule set checks** (under the table and next to rows):

| Check | Meaning | Type |
|---|---|---|
| Duplicate | The same rule twice. The second one is ignored. | warning |
| Identical prefixes | Old and new prefixes are the same. The rule is ignored. | warning |
| Conflict | The same old prefix with different new ones. | error |
| Cycle and repeated replacement | A new prefix matches some rule again (including A→B and B→A): a repeated run would change the link once more. | error |
| Empty or invalid prefix, host-less path without `baseHost`, empty set | — | error |

## 3.3. Profile

**Profiles → [name]**. Buttons at the top:
- **JSON** — export;
- **Run…** — a new run with this profile;
- **Validate** — full check, as before a start;
- **Save** — saves even with errors, but such a profile will not start.

Errors are shown as a counter on the tab and next to the fields.

### "General and rules" tab

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Name (`name`) | Profile name. | Runs are started by it (GUI and CLI). | — |
| baseHost (`baseHost`) | Tenant address, e.g. `https://contoso.sharepoint.com`. | Required for rules and links without a host (`/sites/…`, `/Document/…`). Without it such links cannot be resolved and end up as `unresolved`. | empty |
| Rule templates (`templates`) | Select one or more templates. | Template rules are added to the profile rules and validated together (including conflicts between templates). A run gets a copy of the rules as of its start. | — |
| Rules from templates | Read-only table. | Shows exactly what comes from the templates. | — |
| Profile's own rules (`rules`) | Rules for this profile only. | For exceptions and additions to the template. Fields — as in [3.2](#32-rule-template). | — |

### "Targets" tab

Targets table:
- **Files ≈** — count from the latest run of this profile that listed the library (inventory, scan, apply or verify by profile);
- ✎ — edit the target;
- 🗑 — delete the target;
- **Add target** — opens the dialog below.

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Connection for browsing | Which connection to use for browsing sites in the dialog. | For selection only. The run uses the profile's connection. | the profile's first connection |
| Site address (`targets[].siteUrl`) | `https://tenant.sharepoint.com/sites/x`. **Open** — load the site's libraries. | The main way to select a site: with the Sites.Selected permission, search often finds nothing. | — |
| Search site by name | Site search. | Secondary method. | — |
| Library (`targets[].library`) | Library name (Documents, Документи), its URL name (Shared Documents) or full address. | Which library to process. | — |
| Folder (`targets[].folder`) | Path inside the library, e.g. `Fin/2024`. Browsing — **Up** button, click a folder, **Use this folder**. | Process only part of the library (handy for a pilot and batches). Empty — the whole library. | whole library |
| Old address (`targets[].oldLocation`) | Where this library or folder was before the migration, e.g. `https://contoso.sharepoint.com/Document/Fin`. | Relative links (`../Analytics/a.xlsx`) are resolved from the document's old location, because they were written while the file was there. Empty — the old address is derived from the rules (in reverse), if this is unambiguous. If not — relative links are left alone and go to `unresolved`. | from rules |

### "Filters" tab

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Extensions (`filters.extensions`) | .xlsx, .xlsm, .xlsb, .xls. | Files with other extensions are not included in the run. Office lock files (`~$…`) are always skipped. | all four |
| Include only paths (`filters.include`) | Path masks relative to the target, one per line. `*` — any characters within a folder, `**` — any depth, `?` — one character. | Process only some files, e.g. `Fin/**`. The rest — `skipped:filtered`. Empty — all. | empty |
| Exclude paths (`filters.exclude`) | Masks, one per line. | E.g. `**/Archive/**`. Such files — `skipped:filtered`. | empty |
| Size from, KB (`filters.minFileSizeKb`) | Minimum size. | Skip small files (empty templates). Smaller ones — `skipped:filtered`. 0 — no limit. | 0 |
| Size up to, MB (`filters.maxFileSizeMb`) | Maximum size. | Very large files take long to process and need memory. Larger ones — `skipped:too_large`. 0 — no limit. | 500 |
| Modified from (`filters.modifiedAfter`) | Date. | Only files modified on or after this date. The rest — `skipped:filtered`. | empty |
| Modified before (`filters.modifiedBefore`) | Date. | Only files modified before this date. | empty |
| Count files | Button. | How many files match the filters, based on the latest inventory. Without an inventory — a hint to run inventory. | — |

### "Replacement locations" tab

The table shows which formats support each location. An old address in a location that is not changed is still found and goes to residuals.

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Hyperlinks (`places.hyperlinks`) | Hyperlinks on cells and shapes. | Turn off if hyperlinks are deliberately left alone. | on |
| Link Sources (`places.linkSources`) | Links to other workbooks (Data → Edit Links), including the alternate paths Excel stores in the file. | The main location for formulas that pull data from other workbooks. Cached values are not changed; the new source is read when the user refreshes the links. | on |
| Formulas (`places.formulas`) | Addresses in formula text (`HYPERLINK("…")`, `INDIRECT("…")`) and the cached results of these formulas. | Turn off if formulas are left alone. | on |
| Other external links (`places.otherExternal`) | Linked OLE objects, linked images. | — | on |
| Pivot caches, Data connections, Power Query, Cell text (`places.pivotCaches`, `connections`, `powerQuery`, `cellText`) | Not implemented yet. | Visible but disabled. Enabling them in JSON is a profile error. Old addresses in these locations appear in the residuals report. | off |

You cannot turn off all locations (profile error).

### "Performance" tab

**Presets** fill in the fields with one click (then save the profile):

| Preset | Downloads / processors / uploads | RU budget | Excel processes | Windows |
|---|---|---|---|---|
| Cautious | 4 / 2 / 2 | 50% | 1 | no windows |
| Standard | 8 / 4 / 4 | 80% | 2 | no windows |
| Off-hours | 8 / 4 / 4 | 80% in windows; outside windows the run waits | 2 | Mon–Fri 20:00–07:00 and Sat–Sun all day; "Run only in windows" on |

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Concurrent downloads (`performance.downloads`) | How many files are downloaded in parallel. | More — faster, but more load and memory. Actual speed is limited by the RU budget. 1…64. | 8 |
| Processors (CPU) (`performance.processors`) | How many files are processed in parallel on the CPU. | Rarely the bottleneck; more — more memory (≈ 2× the file size per processor). 1…64. | half the cores, at least 2 |
| Concurrent uploads (`performance.uploads`) | How many files are written in parallel. | Writing costs SharePoint more than reading. 1…64. | 4 |
| RU budget, % of limit (`performance.budgetPercent`) | What share of the app's limit to use. | Lower — less impact on others and lower risk of 429; higher — faster. This value applies outside work windows. 1…100. | 80 |
| Request timeout, s (`performance.requestTimeoutSeconds`) | How long to wait for a response to a single request. | After that the request is retried. Minimum 5. | 120 |
| File timeout, min (`performance.fileTimeoutMinutes`) | Maximum time for the full cycle of one file. | Protection against hangs. Exceeding it is a transient error with a retry (at the Excel stage — an Excel error). Must be at least three Excel timeouts (otherwise a warning). | 30 |
| Simple upload up to, MB (`performance.simpleUploadMaxMb`) | Size limit for writing in a single request. | Larger files are written in chunks, resuming after a dropped connection. 1…250. | 60 |
| Excel processes (`performance.excelProcesses`) | How many hidden Excel instances run in parallel for xls and xlsb residuals. | 0 — Excel is not used; such files get `needs_excel`. Each process takes hundreds of MB of memory. 0…16. | 2 |
| Excel timeout, min (`performance.excelTimeoutMinutes`) | Maximum time for one Excel operation (open and save). | After that Excel is considered hung: its process is terminated and the file gets one more attempt. 1…120. | 5 |
| Window time zone (`performance.timeZone`) | Time zone the windows are defined in. | Empty — machine time. | empty |
| Run only in windows (`performance.onlyInWindows`) | Toggle. | Outside windows the run waits and no new files are picked up. Requires at least one window. | off |
| Work windows (`performance.windows`) | List: **Days** (Mon…Sun; empty — every day), **From** and **To** (HH:MM, `24:00` — end of day; a window may cross midnight and belongs to its start day), **Budget, %**. | A separate RU budget for off-hours, automatic pause and resume. | none |

### "Retries" tab

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Attempts per request (`retry.requestAttempts`) | Total number of times a single SharePoint request is made (first attempt + retries on network failures, 5xx, 429 with a pause per `Retry-After`). | More — more resilient to short outages, but longer waits. 1…20. | 6 |
| Attempts per file (`retry.fileAttempts`) | How many times a file is taken again after a transient error or conflict. | After the last one — `failed:transient` / `failed:conflict`. 1…20. | 3 |
| Pauses between file attempts, min (`retry.filePausesMinutes`) | Comma-separated list. | Pause before the 2nd, 3rd… attempt; the last value applies from then on. With 3 attempts the first two pauses are used. | 1, 5, 15 |
| Locked wave, min (`retry.lockedWaveMinutes`) | How often to retry deferred (locked, checked-out) files. | — | 10 |
| Wait for locked, min (`retry.lockedLimitMinutes`) | Maximum time to wait for a file to be released. | After that — `failed:locked`; retry manually later. | 120 |
| Circuit breaker: recent files (`retry.breakerWindow`) | How many recent processing results the circuit breaker considers (intermediate failed attempts count too). | — | 50 |
| Circuit breaker: error ratio (`retry.breakerFailureRatio`) | Share of transient, auth and internal errors in the window that stops the run. | A widespread problem (network, SharePoint, permissions) stops the run instead of "burning" all files into errors. 0…1. | 0.5 |
| Consecutive auth errors before stop (`retry.authFailuresToStop`) | 401/403 in a row. | Revoked access or an expired certificate stops the run. | 10 |
| Only 503 for this many min — block (`retry.blockedMinutes`) | How long only 503 responses indicate a possible app block. | The run stops with the reason "possible SPO block", without further attempts. | 15 |

### "Connections" tab

| Field (JSON) | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Connections (`connections`) | One or more connections. | Which app to work with. More than one — only with a pool. | — |
| App pool (`pool.enabled`) | Several apps in one run. Enabled with a warning. | Each app has its own budget and throttling state. A file is processed from start to finish by one app. A pause of one app does not stop the others. Before the start, each app is checked to see all targets. | off |
| Tenant budget ceiling, % (`pool.tenantCeilingPercent`) | Share of the whole-tenant limit the entire pool may use. | Protects users and other apps. 1…100. | 50 |

## 3.4. New run

**Runs → New run** or **Profile → Run…**

| Field | What it is | Purpose, effect and result |
|---|---|---|
| Profile | Profile for the run. | The profile is validated immediately; errors are shown in the dialog, and the run will not start with them. |
| Mode | See the table below. | Determines whether anything is written. |
| One-off changes for this run | Downloads, processors, uploads, RU budget, Excel processes, include/exclude masks, Hyperlinks / Link Sources / Formulas locations, "Run only in windows". **Reset to profile** — discard the changes. | Change parameters of a single run without changing the profile. The changes go into the run snapshot. |
| Estimated file count | Based on the latest inventory: total Excel files and how many match the filters. | To know the scale before starting. |
| I confirm writing… (apply only) | Checkbox. | apply will not start without it. |

| Mode | What it does | Writes to SharePoint | Result |
|---|---|---|---|
| **inventory** | Lists files, downloads and parses each file, collects all links and a prefix histogram. Costs the same as scan. | No | "Prefixes" report, file count for the profile; statuses as in scan. |
| **scan** | Full in-memory processing. | No | What will be replaced: `planned` statuses, "Links", "Sample", "Residuals" reports. |
| **apply** | Replaces and writes changed files. | Yes | `done` — written, `unchanged` — nothing to change. |
| **verify** | Downloads again and searches for old addresses in all files of the profile. | No | `verified` — clean; `changes_pending` — there is still something to change. |

The **Verify** button on an apply run card checks only the files written by that run ([details](#the-verify-button)).

## 3.5. Retry run

**Run card → Retry…**, error group → ↻, or **Files → Retry selected / Retry all matching filter**.

| Field | What it is | Purpose, effect and result | Default |
|---|---|---|---|
| Statuses | File statuses of the parent run. `failed` — all errors, `done` — also `done:residual` etc. **`attention — needs attention`** — everything counted by the "Needs attention" card: errors, files with residuals (`*:residual`), `needs_excel`, `changes_pending`, `skipped:encrypted`. | Which files to take. Retrying residuals and `skipped:encrypted` makes sense only after extending the rules (the "Use current profile version" checkbox) or removing the password. | `failed` |
| Error classes | transient, conflict, locked, excel, auth, file, validation, internal. | Narrow the selection. | all |
| Extensions | Comma-separated. | E.g. only `.xls` after installing Excel. | all |
| Path mask | `*` and `?`. Unlike profile filters: relative to the library root (not the target) and case-sensitive. | Only some folders, e.g. `Fin/2024/*`. | — |
| Mode | inventory / scan / apply / verify. | E.g. retry scan with new rules before apply. | same as parent |
| Use current profile version | Checkbox. | Retry with extended rules or other parameters. Targets must match the parent run. Without the checkbox — the parent run's snapshot. | off |

Ignored files are not taken. Each file is downloaded again in its latest version. The result is a new run labeled "retry of run N".

## 3.6. Dashboard and run card

**Header:**
- run number, state (running, paused, stopped, aborted, completed, completed with errors), mode, profile, "retry of run N";
- for a stopped or aborted run — the reason.

**Progress:** processed N of M (%), start time, duration, estimated time to completion (at the current speed).

**Cards:**

| Card | What it shows | What to look for |
|---|---|---|
| Files found | How many files are in the run (grows during listing). | — |
| Changed | apply — written to SharePoint; scan — will be changed. | — |
| Still to change (verify) | Files where verify found addresses that apply would still change. | Should be 0. Otherwise — run apply for these files. |
| Unchanged | Files where nothing needs changing. | — |
| Link replacements | How many links were replaced; the tooltip breaks it down by location. | — |
| Errors | Files with `failed:*`; the tooltip breaks it down by class. Clicking the card opens the errors tab with errors only. | [S8](#s8-errors-and-retries) |
| Deferred | Locked files waiting for the next wave. | If there are many during the day — work during off-hours. |
| With residuals | Files with old addresses after replacement. | [S9](#s9-residual-old-addresses) |
| Needs attention | Errors + residuals + needs_excel + changes_pending + skipped:encrypted. Clicking the card opens the errors tab with only these files. | Detailed list — the "Needs attention" report ([3.10](#310-reports)). |
| Speed | Files per minute, MB/s. | — |
| Active processors | How many files are being processed now; throttling state. | "paused" — SharePoint asked to wait. |
| Excel (xls, xlsb) | Busy / started out of N; operations, hangs, crashes, closed dialogs (if any); or "unavailable" with the reason. | Many hangs — increase the Excel timeout or check the files. Closed dialogs — see [S10](#s10-xls-and-xlsb-files-excel). |

**SharePoint load** — a row for each app and a tenant row:

| Column | What it is |
|---|---|
| RU budget per minute | What share of the per-minute budget is in use right now. |
| RU in 24 h, RU total | Spent over the last day and since the run was last started (after "Resume run" the count starts from zero; the total across all starts is in the pilot report). |
| Requests | How many requests were made. |
| 429 / 503 | How many times SharePoint asked to wait or was overloaded. |
| Concurrency | Current limit on concurrent requests (lowered after 429 and gradually restored). |
| Files now | How many files the app is processing. |
| State | running / paused until HH:MM:SS / possible block. |
| Tenant (pool ceiling) | How much of the 5-minute tenant budget the run uses. |

**Trends** (snapshot once a minute):
- "Progress" — files found and processed;
- "Speed and errors" — files per minute between adjacent snapshots (not the average since the start, so a slowdown shows immediately); on the right axis:
  - "Files with error" — files currently in a `failed:*` state (the same number as on the "Errors" card, with up to a minute's delay);
  - "Retries" — failed attempts after which the file went for another one (transient error, locked file, retry after Excel); many retries without growing errors — SharePoint or the network cannot keep up, consider reducing concurrency;
  - 429/503 — since the run was last started;
- hover the mouse — values at that moment appear under the chart;
- select a region with the mouse — both charts zoom into it; the "1 h", "6 h", "24 h" buttons show the latest interval and keep moving with new snapshots; "All" or a double-click — the whole run;
- a gap in the speed line — the run was not working at that time (paused, app closed).

**Breakdowns:**
- by format — table: total, processed, changed, errors;
- by target — a block for each library: name (and folder), site address, share processed, the same numbers.

**Buttons:**
- **Pause / Continue / Stop** — [S11](#s11-pause-stop-closing-the-window-reboot);
- **Start / Resume run** — for a created, stopped or aborted run. For apply — with confirmation;
- **Retry…** — [3.5](#35-retry-run);
- **Verify** — for a completed apply, [see below](#the-verify-button);
- **Reports** — [3.10](#310-reports).

**Run card tabs:**
- **Overview** — the dashboard;
- **Files**, **Errors and attention** — [3.7](#37-files-table-errors-file-card);
- **Run settings** — snapshot of the profile, templates and connections the run works with.

### The "Verify" button

**Purpose:** make sure that SharePoint now holds exactly the fixed files and no old addresses remain in them. The button is on the card of a completed **apply** run (including one completed with errors).

**What happens after clicking:**

1. A new run in `verify` mode labeled "retry of run N" is created and started immediately; its card opens.
2. The run includes **only files with the `done` status** — those this apply wrote. Unchanged, failed, deferred and ignored files are not taken. The library is not listed again.
3. For each file, the **current version from SharePoint** is taken (fresh metadata and download), not the copy from apply. So if someone overwrote the file after apply, verify will see it.
4. The file is processed with the same rules — from the apply run's profile snapshot, not from the current profile. xls and xlsb go through Excel, as in apply ([S10](#s10-xls-and-xlsb-files-excel)).
5. **Nothing is written.** Verify only counts what apply would change.

**Result for each file:**

| Status | Meaning | What to do |
|---|---|---|
| `verified` | Nothing for the rules to change: no old addresses covered by the rules. | — |
| `changes_pending` | There are links apply would change. Most often the file was overwritten after apply (an old copy, a restored version). | **Retry…** with apply mode and the `changes_pending` status. |
| `…:residual` | In addition to the status: the file contains old addresses not covered by any rule. | [S9](#s9-residual-old-addresses) |
| `failed:<class>` | The file could not be verified (access, file deleted, Excel, etc.). | [3.9](#39-error-classes-and-what-to-do) |

The verify dashboard shows a **Still to change** card — the number of `changes_pending`. Expected result: 0 on this card and 0 residuals, except agreed ones.

**Cost:** same as scan over the same files — each file is downloaded and parsed, consuming RU. For a large apply, run verify during off-hours.

**Difference from verify by profile:** **New run → verify** (or `verify --profile <profile>` in the CLI) lists and checks all files of the profile, including those apply did not change. The button on the card checks only what this run wrote. The command-line equivalent is `verify <apply run>` ([3.11](#311-command-line-cli)).

## 3.7. Files table, errors, file card

**Files:**
- **filters** — status (including `attention — needs attention`, as in [3.5](#35-retry-run)), error class, format, target, path search;
- **sorting** — click a column header;
- **actions on selected files** (checkboxes) — **Retry selected**, **Ignore selected** (with a comment);
- **Retry all matching filter** — carries only the selected statuses into the retry dialog; set class, format and mask in the dialog;
- **CSV / XLSX** — export the current view with filters;
- click a row — the file card.

**Errors and attention:** "status + message" groups with the file count, class and recommendation. Group actions: ☰ files, ↻ retry, ⊘ ignore. Switch above the table: **All** (including deferred), **Errors** (`failed:*`), **Needs attention** (what the card of the same name counts). The tab and the switch are reflected in the address, so dashboard cards open their own list.

**File card:**

| Section | What it shows |
|---|---|
| Buttons | **Open in SharePoint**; **Save pre-change version** (copy of the version before writing); **Ignore** (with a comment — the file will not be taken into retries and is counted separately). |
| Metadata | Size, who changed it in SharePoint and when, attempts, app, version before writing, eTag after writing, error class, message, comment. |
| Attempt history | Each attempt: time, number, stage (download, process, excel, upload), result, HTTP code, `request-id` (for Microsoft support), app. |
| Links | Each link found: action, location, file part, old value (and how a relative one was resolved), new value, rule number. |
| Residuals | File part, encoding, match, context. |

Link actions:

| Action | Meaning |
|---|---|
| Replaced | Replaced (or will be replaced, in scan). |
| NoRule | No rule for the address — not changed. |
| Current | Already the new address — not changed. |
| Unresolved | A relative link that could not be resolved (no old target address or `baseHost`). |
| PlaceDisabled | A rule exists, but the replacement location is disabled in the profile — not changed, becomes a residual. |

Link locations: hyperlinks, Link Sources, formulas, formula cache (cached formula result), other external, save path (where Excel last saved the file — updated along with other changes, but on its own is not a reason to write the file).

## 3.8. Statuses

**File statuses — final:**

| Status | Meaning | What to do |
|---|---|---|
| `done` | Changed and written to SharePoint. | — |
| `unchanged` | Nothing needs changing. | — |
| `planned` | scan: will be changed. | apply |
| `verified` | verify: no old addresses. | — |
| `changes_pending` | verify: there are addresses apply would change. | apply for these files |
| `needs_excel` | Excel is required, but it is missing or set to 0 processes. | Retry on a machine with Excel |
| `ignored` | Marked manually with a comment. | — |
| `skipped:filtered` | Did not pass the filters (path, date, minimum size). | — |
| `skipped:too_large` | Larger than the maximum size. | Increase the limit and retry, if needed |
| `skipped:encrypted` | Protected with a password to open or to modify; or Excel asked for a password. | Remove the password and retry ([S10a](#s10a-password-protected-files)) |
| `skipped:deleted` | Deleted or moved after listing. | — |
| `failed:<class>` | Error, see [3.9](#39-error-classes-and-what-to-do). | — |

**Suffixes:** `:residual` — old addresses remain in the file ([S9](#s9-residual-old-addresses)); `:needs_excel` — part of the file can be fixed only by Excel.

**Intermediate statuses:** `pending` (queued), `in_progress` (being processed), `deferred` (locked, waiting for a wave), `waiting_path` (listing has not reached the file's folder yet).

**Run states:**

| State | Meaning |
|---|---|
| created | Not started yet. |
| running / paused | — |
| stopped | Stopped manually; can be resumed. |
| aborted | Stopped by the circuit breaker, or due to a listing, access or block failure; the reason is on the dashboard. Can be resumed after the cause is fixed. |
| completed | All files have a final status, no errors. |
| completed with errors | All files have a final status, there are `failed:*`. |

## 3.9. Error classes and what to do

| Class | Examples | Auto-retry | What to do |
|---|---|---|---|
| transient (`transient`) | Network drop, 5xx, 429 after retries are exhausted, timeout | Yes, up to 3 file attempts | Retry |
| conflict (`conflict`) | 412: the file was changed during processing | Yes, with the new version | Retry |
| locked (`locked`) | 423, check-out, file open | Yes, in waves for up to 2 h | Ask the owner to close the file or check it in; retry later |
| Excel (`excel`) | Excel hung or crashed | Yes, once in a new process | Check the file manually; increase the Excel timeout |
| auth (`auth`) | 401, 403 | No | Check the app's permissions on the site and the certificate; retry |
| file (`file`) | Corrupted, not Excel, unsupported format | No | Check manually or ignore |
| self-check (`validation`) | The result failed validation; a link to a local drive in xls/xlsb ([S10](#s10-xls-and-xlsb-files-excel)) | No | The file was not changed. Read the message; pass the file to the developer |
| internal (`internal`) | An error in the app itself | No | Pass the log to the developer; retry |

## 3.10. Reports

**Run card → Reports** (CSV or XLSX; summary and pilot report — TXT) or `spo-replace-link export <run> -o <folder>`.

| Report | Contents | Purpose |
|---|---|---|
| Files (`files`) | `id`, `path`, `format`, `size`, `status`, `error_class`, `error`, `attempts`, `replacements`, `residuals`, `version_before`, `etag_after`, `app`, `comment` | Complete list and archive; `version_before` — for rollback. |
| Links (`links`) | `path`, `part`, `kind`, `action`, `old_value`, `new_value`, `rule`, `resolved_from` | What was or will be replaced, and by which rule. |
| Residuals (`residuals`) | `path`, `part`, `encoding`, `match`, `context` | What was left unfixed and where. |
| Needs attention (`attention`) | `id`, `path`, `status`, `error_class`, `message`, `attempts`, `action` | Action list: what to do with each file. |
| Prefixes (`prefixes`) | `prefix`, `links`, `files`, `replaced`, `no_rule`, `unresolved` | Extending the rules ([S4](#s4-inventory-and-completing-the-rules)). |
| Verification sample (`sample`) | `id`, `path`, `status`, `replacements`, `residuals`, `version_before`, `url`, `examples` (up to three replacements), `checked`, `comment` | Manual check of 30 changed files. The selection is random, but always the same for the same run. |
| Summary (`summary`) | Duration, files, volume, replacements, speed, counts by status | Short report. |
| Pilot report (`pilot`) | Files by format; RU, 429/503, RU per file (across all starts of the run); speed over active processing time, s per file, RU budget limit, forecast for 10,000 files and for the profile; SharePoint behavior on writes; attempts with conflict and lock; Excel; errors with recommendations; what to check manually | Pilot conclusion and planning of the production run ([S6](#s6-pilot-on-a-single-library)). |

## 3.11. Command line (CLI)

`spo-replace-link <command>`. Commands that work with SharePoint are executed through the app, which must be open **under the same user account**. The CLI takes the address and API key from `api.json` in the app's data folder (`--data`, variable `SPO_RL_DATA`, default `%LOCALAPPDATA%\SPOReplaceLink`; if the app was started with `--data`, pass the same folder). `--host` (variable `SPO_RL_HOST`) sets the address explicitly; without `api.json` — `http://localhost:5080`.

| Command | What it does |
|---|---|
| `profiles\|connections\|templates list` | List. |
| `… get <name> [-o file.json]` | Show or export. |
| `… import <file.json> [--name]` | Create or replace. |
| `… delete <name>` | Delete. |
| `profiles validate <name>` | Validate a profile as before a start. |
| `connections test <name> [--profile <profile>]` | Token and access to the profile's targets. |
| `inventory\|scan\|apply <profile> [--wait] [--yes]` | New run; `--wait` — wait, showing progress; `--yes` — no confirmation prompt for apply. |
| `verify <apply run>` or `verify --profile <profile>` | Verify the files of a run or the whole profile. |
| `retry <run> [--status …] [--class …] [--ext …] [--path-mask …] [--file-id …] [--list file.csv] [--mode …] [--current-profile]` | Retry run by selection. `--status attention` — all files that need attention. `--list` — CSV with paths (the `path` column or the first one). |
| `runs [--take N]`, `status <run>`, `wait <run>` | List, state, waiting. |
| `start\|pause\|resume\|stop <run>` | Control. |
| `files <run> [--status] [--class] [--search] [--take]` | Run files. |
| `ignore <file> --comment "…"` | Ignore a file. |
| `delete-run <run> [--yes]` | Delete a run that is not running, with all its results ([S11](#s11-pause-stop-closing-the-window-reboot)). Asks for confirmation. |
| `export <run> -o <folder> [--prefix-depth N]` | All reports ([3.10](#310-reports)); `--prefix-depth` — prefix depth in `prefixes` (1–10, default 2). |
| `rules validate -c profile.json`, `rules test -c profile.json <link…> [--base-url]` | Check rules without the app. |
| `local <folder> -c profile.json [--write --out] [--base-url] [--excel N] [--temp] [--report] [--parallel] [--csv-delimiter]` | Local processing ([S14](#s14-processing-a-local-folder-without-sharepoint)). |

Ctrl+C during `--wait` stops the run gracefully. Exit codes: `0` — OK, `1` — something needs attention, `2` — configuration error or app unavailable, `3` — run stopped or aborted.

## 3.12. App files and folders

In the data folder (`--data`):

| Path | What it is |
|---|---|
| `state.db` | Database: settings, runs, files, attempts, replacements, residuals, metrics. Back it up while the app is closed. |
| `logs\host-*.jsonl` | Log (JSON lines): every error with details. Send to the developer when there are problems. |
| `temp\` | Temporary file copies; each file is deleted right after processing. |
| `temp\excel-*.pids` | Internal: Excel processes started by the app (so they can be cleaned up after a crash). |
| `api.json` | Address and key of the running instance's local API — for the CLI. Created on start, deleted on exit. |

## 3.13. Limitations

- **Not fixed** (only reported in residuals): pivot caches, data connections, Power Query, cell text, paths in VBA code, links like `AllItems.aspx?id=…` and `/:x:/r/…`.
- **The Modified and Modified By fields change:** writing a file creates a new version on behalf of the app.
- **No bulk rollback** — only manually via version history ([S12](#s12-restoring-a-previous-version-of-a-file)).
- **Prefix rules only** (no regular expressions or variables).
- **Import from a ReplaceMagic list** reads two columns or "old → new"; the exact Export list format is still being clarified.
- **Passwords and Excel.** Password-protected files are not processed, only marked for retry ([S10a](#s10a-password-protected-files)). xls and some xlsb files need Excel on the machine, so the app runs in a user session rather than as a Windows service.
- **xls/xlsb with a link to a local drive** — see [S10](#s10-xls-and-xlsb-files-excel).
- **Window access** — local machine only; there are no roles or Entra sign-in for the GUI.

## 3.14. Troubleshooting

| Situation | Cause and what to do |
|---|---|
| "Certificate … not found in store" | The app runs under a different account than the one the certificate was created under, or the wrong store is selected. |
| "without a private key" | The selected store has only an imported `.cer` without a key. Select the store where the certificate was created with `New-SelfSignedCertificate`, or create a new certificate. |
| "No access to the private key" | The certificate is in `LocalMachine`, and the app's account has not been granted Read on the key (Manage Private Keys). |
| `AADSTS700027`, `AADSTS700016` | The certificate is not uploaded to the Entra app, or the Client ID or tenant is wrong. |
| Token is acquired, but `403` on a target | No admin consent for Sites.Selected, or `write` on the site has not been granted. |
| Site search finds nothing | Normal for Sites.Selected — enter the site address. |
| Run "aborted: Circuit breaker…" | A widespread problem: network, SharePoint, permissions. Read the reason, fix it, **Resume run**. |
| "possible SPO block" | SharePoint has been responding only with 503 for a long time. Do not resume right away; wait, lower the RU budget, work during off-hours. |
| Many 429s | Lower the RU budget, check the license range in the connection, find out whether other heavy apps are running in parallel (migration, backup). |
| Many `deferred` / `failed:locked` | Files are open by users. Work during off-hours; retry later. |
| `needs_excel` | Excel is not installed, not activated for this account, or the profile has 0 processes. |
| `failed:validation` in xls/xlsb with a "drive letter" | [S10](#s10-xls-and-xlsb-files-excel): the data folder is on a different drive. |
| CLI: "Host unavailable" | The app is not running under this account, or it was started with a different `--data` (pass the same one to the CLI), or an explicit `--host` points elsewhere. |
| CLI: "No access key for this instance" | The app at this address belongs to another user or another data folder. Start the app under your own account and drop `--host`, or specify the same `--data`. |
| Window does not open, starting again does nothing | The app is already running (possibly in the tray). Double-click the icon next to the clock. |
| Anything unclear | `spo-replace-link export <run> -o …` and `logs\host-*.jsonl` from the data folder — to the developer. |
