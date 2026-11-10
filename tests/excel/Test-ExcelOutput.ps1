<#
.SYNOPSIS
  Opens processed workbooks in Excel and checks the stage 1 readiness criterion:
  the file opens without repair, Edit Links and hyperlinks point to the new location,
  and the external reference reads the migrated source.

.PARAMETER Path
  Workbooks to check (output of `local --write` on fixtures from New-ExcelFixtures.ps1).

.PARAMETER Expect
  Text every link source and every non-internal hyperlink must start with (one of).

.PARAMETER ExpectedValue
  Value of Sheet1!A4 after the external link is updated from its (new) source; skip with $null.
#>
param(
    [Parameter(Mandatory)] [string[]] $Path,
    [Parameter(Mandatory)] [string[]] $Expect,
    $ExpectedValue = $null,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$failed = 0

function Test-Prefix([string] $value) {
    foreach ($e in $Expect) { if ($value.StartsWith($e, [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    return $false
}

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    foreach ($file in $Path) {
        $problems = @()
        $wb = $null
        try {
            # UpdateLinks = 0, ReadOnly. CorruptLoad stays at its default (xlNormalLoad).
            $wb = $xl.Workbooks.Open($file, 0, $true)
        }
        catch {
            $problems += "не відкривається: $($_.Exception.Message)"
        }

        if ($wb) {
            if ($wb.Name -match 'Repaired|Відновлено') { $problems += 'Excel відновлював файл' }

            # LinkSources returns a 1-based COM array; enumerate it instead of indexing.
            $raw = $wb.LinkSources(1)
            $sources = @($raw | ForEach-Object { $_ })
            if ($sources.Count -eq 0) { $problems += 'немає зовнішніх зв''язків (Edit Links порожній)' }
            foreach ($s in $sources) {
                if ($s -and -not (Test-Prefix $s)) { $problems += "Edit Links: $s" }
            }

            $ws = $wb.Worksheets.Item(1)
            foreach ($h in $ws.Hyperlinks) {
                if ($h.Address -and -not (Test-Prefix $h.Address)) { $problems += "гіперпосилання $($h.Range.Address(0,0)): $($h.Address)" }
            }
            $formula = $ws.Range('A3').Formula
            if ($formula -match 'HYPERLINK\("([^"]+)"' -and -not (Test-Prefix $Matches[1])) { $problems += "формула A3: $formula" }

            if ($null -ne $ExpectedValue -and $sources.Count -gt 0) {
                $wb.UpdateLink($raw, 1)
                $value = $ws.Range('A4').Value2
                if ($value -ne $ExpectedValue) { $problems += "A4 після оновлення зв'язку = $value, очікувалось $ExpectedValue" }
            }

            if (-not $Quiet) {
                Write-Host "  Edit Links: $($sources -join '; ')"
                foreach ($h in $ws.Hyperlinks) { Write-Host "  $($h.Range.Address(0,0)): $($h.Address)$(if ($h.SubAddress) { '#' + $h.SubAddress })" }
                Write-Host "  A3: $formula"
            }
            $wb.Close($false)
        }

        if ($problems.Count -eq 0) {
            Write-Host "OK    $file" -ForegroundColor Green
        }
        else {
            $failed++
            Write-Host "FAIL  $file" -ForegroundColor Red
            $problems | ForEach-Object { Write-Host "      $_" -ForegroundColor Red }
        }
    }
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $failed



