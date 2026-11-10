<#
.SYNOPSIS
  Creates real Excel workbooks with links to the old location (needs Excel installed).

.DESCRIPTION
  Layout under -Root:
    old\Document\Analytics\Source.xlsx      external source before migration (A1 = 111)
    new\sites\finance\analytics_migrated\Source.xlsx  the same source after migration (A1 = 222)
    input\Fin\Report.xlsx|.xlsm|.xlsb|.xls workbooks to fix; they lived in old\Document\Fin
    input\Fin\Locked.xlsx                  encrypted with a password to open
    profile.json                           rules for this layout

  Each report has: two cell hyperlinks to https://contoso.sharepoint.com/Document/…,
  a HYPERLINK() formula, and an external reference to old\Document\Analytics\Source.xlsx.
#>
param(
    [Parameter(Mandatory)] [string] $Root
)

$ErrorActionPreference = 'Stop'
$Root = [IO.Path]::GetFullPath($Root)
if (Test-Path $Root) { Remove-Item -Recurse -Force $Root }

$oldFin = New-Item -ItemType Directory -Force (Join-Path $Root 'old\Document\Fin')
$oldSrc = New-Item -ItemType Directory -Force (Join-Path $Root 'old\Document\Analytics')
$newSrc = New-Item -ItemType Directory -Force (Join-Path $Root 'new\sites\finance\analytics_migrated')
$inputDir = New-Item -ItemType Directory -Force (Join-Path $Root 'input\Fin')

$host_ = 'https://contoso.sharepoint.com'
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false

    foreach ($pair in @(@($oldSrc, 111), @($newSrc, 222))) {
        $wb = $xl.Workbooks.Add()
        $wb.Worksheets.Item(1).Name = 'Sheet1'
        $wb.Worksheets.Item(1).Range('A1').Value2 = $pair[1]
        $wb.SaveAs((Join-Path $pair[0] 'Source.xlsx'), 51)
        $wb.Close($false)
    }

    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Sheet1'
    $null = $ws.Hyperlinks.Add($ws.Range('A1'), "$host_/Document/Fin/Звіт 2024.xlsx", '', '', 'Звіт 2024')
    $null = $ws.Hyperlinks.Add($ws.Range('A2'), "$host_/Document/Analytics/Аналіз.xlsx", 'Sheet1!A1', '', 'Аналіз')
    $ws.Range('A3').Formula = "=HYPERLINK(""$host_/Document/Analytics/b.xlsx"",""b"")"
    $ws.Range('A4').Formula = "='$($oldSrc.FullName)\[Source.xlsx]Sheet1'!`$A`$1"
    $ws.Range('A5').Value2 = 'internal'
    $null = $ws.Hyperlinks.Add($ws.Range('A5'), '', 'Sheet1!A4', '', 'internal')

    # Save in the old location first so Excel records the external link relative to it, then copy.
    $wb.SaveAs((Join-Path $oldFin 'Report.xlsx'), 51)
    $wb.SaveAs((Join-Path $oldFin 'Report.xlsm'), 52)
    $wb.SaveAs((Join-Path $oldFin 'Report.xlsb'), 50)
    $wb.SaveAs((Join-Path $oldFin 'Report.xls'), 56)
    $wb.SaveAs((Join-Path $oldFin 'Locked.xlsx'), 51, 'secret')
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-ChildItem $oldFin | Copy-Item -Destination $inputDir

$profileData = [ordered]@{
    name     = 'excel-fixtures'
    baseHost = $host_
    rules    = @(
        @{ old = "$host_/Document"; new = "$host_/sites/finance" },
        @{ old = "$host_/Document/Analytics"; new = "$host_/sites/finance/analytics_migrated" },
        @{ old = (Join-Path $Root 'old\Document'); new = (Join-Path $Root 'new\sites\finance') },
        @{ old = (Join-Path $Root 'old\Document\Analytics'); new = (Join-Path $Root 'new\sites\finance\analytics_migrated') }
    )
}
$profileData | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $Root 'profile.json')
Write-Host "Fixtures: $Root"

