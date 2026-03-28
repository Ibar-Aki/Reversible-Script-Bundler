[CmdletBinding()]
param(
    [string]$OutputDir
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $OutputDir) {
    $OutputDir = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'samples'
}

Ensure-Directory -Path $OutputDir
$excel = $null
$workbook = $null
$sheet1 = $null
$sheet2 = $null
$sheet3 = $null
$window = $null

try {
    $excel = New-ExcelApplication
    $workbook = $excel.Workbooks.Add()

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet1.Name = 'Summary'
    $sheet1.Range('A1').Value2 = 'Item'
    $sheet1.Range('B1').Value2 = 'Qty'
    $sheet1.Range('C1').Value2 = 'Price'
    $sheet1.Range('D1').Value2 = 'Total'
    $sheet1.Range('E1').Value2 = 'HiddenCol'
    $sheet1.Range('A2').Value2 = 'Desk'
    $sheet1.Range('B2').Value2 = 2
    $sheet1.Range('C2').Value2 = 15000
    $sheet1.Range('D2').Formula = '=B2*C2'
    $sheet1.Range('A3').Value2 = 'Chair'
    $sheet1.Range('B3').Value2 = 4
    $sheet1.Range('C3').Value2 = 8000
    $sheet1.Range('D3').Formula = '=B3*C3'
    $sheet1.Range('A5:C5').Merge()
    $sheet1.Range('A5').Value2 = 'Merged Header'
    $sheet1.Range('A6').Value2 = 'Visible'
    $sheet1.Range('A6').Interior.Color = 65535
    $sheet1.Range('A6').Borders.Item(7).LineStyle = 1
    $sheet1.Rows.Item(7).Hidden = $true
    $sheet1.Columns.Item(5).Hidden = $true
    $sheet1.Range('A8').Value2 = 'Example'
    $sheet1.Hyperlinks.Add($sheet1.Range('A8'), 'https://example.com') | Out-Null
    $sheet1.Range('B8').AddComment('Legacy comment') | Out-Null
    [void]$sheet1.Activate()
    $window = $excel.ActiveWindow
    $window.SplitRow = 1
    $window.SplitColumn = 1
    $window.FreezePanes = $true

    $sheet2 = $workbook.Worksheets.Add()
    $sheet2.Name = 'WideTable'
    for ($column = 1; $column -le 100; $column++) {
        $sheet2.Cells.Item(1, $column).Value2 = "Col$column"
    }
    for ($row = 2; $row -le 51; $row++) {
        for ($column = 1; $column -le 100; $column++) {
            if ($column -eq 100) {
                $startAddress = Convert-CoordinateToA1 -Row $row -Column 1
                $endAddress = Convert-CoordinateToA1 -Row $row -Column 3
                $sheet2.Cells.Item($row, $column).Formula = "=$startAddress&""-""&$endAddress"
            }
            else {
                $sheet2.Cells.Item($row, $column).Value2 = "{0:D2}-{1:D3}" -f $row, $column
            }
        }
    }

    $sheet3 = $workbook.Worksheets.Add()
    $sheet3.Name = 'Calc'
    $sheet3.Range('A1').Value2 = 10
    $sheet3.Range('A2').Value2 = 20
    $sheet3.Range('A3').Formula = '=SUM(A1:A2)'
    $sheet3.Range('B1').Value2 = '2026-03-10'
    $sheet3.Range('B2').Formula = '=TEXT(TODAY(),"yyyy-mm-dd")'

    $xlsxPath = Join-Path $OutputDir 'sample.xlsx'
    $xlsmPath = Join-Path $OutputDir 'sample.xlsm'
    $workbook.SaveAs($xlsxPath, 51)
    $workbook.SaveAs($xlsmPath, 52)

    Write-Host "Created sample workbook -> $xlsxPath"
    Write-Host "Created sample macro workbook -> $xlsmPath"
}
catch {
    throw "create_sample_workbook.ps1 line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
}
finally {
    if ($null -ne $window) {
        Release-ComReference $window
    }
    if ($null -ne $sheet3) {
        Release-ComReference $sheet3
    }
    if ($null -ne $sheet2) {
        Release-ComReference $sheet2
    }
    if ($null -ne $sheet1) {
        Release-ComReference $sheet1
    }
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
        }
        Release-ComReference $workbook
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        Release-ComReference $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
