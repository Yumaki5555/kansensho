param([string]$Path)
# Open in Excel, recalculate every formula, list error cells and the summary sheet, then save.
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $xl.CalculateFull()
    $formulas = 0; $errors = @()
    foreach ($ws in $wb.Worksheets) {
        foreach ($c in $ws.UsedRange.Cells) {
            if ($c.HasFormula) {
                $formulas++
                if ($c.Text -like '#*') { $errors += "$($ws.Name)!$($c.Address($false,$false)) $($c.Formula) -> $($c.Text)" }
            }
        }
    }
    "formulas: $formulas  errors: $($errors.Count)"
    $errors | Select-Object -First 30
    $sm = $wb.Worksheets.Item(2)
    for ($r = 4; $r -le 8; $r++) {
        $vals = @()
        for ($col = 1; $col -le 14; $col++) { $vals += $sm.Cells.Item($r, $col).Text }
        $vals -join ' | '
    }
    $wb.Save()
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
