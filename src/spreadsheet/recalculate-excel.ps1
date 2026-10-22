$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class AhmedExcelPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p); }'
    [uint32]$excelProcessId = 0
    [AhmedExcelPid]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId) | Out-Null
    [IO.File]::WriteAllText($env:AHMED_XLSX_PID, [string]$excelProcessId)
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($env:AHMED_XLSX_INPUT, 0, $true)
    $excel.CalculateFullRebuild()
    $book.SaveCopyAs($env:AHMED_XLSX_OUTPUT)
} finally {
    if ($book) { $book.Close($false); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null; Remove-Item -LiteralPath $env:AHMED_XLSX_PID -ErrorAction SilentlyContinue }
}
