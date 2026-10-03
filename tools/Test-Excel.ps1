param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'
$sourcePath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$projectPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testDirectory = Join-Path $projectPath ('.build\native-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testDirectory)
$copyPath = Join-Path $testDirectory 'native-test.xlsm'
Copy-Item -LiteralPath $sourcePath -Destination $copyPath

$excelApp = $null
$excelBook = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.EnableEvents = $false
    # Honor the existing Office security policy. No registry or trust changes.
    $excelBook = $excelApp.Workbooks.Open($copyPath, 0, $false)
    $macroPrefix = "'" + $excelBook.Name.Replace("'", "''") + "'!"
    [void]$excelApp.Run($macroPrefix + 'ProtectSheets')
    $settings = $excelBook.Worksheets.Item(6)
    $board = $excelBook.Worksheets.Item(1)
    $archive = $excelBook.Worksheets.Item(2)
    $editor = $excelBook.Worksheets.Item(3)
    $data = $excelBook.Worksheets.Item(4)
    $newMarker = [string]$editor.Range('B4').Value2
    if ($newMarker -ne ([char]0x65B0).ToString() + ([char]0x898F).ToString()) {
        throw 'Run this test on the unused empty workbook.'
    }
    $taskId = $excelApp.Run($macroPrefix + 'StoreTask', $newMarker, 'Native smoke test', "line1`nline2", '=1+1', [string]$settings.Cells.Item(1, 1).Value2)
    $created = $data.Range('F7').Value2
    for ($index = 1; $index -le 5; $index++) {
        $status = [string]$settings.Cells.Item($index, 1).Value2
        $returnedId = $excelApp.Run($macroPrefix + 'StoreTask', $taskId, 'Native smoke test', "line1`nline2", '=1+1', $status)
        if ($returnedId -ne $taskId) { throw 'Update changed the ID.' }
        [void]$excelApp.Run($macroPrefix + 'RenderBoards')
        $targetSheet = if ($index -eq 5) { $archive } else { $board }
        $card = $targetSheet.Shapes.Item('card_' + $taskId)
        $expectedColor = $excelApp.Run($macroPrefix + 'StatusColor', $status)
        if ($card.Fill.ForeColor.RGB -ne $expectedColor) { throw 'Unexpected card color.' }
        if ($card.TextFrame2.TextRange.Font.Name -ne 'Meiryo UI') { throw 'Unexpected font.' }
        if ($data.Range('E7').Value2 -ne $status) { throw 'Unexpected stored status.' }
        if ($data.Range('F7').Value2 -ne $created) { throw 'Creation time changed.' }
        if ($data.Range('D7').HasFormula) { throw 'Text was stored as a formula.' }
    }
    if ($board.Shapes.Count -ne 4 -or $archive.Shapes.Count -ne 5) { throw 'Completed card routing failed.' }
    [void]$excelApp.Run($macroPrefix + 'StoreTask', $taskId, 'Reopened', "line1`nline2", '=1+1', [string]$settings.Cells.Item(2, 1).Value2)
    [void]$excelApp.Run($macroPrefix + 'RenderBoards')
    if ($board.Shapes.Count -ne 5 -or $archive.Shapes.Count -ne 4) { throw 'Reopening routing failed.' }
    if ($data.Range('H7').Value2) { throw 'Completion date was not cleared.' }
    $excelBook.Save()
    $excelBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelBook)
    $excelBook = $excelApp.Workbooks.Open($copyPath, 0, $false)
    $data = $excelBook.Worksheets.Item(4)
    if ($data.Range('A7').Value2 -ne $taskId -or $data.Range('C7').Value2 -ne "line1`nline2") {
        throw 'Saved data did not survive reopening.'
    }
    Write-Output ('PASS: native Excel status/color/archive/reopen/persistence checks. Copy: ' + $copyPath)
} finally {
    if ($null -ne $excelBook) {
        $excelBook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelBook)
    }
    if ($null -ne $excelApp) {
        $excelApp.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
