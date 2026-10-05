$ErrorActionPreference = 'Stop'

$outputDirectory = 'E:\strategy\research\render_compare'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$documents = @(
    'E:\strategy\research\render_compare\pg_whole.docx',
    'E:\strategy\research\render_compare\pg_megu.docx'
)

$wordApplication = New-Object -ComObject Word.Application
$wordApplication.Visible = $false
$wordApplication.DisplayAlerts = 0

try {
    foreach ($documentPath in $documents) {
        $document = $wordApplication.Documents.Open($documentPath, $false, $true)
        try {
            $document.Repaginate()
            $pageCount = $document.ComputeStatistics(2)
            $pdfPath = Join-Path $outputDirectory (([System.IO.Path]::GetFileNameWithoutExtension($documentPath)) + '.pdf')
            $document.ExportAsFixedFormat($pdfPath, 17)
            Write-Output (([System.IO.Path]::GetFileName($documentPath)) + "`t" + $pageCount + " pages`t" + $pdfPath)
        }
        finally {
            $document.Close($false)
        }
    }
}
finally {
    $wordApplication.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApplication) | Out-Null
}
