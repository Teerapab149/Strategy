$ErrorActionPreference = 'Stop'
$inputDocx = 'E:\strategy\PG_Functional_Uniform_Market_Research_TH_SEA_2026.docx'
$outputPdf = 'E:\strategy\deep_research_functional_uniform_2026\rendered\PG_Functional_Uniform_Market_Research_TH_SEA_2026.pdf'
$outputDir = Split-Path -Parent $outputPdf
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($inputDocx, $false, $true)
    try {
        $doc.ExportAsFixedFormat($outputPdf, 17)
    }
    finally {
        $doc.Close($false)
    }
}
finally {
    $word.Quit()
}

Write-Output $outputPdf
