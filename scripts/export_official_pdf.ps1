$word = New-Object -ComObject Word.Application
$word.Visible = $false
$docPath = (Resolve-Path "output/1160102_吳修齊基金會_重溪國小第一期款函復公文_v1.docx").Path
$pdfPath = [System.IO.Path]::ChangeExtension($docPath, ".pdf")
$doc = $word.Documents.Open($docPath)
$pages = $doc.ComputeStatistics(2)
Write-Output "Official doc pages: $pages"
$doc.ExportAsFixedFormat($pdfPath, 17)
$doc.Close([ref]0)
$word.Quit()
Write-Output "Exported official doc PDF to: $pdfPath"
