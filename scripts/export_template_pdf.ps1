$word = New-Object -ComObject Word.Application
$word.Visible = $false
$docPath = (Resolve-Path "templates/重溪國小_正式信函公版_防裁切標準格式.docx").Path
$pdfPath = [System.IO.Path]::ChangeExtension($docPath, ".pdf")
$doc = $word.Documents.Open($docPath)
$pages = $doc.ComputeStatistics(2)
Write-Output "Template pages: $pages"
$doc.ExportAsFixedFormat($pdfPath, 17)
$doc.Close([ref]0)
$word.Quit()
Write-Output "Exported template PDF to: $pdfPath"
