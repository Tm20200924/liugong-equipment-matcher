param($docPath)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docPath)
$text = $doc.Content.Text
$doc.Close()
$word.Quit()

$txtPath = $docPath + ".txt"
[System.IO.File]::WriteAllText($txtPath, $text, [System.Text.UTF8Encoding]::new($false))
Write-Host "OK: $txtPath"
