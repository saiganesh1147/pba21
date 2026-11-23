$docxPath = "D:\pba_21\III-ML-PBA-Report-Completed.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath)
    
    Write-Host "Iterating paragraphs to find exact chapter headings..."
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $txt = $p.Range.Text.Trim()
        if ($txt -match "^CHAPTER\s+[1-9]") {
            $absPage = $p.Range.Information(1) # wdActiveEndAdjustedPageNumber
            $docPage = $p.Range.Information(3) # wdActiveEndPageNumber
            # Only report if it's beyond the TOC (docPage > 12)
            if ($docPage -gt 12) {
                Write-Host ("FOUND HEADING: '" + ($txt -replace '\r', ' ') + "' at Absolute Doc Page: " + $docPage + " | Section Relative Page: " + $absPage)
            }
        }
    }
}
finally {
    $doc.Close([ref]0)
    $word.Quit()
}
