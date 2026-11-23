$docxPath = "D:\pba_21\III-ML-PBA-Report-Completed.docx"
$pdfPath = "D:\pba_21\III-ML-PBA-Report-Completed.pdf"

Write-Host "Starting Word COM application..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    Write-Host "Opening document: $docxPath"
    $doc = $word.Documents.Open($docxPath)
    
    Write-Host "Computing document statistics..."
    $pages = $doc.ComputeStatistics(2)
    Write-Host "Total Pages: $pages"
    
    # Check find headings and their page numbers
    $headings = @("CHAPTER 1", "1.1 PROBLEM STATEMENT", "1.2 PROBLEM UNDERSTANDING", "CHAPTER 2", "2.1 REVIEW OF LITERATURE SURVEY", "CHAPTER 3", "3.1.1 Dataset Gathering and Description", "CHAPTER 4", "4.1 PROPOSED ALGORITHM", "CHAPTER 5", "5.1 TECHNOLOGIES", "CHAPTER 6", "6.1 RESULT ANALYSIS", "6.2 INTERPRETATION", "6.4 OUTPUT SCREENS", "CHAPTER 7", "CHAPTER 8", "CHAPTER 9")
    
    foreach ($h in $headings) {
        $range = $doc.Content
        $find = $range.Find
        $find.Text = $h
        $find.Forward = $true
        $find.Wrap = 0
        if ($find.Execute()) {
            $pageNum = $range.Information(3) # wdActiveEndPageNumber = 3
            Write-Host "Heading '$h' -> Page $pageNum"
        }
    }
    
    Write-Host "Exporting to PDF: $pdfPath"
    $wdFormatPDF = 17
    $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
    Write-Host "PDF export complete!"
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    if ($doc) {
        $doc.Close([ref]0)
    }
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
