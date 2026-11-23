$docxPath = "D:\pba_21\III-ML-PBA-Report-Completed.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath)
    $targets = @(
        "CHAPTER 1",
        "1.1 PROBLEM STATEMENT",
        "1.2 PROBLEM UNDERSTANDING",
        "CHAPTER 2",
        "2.1 REVIEW OF LITERATURE SURVEY",
        "CHAPTER 3",
        "3.1.1 Dataset Gathering and Description",
        "CHAPTER 4",
        "4.1 PROPOSED ALGORITHM",
        "CHAPTER 5",
        "5.1 TECHNOLOGIES",
        "CHAPTER 6",
        "6.1 RESULT ANALYSIS",
        "6.2 INTERPRETATION",
        "6.3 ITERATIVE IMPROVEMENTS / REFLECTION",
        "6.4 OUTPUT SCREENS",
        "CHAPTER 7",
        "CHAPTER 8",
        "CHAPTER 9"
    )
    
    foreach ($t in $targets) {
        $range = $doc.Content
        $range.Find.ClearFormatting()
        $range.Find.Forward = $true
        $range.Find.Wrap = 0
        $range.Find.Text = $t
        while ($range.Find.Execute()) {
            $absPage = $range.Information(1) # wdActiveEndAdjustedPageNumber
            $docPage = $range.Information(3) # wdActiveEndPageNumber
            if ($docPage -gt 12) {
                Write-Host ($t + " -> Absolute Doc Page: " + $docPage + " | Section Relative Page: " + $absPage)
                break
            }
        }
    }
}
finally {
    $doc.Close([ref]0)
    $word.Quit()
}
