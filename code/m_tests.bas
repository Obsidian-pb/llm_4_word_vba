Attribute VB_Name = "m_tests"
' Module: m_tests - Empty module for testing purposes


Public Sub GetDocText()
Dim rng As Word.Range
    
    Set rng = Application.ActiveDocument.Range
    
    Debug.Print rng.text
End Sub


Public Sub GetDocParagraphs()
Dim rng As Word.Range
Dim par As Word.Paragraph
    
    For Each par In Application.ActiveDocument.Paragraphs
        Debug.Print par.Range.Start & " - " & par.Range.text
    Next par
End Sub
