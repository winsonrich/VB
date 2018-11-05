Sub copy()
Dim r As Long
    Dim SourcePath As String
    Dim dstPath As String
    Dim myFile As String
    On Error GoTo ErrHandler
    For r = 2 To Range("A" & Rows.Count).End(xlUp).Row
    SourcePath = Range("C" & r)
    dstPath = Range("D" & r)
        myFile = Range("A" & r)
        FileCopy SourcePath & "\" & myFile, dstPath & "\" & myFile
        If Range("A" & r) = "" Then
           Exit For
        End If
    Next r
        MsgBox "The file(s) can found in: " & vbNewLine & dstPath, , "COPY COMPLETED"
ErrHandler:
    MsgBox "Copy error: " & SourcePath & "\" & myFile & vbNewLine & vbNewLine & _
    "File could not be found in the source folder", , "MISSING FILE(S)"
Range("A" & r).copy Range("F" & r)
Resume Next
End Sub
