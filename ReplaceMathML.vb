Sub ReplaceMathML()

Dim nNumber As Integer
Dim strMath As String
Dim objData As New MSForms.DataObject
Dim oRange As Range

If ActiveDocument.OMaths.Count > 0 Then
    With ActiveDocument
        For nNumber = 1 To .OMaths.Count Step 1
            .OMaths(1).Range.Select
            Set oRange = .OMaths(1).Range
            Selection.Copy
            objData.GetFromClipboard
            strMath = Replace(objData.GetText(), "mml:", "")
            strMath = Replace(strMath, "xmlns:mml=""http://www.w3.org/1998/Math/MathML"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math""", "")
            oRange.Select
            Selection.Next.Select
            .OMaths(1).Range.Select
            Selection.Delete
            Selection.EndOf
            oRange.Select
            Selection.InsertAfter (" " + strMath + " ")
        Next nNumber

    End With

Else
    MsgBox ("No math")

End If

End Sub
