'Sourced from https://word.tips.net/T001690_Removing_All_Text_Boxes_In_a_Document.html

Sub RemoveTextBox2()
    Dim shp As Shape
    Dim oRngAnchor As Range
    Dim sString As String

    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoTextBox Then
            ' copy text to string, without last paragraph mark
            sString = Left(shp.TextFrame.TextRange.Text, _
              shp.TextFrame.TextRange.Characters.Count - 1)
            If Len(sString) > 0 Then
                ' set the range to insert the text
                Set oRngAnchor = shp.Anchor.Paragraphs(1).Range
                ' insert the textbox text before the range object
                oRngAnchor.InsertBefore _
                  "Textbox start << " & sString & " >> Textbox end"
            End If
            shp.Delete
        End If
    Next shp
End Sub
