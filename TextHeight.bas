Attribute VB_Name = "TextHeight"
' TextHeight - prompt for a new height (master units), apply to every
' selected text / text node element.
'
' Key-in:   VBA RUN [ProjectName] TextHeight.TextHeightMain

Option Explicit

Public Sub TextHeightMain()
    Dim newHt As Double
    Dim s As String
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oTxt As TextElement
    Dim count As Long

    s = InputBox("New text height (master units):", "TextHeight")
    If Len(s) = 0 Then Exit Sub
    newHt = CDbl(s)
    If newHt <= 0# Then Exit Sub

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        Select Case oEle.Type
            Case msdElementTypeText
                Set oTxt = oEle.AsTextElement
                oTxt.TextStyle.Height = newHt
                oTxt.TextStyle.Width = newHt
                oTxt.Rewrite
                count = count + 1
            Case msdElementTypeTextNode
                Dim sub_ As ElementEnumerator
                Set sub_ = oEle.AsTextNodeElement.GetSubElements
                Do While sub_.MoveNext
                    Set oTxt = sub_.Current.AsTextElement
                    oTxt.TextStyle.Height = newHt
                    oTxt.TextStyle.Width = newHt
                    oTxt.Rewrite
                Loop
                count = count + 1
        End Select
    Loop
    Application.ShowStatus count & " text element(s) re-sized."
    Exit Sub
Cleanup:
    MsgBox "TextHeight error: " & Err.Description, vbExclamation
End Sub
