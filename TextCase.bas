Attribute VB_Name = "TextCase"
' TextCase - three subs that change the case of every selected text
' or text node element.
'   TextUpperMain     -> UPPER
'   TextLowerMain     -> lower
'   TextTitleCaseMain -> Title Case
'
' Key-ins:
'   VBA RUN [ProjectName] TextCase.TextUpperMain
'   VBA RUN [ProjectName] TextCase.TextLowerMain
'   VBA RUN [ProjectName] TextCase.TextTitleCaseMain

Option Explicit

Public Sub TextUpperMain():     ApplyCase 1: End Sub
Public Sub TextLowerMain():     ApplyCase 2: End Sub
Public Sub TextTitleCaseMain(): ApplyCase 3: End Sub

Private Sub ApplyCase(ByVal mode As Long)
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oTxt As TextElement
    Dim oTNode As TextNodeElement
    Dim count As Long

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        Select Case oEle.Type
            Case msdElementTypeText
                Set oTxt = oEle.AsTextElement
                oTxt.Text = TransformCase(oTxt.Text, mode)
                oTxt.Rewrite
                count = count + 1
            Case msdElementTypeTextNode
                Set oTNode = oEle.AsTextNodeElement
                Dim lines As ElementEnumerator
                Set lines = oTNode.GetSubElements
                Do While lines.MoveNext
                    Set oTxt = lines.Current.AsTextElement
                    oTxt.Text = TransformCase(oTxt.Text, mode)
                    oTxt.Rewrite
                Loop
                count = count + 1
        End Select
    Loop
    Application.ShowStatus count & " text element(s) updated."
    Exit Sub
Cleanup:
    MsgBox "TextCase error: " & Err.Description, vbExclamation
End Sub

Private Function TransformCase(ByVal s As String, ByVal mode As Long) As String
    Select Case mode
        Case 1: TransformCase = UCase$(s)
        Case 2: TransformCase = LCase$(s)
        Case 3: TransformCase = TitleCase(s)
    End Select
End Function

Private Function TitleCase(ByVal s As String) As String
    Dim i As Long, ch As String, prev As String, out As String, cap As Boolean
    cap = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case " ", vbTab, "-", "_", ".", "/"
                out = out & ch
                cap = True
            Case Else
                If cap Then
                    out = out & UCase$(ch)
                    cap = False
                Else
                    out = out & LCase$(ch)
                End If
        End Select
    Next i
    TitleCase = out
End Function
