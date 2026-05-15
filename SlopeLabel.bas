Attribute VB_Name = "SlopeLabel"
' SlopeLabel - take two selected line endpoints (first two selected line
' elements, or first selected line) and label the slope between them.
'
' If exactly ONE line is selected: uses its endpoints.
' If TWO line elements are selected: uses their start points.
' Reports percent slope and H:V ratio.
'
' Key-in:   VBA RUN [ProjectName] SlopeLabel.SlopeLabelMain

Option Explicit

Public Sub SlopeLabelMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim p1 As Point3d, p2 As Point3d, mid As Point3d
    Dim found As Long

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    found = 0
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        If oEle.Type = msdElementTypeLine Then
            If found = 0 Then
                p1 = oEle.AsLineElement.StartPoint
                p2 = oEle.AsLineElement.EndPoint
            ElseIf found = 1 Then
                p2 = oEle.AsLineElement.StartPoint
            End If
            found = found + 1
        End If
    Loop
    If found = 0 Then
        MsgBox "Select a line (or two lines).", vbExclamation
        Exit Sub
    End If

    Dim dx As Double, dy As Double, dz As Double, dh As Double
    Dim pct As Double, ratio As String
    dx = p2.X - p1.X: dy = p2.Y - p1.Y: dz = p2.Z - p1.Z
    dh = Sqr(dx * dx + dy * dy)
    If dh > 0# Then pct = 100# * dz / dh Else pct = 0#
    If dh > 0# And dz <> 0# Then
        ratio = Format(dh / Abs(dz), "0.00") & ":1"
    Else
        ratio = "flat"
    End If

    mid.X = (p1.X + p2.X) / 2#
    mid.Y = (p1.Y + p2.Y) / 2#
    mid.Z = (p1.Z + p2.Z) / 2#

    Dim oTxt As TextElement
    Set oTxt = CreateTextElement1(Nothing, _
        Format(pct, "0.00") & "% (" & ratio & ")", mid, Matrix3dIdentity)
    ActiveModelReference.AddElement oTxt

    Application.ShowStatus "Slope: " & Format(pct, "0.00") & _
        "%   dv=" & Format(dz, "0.000") & "   dh=" & Format(dh, "0.000")
    Exit Sub
Cleanup:
    MsgBox "SlopeLabel error: " & Err.Description, vbExclamation
End Sub
