Attribute VB_Name = "PointGrid"
' PointGrid - prompts for two corners and a spacing, then drops a
' rectangular grid of point elements (zero-length lines) for use as a
' grading-volume checkpoint set or quick survey reference grid.
'
' Key-in:   VBA RUN [ProjectName] PointGrid.PointGridMain

Option Explicit

Public Sub PointGridMain()
    Dim p1 As Point3d, p2 As Point3d
    Dim dx As Double, dy As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim x As Double, y As Double
    Dim pt As Point3d
    Dim oLine As LineElement
    Dim count As Long

    On Error GoTo Cleanup
    p1 = PromptPoint("Pick first corner")
    p2 = PromptPoint("Pick opposite corner")
    dx = CDbl(InputBox("X spacing (master units):", "PointGrid", "10"))
    dy = CDbl(InputBox("Y spacing (master units):", "PointGrid", CStr(dx)))
    If dx <= 0# Or dy <= 0# Then Exit Sub

    x0 = MinD(p1.X, p2.X): x1 = MaxD(p1.X, p2.X)
    y0 = MinD(p1.Y, p2.Y): y1 = MaxD(p1.Y, p2.Y)
    y = y0
    Do While y <= y1 + 0.0001
        x = x0
        Do While x <= x1 + 0.0001
            pt.X = x: pt.Y = y: pt.Z = 0#
            Set oLine = CreateLineElement2(Nothing, pt, pt)
            ActiveModelReference.AddElement oLine
            count = count + 1
            x = x + dx
        Loop
        y = y + dy
    Loop
    Application.ShowStatus count & " grid points created."
    Exit Sub
Cleanup:
    MsgBox "PointGrid error: " & Err.Description, vbExclamation
End Sub

Private Function PromptPoint(ByVal msg As String) As Point3d
    Dim pts As Variant
    Dim s As String
    s = InputBox(msg & " as X,Y (master units):", "PointGrid")
    pts = Split(s, ",")
    PromptPoint.X = CDbl(pts(0))
    PromptPoint.Y = CDbl(pts(1))
    PromptPoint.Z = 0#
End Function

Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function
Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
