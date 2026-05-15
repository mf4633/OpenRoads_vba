Attribute VB_Name = "CoordLabel"
' CoordLabel - two subs for survey-style labeling at a picked location.
'   CoordNEMain    -> two-line N/E label  (Northing / Easting)
'   CoordElevMain  -> one-line Z label    ("EL ###.##")
'
' Each pulls X,Y,Z from the FIRST selected element's origin/range center,
' or prompts the user if no selection. Stack invocations to label many
' points: pick a survey monument, run, repeat.
'
' Key-ins:
'   VBA RUN [ProjectName] CoordLabel.CoordNEMain
'   VBA RUN [ProjectName] CoordLabel.CoordElevMain

Option Explicit

Public Sub CoordNEMain()
    Dim p As Point3d
    If Not GetTargetPoint(p) Then Exit Sub
    PlaceMultilineLabel p, _
        "N " & Format(p.Y, "0.00") & vbCrLf & _
        "E " & Format(p.X, "0.00")
End Sub

Public Sub CoordElevMain()
    Dim p As Point3d
    If Not GetTargetPoint(p) Then Exit Sub
    Dim oTxt As TextElement
    Set oTxt = CreateTextElement1(Nothing, _
        "EL " & Format(p.Z, "0.00"), p, Matrix3dIdentity)
    ActiveModelReference.AddElement oTxt
End Sub

Private Function GetTargetPoint(ByRef p As Point3d) As Boolean
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Set oEnum = ActiveModelReference.GetSelectedElements()
    If oEnum.MoveNext Then
        Set oEle = oEnum.Current
        p.X = (oEle.Range.Low.X + oEle.Range.High.X) / 2#
        p.Y = (oEle.Range.Low.Y + oEle.Range.High.Y) / 2#
        p.Z = (oEle.Range.Low.Z + oEle.Range.High.Z) / 2#
        GetTargetPoint = True
    Else
        Dim s As String, parts As Variant
        s = InputBox("Enter X,Y,Z (master units):", "CoordLabel")
        If Len(s) = 0 Then Exit Function
        parts = Split(s, ",")
        p.X = CDbl(parts(0)): p.Y = CDbl(parts(1))
        p.Z = IIf(UBound(parts) >= 2, CDbl(parts(2)), 0#)
        GetTargetPoint = True
    End If
End Function

Private Sub PlaceMultilineLabel(p As Point3d, ByVal txt As String)
    Dim oTNode As TextNodeElement
    Set oTNode = CreateTextNodeElement1(Nothing, p, Matrix3dIdentity)
    Dim lines() As String
    lines = Split(txt, vbCrLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        oTNode.AddTextLine lines(i)
    Next i
    ActiveModelReference.AddElement oTNode
End Sub
