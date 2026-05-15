Attribute VB_Name = "StationOffset"
' StationOffset - given a selected alignment line/linestring AND a target
' point (entered as X,Y), compute and label the station along the
' alignment and perpendicular offset (L/R) at that point.
'
' Conservative implementation that works on plain MicroStation geometry,
' so it functions even when the OpenRoads CivilModel API is not loaded.
' For full OpenRoads alignment objects, prefer the native annotation tools.
'
' Key-in:   VBA RUN [ProjectName] StationOffset.StationOffsetMain

Option Explicit

Private Const PI As Double = 3.14159265358979

Public Sub StationOffsetMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim segs() As Point3d
    Dim has As Boolean

    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        Select Case oEle.Type
            Case msdElementTypeLine
                ReDim segs(1)
                segs(0) = oEle.AsLineElement.StartPoint
                segs(1) = oEle.AsLineElement.EndPoint
                has = True: Exit Do
            Case msdElementTypeLineString
                segs = oEle.AsLineStringElement.GetVertices
                has = True: Exit Do
            Case msdElementTypeShape
                segs = oEle.AsShapeElement.GetVertices
                has = True: Exit Do
        End Select
    Loop
    If Not has Then
        MsgBox "Select an alignment line / linestring / shape.", vbExclamation
        Exit Sub
    End If

    Dim s As String, parts As Variant
    Dim p As Point3d
    s = InputBox("Target point X,Y (master units):", "StationOffset")
    If Len(s) = 0 Then Exit Sub
    parts = Split(s, ",")
    p.X = CDbl(parts(0)): p.Y = CDbl(parts(1)): p.Z = 0#

    Dim bestSta As Double, bestOff As Double, bestSide As String
    Dim cum As Double, i As Long
    Dim a As Point3d, b As Point3d, foot As Point3d
    Dim segLen As Double, t As Double, dxs As Double, dys As Double
    Dim cross As Double, foundIt As Boolean

    cum = 0#: bestOff = 1E+99
    For i = LBound(segs) To UBound(segs) - 1
        a = segs(i): b = segs(i + 1)
        dxs = b.X - a.X: dys = b.Y - a.Y
        segLen = Sqr(dxs * dxs + dys * dys)
        If segLen > 0# Then
            t = ((p.X - a.X) * dxs + (p.Y - a.Y) * dys) / (segLen * segLen)
            If t < 0# Then t = 0#
            If t > 1# Then t = 1#
            foot.X = a.X + t * dxs
            foot.Y = a.Y + t * dys
            Dim dd As Double
            dd = Sqr((p.X - foot.X) ^ 2 + (p.Y - foot.Y) ^ 2)
            If dd < bestOff Then
                bestOff = dd
                bestSta = cum + t * segLen
                cross = dxs * (p.Y - a.Y) - dys * (p.X - a.X)
                If cross >= 0# Then bestSide = "L" Else bestSide = "R"
                foundIt = True
            End If
        End If
        cum = cum + segLen
    Next i

    If Not foundIt Then Exit Sub
    Dim lbl As String
    lbl = "STA " & Format(bestSta, "0.00") & vbCrLf & _
          "OFF " & Format(bestOff, "0.00") & " " & bestSide

    Dim oTN As TextNodeElement
    Set oTN = CreateTextNodeElement1(Nothing, p, Matrix3dIdentity)
    oTN.AddTextLine "STA " & Format(bestSta, "0.00")
    oTN.AddTextLine "OFF " & Format(bestOff, "0.00") & " " & bestSide
    ActiveModelReference.AddElement oTN

    Application.ShowStatus lbl
End Sub
