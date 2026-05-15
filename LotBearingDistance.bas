Attribute VB_Name = "LotBearingDistance"
' LotBearingDistance - for the FIRST selected closed shape, label every
' straight segment with surveyor bearing (N DD'MM"SS E/W) and length,
' plus an area label at the centroid.
'
' Key-in:   VBA RUN [ProjectName] LotBearingDistance.LotBDMain

Option Explicit

Private Const SQFT_PER_ACRE As Double = 43560#
Private Const PI As Double = 3.14159265358979

Public Sub LotBDMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oShape As ShapeElement
    Dim pts() As Point3d
    Dim nPts As Long
    Dim i As Long
    Dim p As Point3d, q As Point3d, mid As Point3d
    Dim mathAng As Double, az As Double, dist As Double
    Dim bear As String, lbl As String
    Dim ptLow As Point3d, ptHigh As Point3d, ptCenter As Point3d
    Dim oTxt As TextElement
    Dim areaSq As Double, areaAc As Double
    Dim rot As Matrix3d

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    If Not oEnum.MoveNext Then
        MsgBox "Select a closed shape first.", vbExclamation
        Exit Sub
    End If
    Set oEle = oEnum.Current
    If oEle.Type <> msdElementTypeShape Then
        MsgBox "First selected element must be a Shape.", vbExclamation
        Exit Sub
    End If
    Set oShape = oEle.AsShapeElement
    If Not oShape.IsClosedElement Then
        MsgBox "Shape must be closed.", vbExclamation
        Exit Sub
    End If
    pts = oShape.GetVertices()
    nPts = UBound(pts) - LBound(pts)

    For i = LBound(pts) To UBound(pts) - 1
        p = pts(i)
        q = pts(i + 1)
        mathAng = Atn2(q.Y - p.Y, q.X - p.X)
        az = (PI / 2#) - mathAng
        dist = Sqr((q.X - p.X) ^ 2 + (q.Y - p.Y) ^ 2)
        bear = AzimuthToBearing(az)
        mid.X = (p.X + q.X) / 2#
        mid.Y = (p.Y + q.Y) / 2#
        mid.Z = (p.Z + q.Z) / 2#
        lbl = bear & "  " & Format(dist, "0.00") & "'"

        Dim flip As Double
        flip = mathAng
        If flip > PI / 2# And flip < 3# * PI / 2# Then flip = flip - PI
        rot = Matrix3dFromVectorAndRotationAngle(Point3dFromXYZ(0, 0, 1), flip)
        Set oTxt = CreateTextElement1(Nothing, lbl, mid, rot)
        ActiveModelReference.AddElement oTxt
    Next i

    areaSq = oShape.Area
    areaAc = areaSq / SQFT_PER_ACRE
    ptLow = oEle.Range.Low
    ptHigh = oEle.Range.High
    ptCenter.X = (ptLow.X + ptHigh.X) / 2#
    ptCenter.Y = (ptLow.Y + ptHigh.Y) / 2#
    ptCenter.Z = (ptLow.Z + ptHigh.Z) / 2#
    Set oTxt = CreateTextElement1(Nothing, _
        Format(areaAc, "0.00") & " AC", ptCenter, Matrix3dIdentity)
    ActiveModelReference.AddElement oTxt

    Application.ShowStatus "Lot labeled: " & nPts & " segment(s)."
    Exit Sub
Cleanup:
    MsgBox "LotBearingDistance error: " & Err.Description, vbExclamation
End Sub

Private Function Atn2(ByVal dy As Double, ByVal dx As Double) As Double
    If dx > 0# Then
        Atn2 = Atn(dy / dx)
    ElseIf dx < 0# And dy >= 0# Then
        Atn2 = Atn(dy / dx) + PI
    ElseIf dx < 0# And dy < 0# Then
        Atn2 = Atn(dy / dx) - PI
    ElseIf dx = 0# And dy > 0# Then
        Atn2 = PI / 2#
    ElseIf dx = 0# And dy < 0# Then
        Atn2 = -PI / 2#
    Else
        Atn2 = 0#
    End If
End Function

Private Function AzimuthToBearing(ByVal az As Double) As String
    Dim pi2 As Double, theta As Double, ns As String, ew As String
    pi2 = 2# * PI
    Do While az < 0#: az = az + pi2: Loop
    Do While az >= pi2: az = az - pi2: Loop
    If az <= PI / 2# Then
        ns = "N": ew = "E": theta = az
    ElseIf az <= PI Then
        ns = "S": ew = "E": theta = PI - az
    ElseIf az <= 3# * PI / 2# Then
        ns = "S": ew = "W": theta = az - PI
    Else
        ns = "N": ew = "W": theta = pi2 - az
    End If
    AzimuthToBearing = ns & " " & RadToDMS(theta) & " " & ew
End Function

Private Function RadToDMS(ByVal rad As Double) As String
    Dim deg As Double, d As Long, m As Long, s As Long
    Dim mD As Double, sD As Double
    deg = rad * 180# / PI
    d = Int(deg)
    mD = (deg - d) * 60#
    m = Int(mD)
    sD = (mD - m) * 60#
    s = CLng(sD)
    If s >= 60 Then: s = 0: m = m + 1
    If m >= 60 Then: m = 0: d = d + 1
    RadToDMS = d & Chr$(176) & Format(m, "00") & "'" & Format(s, "00") & """"
End Function
