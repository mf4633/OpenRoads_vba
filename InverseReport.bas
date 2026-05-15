Attribute VB_Name = "InverseReport"
' InverseReport - takes the FIRST selected line and reports a full COGO
' inverse: bearing, distance, dN, dE, dZ, slope, ratio.
'
' Key-in:   VBA RUN [ProjectName] InverseReport.InverseReportMain

Option Explicit

Private Const PI As Double = 3.14159265358979

Public Sub InverseReportMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim p1 As Point3d, p2 As Point3d

    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        If oEle.Type = msdElementTypeLine Then
            p1 = oEle.AsLineElement.StartPoint
            p2 = oEle.AsLineElement.EndPoint
            Exit Do
        End If
    Loop
    If IsZero(p2) Then
        MsgBox "Select a line.", vbExclamation
        Exit Sub
    End If

    Dim dx As Double, dy As Double, dz As Double, dh As Double, dist As Double
    dx = p2.X - p1.X: dy = p2.Y - p1.Y: dz = p2.Z - p1.Z
    dh = Sqr(dx * dx + dy * dy)
    dist = Sqr(dh * dh + dz * dz)

    Dim mathAng As Double, az As Double, bear As String
    mathAng = Atn2(dy, dx)
    az = (PI / 2#) - mathAng
    bear = AzimuthToBearing(az)

    Dim pct As Double, ratio As String
    If dh > 0# Then pct = 100# * dz / dh Else pct = 0#
    If dh > 0# And dz <> 0# Then
        ratio = Format(dh / Abs(dz), "0.00") & ":1"
    Else
        ratio = "flat"
    End If

    MsgBox _
        "Bearing : " & bear & vbCrLf & _
        "Distance: " & Format(dist, "0.000") & " ft (slope)" & vbCrLf & _
        "  horiz : " & Format(dh, "0.000") & vbCrLf & _
        "dN, dE  : " & Format(dy, "0.000") & "  " & Format(dx, "0.000") & vbCrLf & _
        "dZ      : " & Format(dz, "0.000") & vbCrLf & _
        "Slope   : " & Format(pct, "0.00") & "%  (" & ratio & ")", _
        vbInformation, "Inverse"
End Sub

Private Function IsZero(p As Point3d) As Boolean
    IsZero = (p.X = 0# And p.Y = 0# And p.Z = 0#)
End Function

Private Function Atn2(ByVal dy As Double, ByVal dx As Double) As Double
    If dx > 0#                  Then: Atn2 = Atn(dy / dx)         : Exit Function
    If dx < 0# And dy >= 0#     Then: Atn2 = Atn(dy / dx) + PI    : Exit Function
    If dx < 0# And dy < 0#      Then: Atn2 = Atn(dy / dx) - PI    : Exit Function
    If dx = 0# And dy > 0#      Then: Atn2 = PI / 2#              : Exit Function
    If dx = 0# And dy < 0#      Then: Atn2 = -PI / 2#             : Exit Function
    Atn2 = 0#
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
    Dim deg As Double, d As Long, m As Long, s As Long, mD As Double, sD As Double
    deg = rad * 180# / PI
    d = Int(deg): mD = (deg - d) * 60#: m = Int(mD): sD = (mD - m) * 60#: s = CLng(sD)
    If s >= 60 Then: s = 0: m = m + 1
    If m >= 60 Then: m = 0: d = d + 1
    RadToDMS = d & Chr$(176) & Format(m, "00") & "'" & Format(s, "00") & """"
End Function
