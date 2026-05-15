Attribute VB_Name = "PointsIO"
' PointsIO - PNEZD I/O.
'   ImportPNEZDMain   reads a PNEZD CSV (no header) and places a Line
'                     element of zero length at each row, plus a two-line
'                     text node showing PNO / Z / Description.
'   ExportPNEZDMain   writes a PNEZD CSV from EVERY text node element
'                     whose first line is a numeric point number.
'
' PNEZD = Point, Northing(Y), Easting(X), Z elevation, Description.
'
' Key-ins:
'   VBA RUN [ProjectName] PointsIO.ImportPNEZDMain
'   VBA RUN [ProjectName] PointsIO.ExportPNEZDMain

Option Explicit

Public Sub ImportPNEZDMain()
    Dim fname As String, fnum As Integer
    Dim line As String, parts() As String
    Dim pno As String, n As Double, e As Double, z As Double, desc As String
    Dim pt As Point3d
    Dim oLine As LineElement
    Dim oTN As TextNodeElement
    Dim count As Long

    fname = InputBox("Path to PNEZD .csv:", "ImportPNEZD")
    If Len(fname) = 0 Then Exit Sub

    On Error GoTo Cleanup
    fnum = FreeFile
    Open fname For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, line
        If Len(Trim$(line)) > 0 Then
            parts = Split(line, ",")
            If UBound(parts) >= 3 Then
                pno = Trim$(parts(0))
                n = CDbl(parts(1)): e = CDbl(parts(2)): z = CDbl(parts(3))
                desc = IIf(UBound(parts) >= 4, Trim$(parts(4)), "")
                pt.X = e: pt.Y = n: pt.Z = z
                Set oLine = CreateLineElement2(Nothing, pt, pt)
                ActiveModelReference.AddElement oLine

                Set oTN = CreateTextNodeElement1(Nothing, pt, Matrix3dIdentity)
                oTN.AddTextLine pno
                oTN.AddTextLine Format(z, "0.00")
                If Len(desc) > 0 Then oTN.AddTextLine desc
                ActiveModelReference.AddElement oTN
                count = count + 1
            End If
        End If
    Loop
    Close #fnum
    Application.ShowStatus count & " points imported from " & fname
    Exit Sub
Cleanup:
    If fnum > 0 Then Close #fnum
    MsgBox "ImportPNEZD error: " & Err.Description, vbExclamation
End Sub

Public Sub ExportPNEZDMain()
    Dim fname As String, fnum As Integer
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oTN As TextNodeElement
    Dim subE As ElementEnumerator
    Dim line1 As String, descLine As String, zVal As Double
    Dim pt As Point3d
    Dim count As Long

    fname = InputBox("Save PNEZD .csv to:", "ExportPNEZD")
    If Len(fname) = 0 Then Exit Sub

    On Error GoTo Cleanup
    fnum = FreeFile
    Open fname For Output As #fnum
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        If oEle.Type = msdElementTypeTextNode Then
            Set oTN = oEle.AsTextNodeElement
            pt = oTN.Origin
            Set subE = oTN.GetSubElements
            line1 = "": descLine = ""
            If subE.MoveNext Then line1 = subE.Current.AsTextElement.Text
            If subE.MoveNext Then  ' second line = Z
                zVal = CDbl(subE.Current.AsTextElement.Text)
            Else
                zVal = pt.Z
            End If
            If subE.MoveNext Then descLine = subE.Current.AsTextElement.Text
            If IsNumeric(line1) Then
                Print #fnum, line1 & "," & _
                    Format(pt.Y, "0.000") & "," & _
                    Format(pt.X, "0.000") & "," & _
                    Format(zVal, "0.000") & "," & descLine
                count = count + 1
            End If
        End If
    Loop
    Close #fnum
    Application.ShowStatus count & " points exported -> " & fname
    Exit Sub
Cleanup:
    If fnum > 0 Then Close #fnum
    MsgBox "ExportPNEZD error: " & Err.Description, vbExclamation
End Sub
