Attribute VB_Name = "TotalLength"
' TotalLength - sum the length of every selected line, line string, arc,
' curve, and complex string. Reports total ft and miles.
'
' Key-in:   VBA RUN [ProjectName] TotalLength.TotalLengthMain

Option Explicit

Public Sub TotalLengthMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim total As Double
    Dim count As Long

    On Error GoTo Cleanup
    total = 0#
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        Select Case oEle.Type
            Case msdElementTypeLine
                total = total + oEle.AsLineElement.Length
                count = count + 1
            Case msdElementTypeLineString
                total = total + oEle.AsLineStringElement.Length
                count = count + 1
            Case msdElementTypeArc
                total = total + oEle.AsArcElement.Length
                count = count + 1
            Case msdElementTypeBsplineCurve
                total = total + oEle.AsBsplineCurveElement.Length
                count = count + 1
            Case msdElementTypeComplexString
                total = total + oEle.AsComplexStringElement.Length
                count = count + 1
            Case msdElementTypeCurve
                total = total + oEle.AsCurveElement.Length
                count = count + 1
        End Select
    Loop

    MsgBox count & " linear element(s)" & vbCrLf & _
           Format(total, "#,##0.00") & " ft" & vbCrLf & _
           Format(total / 5280#, "0.000") & " mi", _
           vbInformation, "Total Length"
    Exit Sub
Cleanup:
    MsgBox "TotalLength error: " & Err.Description, vbExclamation
End Sub
