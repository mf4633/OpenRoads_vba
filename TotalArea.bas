Attribute VB_Name = "TotalArea"
' TotalArea - sum the area of every selected closed shape, ellipse, and
' complex shape; report the total in sq ft and acres.
'
' Key-in:   VBA RUN [ProjectName] TotalArea.TotalAreaMain

Option Explicit

Private Const SQFT_PER_ACRE As Double = 43560#

Public Sub TotalAreaMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim total As Double
    Dim count As Long
    Dim areaSq As Double

    On Error GoTo Cleanup
    total = 0#
    count = 0
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        areaSq = 0#
        Select Case oEle.Type
            Case msdElementTypeShape
                If oEle.AsShapeElement.IsClosedElement Then _
                    areaSq = oEle.AsShapeElement.Area
            Case msdElementTypeComplexShape
                areaSq = oEle.AsComplexShapeElement.Area
            Case msdElementTypeEllipse
                areaSq = oEle.AsEllipseElement.Area
        End Select
        If areaSq > 0# Then
            total = total + areaSq
            count = count + 1
        End If
    Loop

    MsgBox count & " shape(s)" & vbCrLf & _
           Format(total, "#,##0.00") & " sq ft" & vbCrLf & _
           Format(total / SQFT_PER_ACRE, "0.000") & " acres", _
           vbInformation, "Total Area"
    Exit Sub
Cleanup:
    MsgBox "TotalArea error: " & Err.Description, vbExclamation
End Sub
