Attribute VB_Name = "LabelAcres"
' LabelAcres - label every selected closed shape with its area in acres
' to the tenth (e.g. "2.3 AC"). Text is placed at the shape's bounding-box
' center using the active text style.
'
' Run from VBA Project Manager, or with the key-in:
'   VBA RUN [ProjectName] LabelAcres.LabelAcresMain
'
' Assumes working units are survey feet. For metric DGNs change
' SQFT_PER_ACRE (1 acre = 4046.856 sq m).

Option Explicit

Private Const SQFT_PER_ACRE As Double = 43560#

Public Sub LabelAcresMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oText As TextElement
    Dim areaSq As Double
    Dim areaAc As Double
    Dim ptCenter As Point3d
    Dim ptLow As Point3d
    Dim ptHigh As Point3d
    Dim labelStr As String
    Dim count As Long

    On Error GoTo Cleanup

    Set oEnum = ActiveModelReference.GetSelectedElements()
    count = 0
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        areaSq = 0#
        Select Case oEle.Type
            Case msdElementTypeShape
                If oEle.AsShapeElement.IsClosedElement Then
                    areaSq = oEle.AsShapeElement.Area
                End If
            Case msdElementTypeComplexShape
                areaSq = oEle.AsComplexShapeElement.Area
            Case msdElementTypeEllipse
                areaSq = oEle.AsEllipseElement.Area
        End Select

        If areaSq > 0# Then
            ptLow = oEle.Range.Low
            ptHigh = oEle.Range.High
            ptCenter.X = (ptLow.X + ptHigh.X) / 2#
            ptCenter.Y = (ptLow.Y + ptHigh.Y) / 2#
            ptCenter.Z = (ptLow.Z + ptHigh.Z) / 2#

            areaAc = areaSq / SQFT_PER_ACRE
            labelStr = Format(areaAc, "0.0") & " AC"

            Set oText = CreateTextElement1(Nothing, labelStr, ptCenter, _
                                           Matrix3dIdentity)
            ActiveModelReference.AddElement oText
            count = count + 1
        End If
    Loop

    Application.ShowStatus count & " shape(s) labeled in acres."
    Exit Sub

Cleanup:
    MsgBox "LabelAcres error: " & Err.Description, vbExclamation
End Sub
