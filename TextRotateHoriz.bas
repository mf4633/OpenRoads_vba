Attribute VB_Name = "TextRotateHoriz"
' TextRotateHoriz - flip upside-down text by 180 degrees so every selected
' text element reads left-to-right. Leaves already-readable text alone.
'
' Key-in:   VBA RUN [ProjectName] TextRotateHoriz.TextRotateHorizMain

Option Explicit

Private Const PI As Double = 3.14159265358979

Public Sub TextRotateHorizMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oTxt As TextElement
    Dim ang As Double
    Dim rot As Matrix3d
    Dim count As Long

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        If oEle.Type = msdElementTypeText Then
            Set oTxt = oEle.AsTextElement
            ang = Matrix3dGetRotationAroundZ(oTxt.Rotation)
            Do While ang < 0#:        ang = ang + 2# * PI: Loop
            Do While ang >= 2# * PI:  ang = ang - 2# * PI: Loop
            If ang > PI / 2# And ang < 3# * PI / 2# Then
                ang = ang - PI
                rot = Matrix3dFromVectorAndRotationAngle( _
                       Point3dFromXYZ(0, 0, 1), ang)
                oTxt.Rotation = rot
                oTxt.Rewrite
                count = count + 1
            End If
        End If
    Loop
    Application.ShowStatus count & " text element(s) flipped readable."
    Exit Sub
Cleanup:
    MsgBox "TextRotateHoriz error: " & Err.Description, vbExclamation
End Sub

Private Function Matrix3dGetRotationAroundZ(m As Matrix3d) As Double
    ' Matrix3d in VBA exposes RowX/RowY/RowZ as Point3d. Atan2(y, x).
    Dim rx As Point3d
    rx = m.RowX
    Matrix3dGetRotationAroundZ = Atn2(rx.Y, rx.X)
End Function

Private Function Atn2(ByVal dy As Double, ByVal dx As Double) As Double
    If dx > 0#                  Then: Atn2 = Atn(dy / dx)         : Exit Function
    If dx < 0# And dy >= 0#     Then: Atn2 = Atn(dy / dx) + PI    : Exit Function
    If dx < 0# And dy < 0#      Then: Atn2 = Atn(dy / dx) - PI    : Exit Function
    If dx = 0# And dy > 0#      Then: Atn2 = PI / 2#              : Exit Function
    If dx = 0# And dy < 0#      Then: Atn2 = -PI / 2#             : Exit Function
    Atn2 = 0#
End Function
