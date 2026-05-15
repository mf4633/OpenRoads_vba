Attribute VB_Name = "ZTools"
' ZTools - bulk Z manipulation of selected elements.
'   FlattenZMain  -> set Z = 0 on every vertex/origin of every selected element
'   MoveZMain     -> prompt for delta, add it to every Z
'   SetZMain      -> prompt for value, set every Z to it
'
' Works on Line, LineString, Shape, Arc, Text, TextNode, Cell elements.
' For meshes / B-spline surfaces / corridor elements, use the OpenRoads
' surface tools instead -- those geometries aren't safe to vertex-edit.
'
' Key-ins:
'   VBA RUN [ProjectName] ZTools.FlattenZMain
'   VBA RUN [ProjectName] ZTools.MoveZMain
'   VBA RUN [ProjectName] ZTools.SetZMain

Option Explicit

Public Sub FlattenZMain():  ApplyZ 1, 0#:           End Sub
Public Sub MoveZMain()
    Dim s As String, d As Double
    s = InputBox("Delta Z (master units):", "MoveZ")
    If Len(s) = 0 Then Exit Sub
    d = CDbl(s)
    ApplyZ 2, d
End Sub
Public Sub SetZMain()
    Dim s As String, d As Double
    s = InputBox("New Z value:", "SetZ")
    If Len(s) = 0 Then Exit Sub
    d = CDbl(s)
    ApplyZ 1, d
End Sub

Private Sub ApplyZ(ByVal mode As Long, ByVal v As Double)
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim count As Long

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        If TransformElementZ(oEle, mode, v) Then
            oEle.Rewrite
            count = count + 1
        End If
    Loop
    Application.ShowStatus count & " element(s) Z-updated."
    Exit Sub
Cleanup:
    MsgBox "ZTools error: " & Err.Description, vbExclamation
End Sub

Private Function TransformElementZ(oEle As Element, mode As Long, v As Double) As Boolean
    Dim pt As Point3d, pts() As Point3d, i As Long
    Select Case oEle.Type
        Case msdElementTypeLine
            With oEle.AsLineElement
                pt = .StartPoint: pt.Z = NewZ(pt.Z, mode, v): .StartPoint = pt
                pt = .EndPoint:   pt.Z = NewZ(pt.Z, mode, v): .EndPoint = pt
            End With
            TransformElementZ = True
        Case msdElementTypeLineString
            pts = oEle.AsLineStringElement.GetVertices
            For i = LBound(pts) To UBound(pts)
                pts(i).Z = NewZ(pts(i).Z, mode, v)
            Next i
            oEle.AsLineStringElement.SetVertices pts
            TransformElementZ = True
        Case msdElementTypeShape
            pts = oEle.AsShapeElement.GetVertices
            For i = LBound(pts) To UBound(pts)
                pts(i).Z = NewZ(pts(i).Z, mode, v)
            Next i
            oEle.AsShapeElement.SetVertices pts
            TransformElementZ = True
        Case msdElementTypeText
            With oEle.AsTextElement
                pt = .Origin: pt.Z = NewZ(pt.Z, mode, v): .Origin = pt
            End With
            TransformElementZ = True
        Case msdElementTypeCellHeader
            With oEle.AsCellElement
                pt = .Origin: pt.Z = NewZ(pt.Z, mode, v): .Origin = pt
            End With
            TransformElementZ = True
    End Select
End Function

Private Function NewZ(ByVal current As Double, ByVal mode As Long, _
                       ByVal v As Double) As Double
    Select Case mode
        Case 1: NewZ = v               ' set
        Case 2: NewZ = current + v     ' add
    End Select
End Function
