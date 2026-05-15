Attribute VB_Name = "CellCount"
' CellCount - tally placed cells in the model grouped by cell name. Match
' to Civil 3D's BCOUNT or AutoCAD's BLOCKCOUNT. Useful for survey QA
' (how many monument cells did you actually drop?).
'
' Key-in:   VBA RUN [ProjectName] CellCount.CellCountMain

Option Explicit

Public Sub CellCountMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oSC As ElementScanCriteria
    Dim dict As Object
    Dim nm As String, total As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set oSC = New ElementScanCriteria
    oSC.ExcludeAllTypes
    oSC.IncludeType msdElementTypeCellHeader
    Set oEnum = ActiveModelReference.Scan(oSC)
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        nm = oEle.AsCellElement.Name
        If dict.Exists(nm) Then
            dict(nm) = dict(nm) + 1
        Else
            dict.Add nm, 1
        End If
        total = total + 1
    Loop

    Dim keys As Variant, k As Variant, out As String
    keys = dict.keys
    out = "Cell                          Count" & vbCrLf & String(40, "-")
    For Each k In keys
        out = out & vbCrLf & Left$(k & String(30, " "), 30) & dict(k)
    Next k
    out = out & vbCrLf & String(40, "-") & vbCrLf & "Total: " & total
    MsgBox out, vbInformation, "Cell Counts"
End Sub
