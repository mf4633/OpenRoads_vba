Attribute VB_Name = "LevelCount"
' LevelCount - count every placed element grouped by level. Output goes
' to a message box (and the status bar shows total).
'
' Key-in:   VBA RUN [ProjectName] LevelCount.LevelCountMain

Option Explicit

Public Sub LevelCountMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oSC As ElementScanCriteria
    Dim dict As Object
    Dim levelName As String
    Dim total As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set oSC = New ElementScanCriteria
    oSC.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oSC)
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        levelName = ActiveDesignFile.Levels.Find(oEle.Level).Name
        If dict.Exists(levelName) Then
            dict(levelName) = dict(levelName) + 1
        Else
            dict.Add levelName, 1
        End If
        total = total + 1
    Loop

    Dim keys As Variant, k As Variant, out As String
    keys = dict.keys
    out = "Level                       Count" & vbCrLf & String(40, "-")
    For Each k In keys
        out = out & vbCrLf & Left$(k & String(28, " "), 28) & dict(k)
    Next k
    out = out & vbCrLf & String(40, "-") & vbCrLf & "Total: " & total
    MsgBox out, vbInformation, "Level Counts"
End Sub
