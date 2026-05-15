Attribute VB_Name = "LevelTranslate"
' LevelTranslate - move elements from one level to another in bulk,
' driven by a CSV file with rows of  OLD_LEVEL,NEW_LEVEL .
' New levels are created if missing.
'
' Key-in:   VBA RUN [ProjectName] LevelTranslate.LevelTranslateMain

Option Explicit

Public Sub LevelTranslateMain()
    Dim fname As String
    fname = InputBox("Path to CSV (OLD,NEW per line):", "LevelTranslate")
    If Len(fname) = 0 Then Exit Sub

    Dim fnum As Integer, line As String, parts() As String
    Dim oldName As String, newName As String
    Dim total As Long

    On Error GoTo Cleanup
    fnum = FreeFile
    Open fname For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, line
        If Len(Trim$(line)) > 0 Then
            parts = Split(line, ",")
            If UBound(parts) >= 1 Then
                oldName = Trim$(parts(0))
                newName = Trim$(parts(1))
                EnsureLevel newName
                total = total + RemapLevel(oldName, newName)
            End If
        End If
    Loop
    Close #fnum
    Application.ShowStatus total & " element(s) re-leveled."
    Exit Sub
Cleanup:
    If fnum > 0 Then Close #fnum
    MsgBox "LevelTranslate error: " & Err.Description, vbExclamation
End Sub

Private Sub EnsureLevel(ByVal name As String)
    Dim lvl As Level
    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels.Find(name)
    If lvl Is Nothing Then ActiveDesignFile.Levels.Add name
    On Error GoTo 0
End Sub

Private Function RemapLevel(ByVal oldName As String, _
                            ByVal newName As String) As Long
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim oSC As ElementScanCriteria
    Dim newLvl As Level
    Dim count As Long

    Set newLvl = ActiveDesignFile.Levels.Find(newName)
    If newLvl Is Nothing Then Exit Function

    Set oSC = New ElementScanCriteria
    oSC.IncludeLevel ActiveDesignFile.Levels.Find(oldName)
    Set oEnum = ActiveModelReference.Scan(oSC)
    Do While oEnum.MoveNext
        Set oEle = oEnum.Current
        oEle.Level = newLvl
        oEle.Rewrite
        count = count + 1
    Loop
    RemapLevel = count
End Function
