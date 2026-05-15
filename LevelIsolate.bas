Attribute VB_Name = "LevelIsolate"
' LevelIsolate - hide every level EXCEPT the level of the first selected
' element. LevelShowAllMain restores all levels.
'
' Key-ins:
'   VBA RUN [ProjectName] LevelIsolate.LevelIsolateMain
'   VBA RUN [ProjectName] LevelIsolate.LevelShowAllMain

Option Explicit

Public Sub LevelIsolateMain()
    Dim oEnum As ElementEnumerator
    Dim oEle As Element
    Dim keepID As Long
    Dim lvl As Level

    On Error GoTo Cleanup
    Set oEnum = ActiveModelReference.GetSelectedElements()
    If Not oEnum.MoveNext Then
        MsgBox "Select an element on the level to isolate.", vbExclamation
        Exit Sub
    End If
    Set oEle = oEnum.Current
    keepID = oEle.Level

    For Each lvl In ActiveDesignFile.Levels
        If lvl.ElementColor <> 0 Then  ' avoid breaking on lvl=0
            If lvl.LevelNumber <> ActiveDesignFile.Levels.Find(keepID).LevelNumber Then
                lvl.IsDisplayed = False
            End If
        End If
    Next lvl
    ActiveModelReference.Redraw
    Application.ShowStatus "Isolated level: " & _
        ActiveDesignFile.Levels.Find(keepID).Name
    Exit Sub
Cleanup:
    MsgBox "LevelIsolate error: " & Err.Description, vbExclamation
End Sub

Public Sub LevelShowAllMain()
    Dim lvl As Level
    For Each lvl In ActiveDesignFile.Levels
        lvl.IsDisplayed = True
    Next lvl
    ActiveModelReference.Redraw
    Application.ShowStatus "All levels displayed."
End Sub
