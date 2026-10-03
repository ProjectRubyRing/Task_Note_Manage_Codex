Option Explicit

Private Sub Workbook_Open()
    TaskBoard.InitializeBoard
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh Is TaskData Then
        If Target.Row >= FIRST_TASK_ROW And Target.Column <= 8 Then
            If Len(CStr(TaskData.Cells(Target.Row, 1).Value2)) > 0 Then
                Cancel = True
                TaskBoard.OpenTask CStr(TaskData.Cells(Target.Row, 1).Value2)
            End If
        End If
    End If
End Sub
