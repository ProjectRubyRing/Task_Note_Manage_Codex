Attribute VB_Name = "TaskBoard"
Option Explicit

Private mBusy As Boolean
Private Const CARD_PREFIX As String = "card_"
Private Const CARD_HEIGHT As Double = 252
Private Const CARD_GAP As Double = 18

Public Sub InitializeBoard()
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ProtectSheets
    BindButtons
    RenderBoards
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "表示の初期化に失敗しました。タスク一覧のデータは保持されています。" & _
           vbCrLf & Err.Description, vbExclamation, "付箋タスク管理"
End Sub

Public Sub ProtectSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFiltering:=True
    Next ws
End Sub

Private Sub BindButtons()
    Dim ws As Worksheet, item As Shape, action As String
    For Each ws In ThisWorkbook.Worksheets
        For Each item In ws.Shapes
            If Left$(item.Name, 4) = "btn_" Then
                action = Mid$(item.Name, 5)
                item.OnAction = MacroName(action)
            End If
        Next item
    Next ws
End Sub

Private Function MacroName(ByVal action As String) As String
    MacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
End Function

Public Function LastTaskRow() As Long
    LastTaskRow = TaskData.Cells(TaskData.Rows.Count, 1).End(xlUp).Row
    If LastTaskRow < FIRST_TASK_ROW Then LastTaskRow = FIRST_TASK_ROW - 1
End Function

Public Function FindTaskRow(ByVal taskId As String) As Long
    Dim rowNumber As Long
    For rowNumber = FIRST_TASK_ROW To LastTaskRow()
        If CStr(TaskData.Cells(rowNumber, 1).Value2) = taskId Then
            If FindTaskRow <> 0 Then
                Err.Raise vbObjectError + 102, "FindTaskRow", "タスクIDが重複しています: " & taskId
            End If
            FindTaskRow = rowNumber
        End If
    Next rowNumber
End Function

Private Function NextTaskId() As String
    Dim rowNumber As Long, taskId As String, largest As Long, number As Long
    For rowNumber = FIRST_TASK_ROW To LastTaskRow()
        taskId = CStr(TaskData.Cells(rowNumber, 1).Value2)
        If Not taskId Like "T######" Then
            Err.Raise vbObjectError + 103, "NextTaskId", "タスクIDが不正です: " & taskId
        End If
        number = CLng(Mid$(taskId, 2))
        If number > largest Then largest = number
    Next rowNumber
    If largest >= 999999 Then Err.Raise vbObjectError + 104, "NextTaskId", "タスクIDの上限に達しました。"
    NextTaskId = "T" & Format$(largest + 1, "000000")
End Function

Private Function ReadEditor(ByVal address As String) As String
    Dim cell As Range
    Set cell = TaskEditor.Range(address)
    If cell.HasFormula Or IsError(cell.Value2) Then
        Err.Raise vbObjectError + 105, "ReadEditor", "入力欄には数式ではなく文字を入力してください。"
    End If
    ReadEditor = CStr(cell.Value2)
End Function

Private Sub WriteText(ByVal cell As Range, ByVal value As String)
    ' Text formatting prevents =, +, -, @ prefixes from becoming formulas.
    cell.NumberFormat = "@"
    cell.Value2 = value
End Sub

Private Function DraftChanged() As Boolean
    Dim taskId As String, rowNumber As Long
    taskId = CStr(TaskEditor.Range("B4").Value2)
    If taskId = "新規" Then
        DraftChanged = Len(ReadEditor("B9") & ReadEditor("B14") & ReadEditor("B24")) > 0
        If ReadEditor("B6") <> "起票済み" Then DraftChanged = True
    Else
        rowNumber = FindTaskRow(taskId)
        If rowNumber = 0 Then
            DraftChanged = True
        Else
            DraftChanged = ReadEditor("B9") <> CStr(TaskData.Cells(rowNumber, 2).Value2) Or _
                           ReadEditor("B14") <> CStr(TaskData.Cells(rowNumber, 3).Value2) Or _
                           ReadEditor("B24") <> CStr(TaskData.Cells(rowNumber, 4).Value2) Or _
                           ReadEditor("B6") <> CStr(TaskData.Cells(rowNumber, 5).Value2)
        End If
    End If
End Function

Private Function CanReplaceDraft() As Boolean
    CanReplaceDraft = True
    If DraftChanged() Then
        CanReplaceDraft = MsgBox("未登録の入力内容を破棄して別の付箋を開きますか？" & _
                         vbCrLf & "残す場合は［いいえ］を選び、先に［登録・更新］を押してください。", _
                         vbYesNo + vbQuestion + vbDefaultButton2, "付箋タスク管理") = vbYes
    End If
End Function

Public Sub NewTask()
    On Error GoTo Failed
    If Not CanReplaceDraft() Then
        TaskEditor.Activate
        Exit Sub
    End If
    ProtectSheets
    WriteText TaskEditor.Range("B4"), "新規"
    WriteText TaskEditor.Range("B6"), "起票済み"
    TaskEditor.Range("B9:P11").ClearContents
    TaskEditor.Range("B14:P21").ClearContents
    TaskEditor.Range("B24:P31").ClearContents
    TaskEditor.Activate
    TaskEditor.Range("B9").Select
    Exit Sub
Failed:
    MsgBox "新しい付箋を開けませんでした。" & vbCrLf & Err.Description, vbExclamation, "付箋タスク管理"
End Sub

Public Sub OpenTask(ByVal taskId As String)
    Dim rowNumber As Long
    On Error GoTo Failed
    rowNumber = FindTaskRow(taskId)
    If rowNumber = 0 Then Err.Raise vbObjectError + 106, "OpenTask", "対象のタスクが見つかりません。"
    If Not CanReplaceDraft() Then
        TaskEditor.Activate
        Exit Sub
    End If
    ProtectSheets
    WriteText TaskEditor.Range("B4"), taskId
    WriteText TaskEditor.Range("B9"), CStr(TaskData.Cells(rowNumber, 2).Value2)
    WriteText TaskEditor.Range("B14"), CStr(TaskData.Cells(rowNumber, 3).Value2)
    WriteText TaskEditor.Range("B24"), CStr(TaskData.Cells(rowNumber, 4).Value2)
    WriteText TaskEditor.Range("B6"), CStr(TaskData.Cells(rowNumber, 5).Value2)
    TaskEditor.Activate
    TaskEditor.Range("B6").Select
    Exit Sub
Failed:
    MsgBox "付箋を開けませんでした。" & vbCrLf & Err.Description, vbExclamation, "付箋タスク管理"
End Sub

Public Sub EditCard()
    Dim callerName As Variant
    callerName = Application.Caller
    If VarType(callerName) <> vbString Then Exit Sub
    If Left$(CStr(callerName), Len(CARD_PREFIX)) <> CARD_PREFIX Then Exit Sub
    OpenTask Mid$(CStr(callerName), Len(CARD_PREFIX) + 1)
End Sub

Public Function StoreTask(ByVal taskId As String, ByVal title As String, ByVal body As String, _
                          ByVal notes As String, ByVal status As String) As String
    Dim rowNumber As Long, message As String, previousStatus As String
    Dim createdAt As Variant, completedAt As Variant, changedAt As Date
    Dim record(1 To 1, 1 To 8) As Variant
    message = ValidationMessage(title, body, notes, status)
    If Len(message) > 0 Then Err.Raise vbObjectError + 107, "StoreTask", message
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 108, "StoreTask", "読み取り専用です。編集できる場所へ名前を付けて保存してください。"
    changedAt = Now
    If taskId = "新規" Then
        taskId = NextTaskId()
        rowNumber = LastTaskRow() + 1
        createdAt = changedAt
    Else
        rowNumber = FindTaskRow(taskId)
        If rowNumber = 0 Then Err.Raise vbObjectError + 109, "StoreTask", "更新対象のタスクがありません。"
        previousStatus = CStr(TaskData.Cells(rowNumber, 5).Value2)
        createdAt = TaskData.Cells(rowNumber, 6).Value
        completedAt = TaskData.Cells(rowNumber, 8).Value
    End If
    record(1, 1) = taskId
    record(1, 2) = CleanTitle(title)
    record(1, 3) = body
    record(1, 4) = notes
    record(1, 5) = status
    record(1, 6) = createdAt
    record(1, 7) = changedAt
    record(1, 8) = CompletionValue(previousStatus, status, completedAt, changedAt)
    TaskData.Cells(rowNumber, 1).Resize(1, 5).NumberFormat = "@"
    TaskData.Cells(rowNumber, 6).Resize(1, 3).NumberFormat = "yyyy/mm/dd hh:mm"
    ' Commit the whole record once, after validation; never delete archival data.
    TaskData.Cells(rowNumber, 1).Resize(1, 8).Value = record
    StoreTask = taskId
End Function

Public Sub SaveTask()
    Dim oldEvents As Boolean, oldScreen As Boolean, dataWritten As Boolean
    Dim taskId As String, status As String, message As String
    If mBusy Then Exit Sub
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ProtectSheets
    status = ReadEditor("B6")
    taskId = StoreTask(ReadEditor("B4"), ReadEditor("B9"), ReadEditor("B14"), ReadEditor("B24"), status)
    dataWritten = True
    WriteText TaskEditor.Range("B4"), taskId
    WriteText TaskEditor.Range("B9"), CleanTitle(ReadEditor("B9"))
    RenderBoards
    If status = "完了" Then
        CompletedBoard.Activate
    Else
        ActiveBoard.Activate
    End If
    ThisWorkbook.Save
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    mBusy = False
    Exit Sub
Failed:
    message = Err.Description
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    mBusy = False
    If dataWritten Then
        MsgBox "ブック内のタスクは更新しましたが、表示またはファイル保存に失敗しました。" & _
               vbCrLf & "ブックを閉じずに［再描画］を押し、Ctrl+Sで保存してください。" & _
               vbCrLf & message, vbExclamation, "付箋タスク管理"
    Else
        MsgBox "登録できませんでした。入力内容は残っています。" & vbCrLf & message, vbExclamation, "付箋タスク管理"
    End If
End Sub

Public Sub RefreshBoard()
    InitializeBoard
End Sub

Public Sub ShowBoard()
    ActiveBoard.Activate
End Sub

Public Sub ShowCompleted()
    CompletedBoard.Activate
End Sub

Public Sub ShowEditor()
    TaskEditor.Activate
End Sub

Private Sub ClearCards(ByVal ws As Worksheet)
    Dim index As Long
    For index = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(index).Name, Len(CARD_PREFIX)) = CARD_PREFIX Then ws.Shapes(index).Delete
    Next index
End Sub

Public Sub RenderBoards()
    Dim rowNumber As Long, lastRow As Long, lane As Long, completedCount As Long
    Dim slots(0 To 3) As Long, status As String
    Dim taskId As String, seen As Collection
    lastRow = LastTaskRow()
    Set seen = New Collection
    ' Validate the source before replacing the derived display.
    For rowNumber = FIRST_TASK_ROW To lastRow
        taskId = CStr(TaskData.Cells(rowNumber, 1).Value2)
        If Not taskId Like "T######" Then Err.Raise vbObjectError + 110, "RenderBoards", "タスクIDが不正です。"
        seen.Add taskId, taskId
        status = CStr(TaskData.Cells(rowNumber, 5).Value2)
        If StatusIndex(status) < 0 Then Err.Raise vbObjectError + 111, "RenderBoards", "ステータスが不正です: " & taskId
    Next rowNumber
    ClearCards ActiveBoard
    ClearCards CompletedBoard
    For rowNumber = FIRST_TASK_ROW To lastRow
        status = CStr(TaskData.Cells(rowNumber, 5).Value2)
        lane = StatusIndex(status)
        If lane = 4 Then
            AddCard CompletedBoard, rowNumber, completedCount Mod 4, completedCount \ 4
            completedCount = completedCount + 1
        Else
            AddCard ActiveBoard, rowNumber, lane, slots(lane)
            slots(lane) = slots(lane) + 1
        End If
    Next rowNumber
    If lastRow < FIRST_TASK_ROW Then
        ActiveBoard.Range("B8").Value2 = "［＋ 新しい付箋］から、最初のタスクを登録しましょう。"
    Else
        ActiveBoard.Range("B8").Value2 = "付箋をクリックして編集。ステータスを変更し［登録・更新］を押すと、色と配置が変わります。"
    End If
    If completedCount = 0 Then
        CompletedBoard.Range("B8").Value2 = "完了した付箋はここに移動します。内容・備考はそのまま残ります。"
    Else
        CompletedBoard.Range("B8").Value2 = "完了した付箋もクリックして編集できます。ステータスを戻すとタスクボードへ復帰します。"
    End If
    FormatData lastRow
    ActiveBoard.Calculate
    CompletedBoard.Calculate
End Sub

Private Sub FormatData(ByVal lastRow As Long)
    Dim savedNumber As Long, savedDescription As String
    On Error GoTo Failed
    TaskData.Unprotect
    If lastRow < FIRST_TASK_ROW Then lastRow = FIRST_TASK_ROW
    With TaskData.Range("A7:H" & CStr(lastRow))
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 64
    End With
    If TaskData.AutoFilterMode Then TaskData.AutoFilterMode = False
    TaskData.Range("A6:H" & CStr(lastRow)).AutoFilter
    TaskData.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
Failed:
    savedNumber = Err.Number
    savedDescription = Err.Description
    TaskData.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     UserInterfaceOnly:=True, AllowFiltering:=True
    Err.Raise savedNumber, "FormatData", savedDescription
End Sub

Private Sub AddCard(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal lane As Long, ByVal slot As Long)
    Dim item As Shape, taskId As String, status As String, title As String, text As String
    Dim left As Double, top As Double, width As Double, column As Long
    taskId = CStr(TaskData.Cells(rowNumber, 1).Value2)
    status = CStr(TaskData.Cells(rowNumber, 5).Value2)
    title = PreviewText(CStr(TaskData.Cells(rowNumber, 2).Value2), 40)
    text = status & "  /  " & taskId & vbLf & title & vbLf & vbLf & _
           "内容" & vbLf & PreviewText(CStr(TaskData.Cells(rowNumber, 3).Value2), 96) & vbLf & vbLf & _
           "備考・関連する実装" & vbLf & PreviewText(CStr(TaskData.Cells(rowNumber, 4).Value2), 62) & _
           vbLf & vbLf & "クリックして全文を表示・編集"
    column = 2 + 14 * lane
    left = ws.Cells(11, column).Left
    top = ws.Cells(11, column).Top + slot * (CARD_HEIGHT + CARD_GAP)
    width = ws.Range(ws.Cells(11, column), ws.Cells(11, column + 12)).Width
    Set item = ws.Shapes.AddShape(1, left, top, width, CARD_HEIGHT)
    item.Name = CARD_PREFIX & taskId
    item.AlternativeText = taskId & " " & status & " " & CStr(TaskData.Cells(rowNumber, 2).Value2)
    item.Placement = xlFreeFloating
    item.Locked = True
    item.OnAction = MacroName("EditCard")
    item.Fill.ForeColor.RGB = StatusColor(status)
    item.Line.Visible = False
    With item.TextFrame2
        .MarginLeft = 14
        .MarginRight = 14
        .MarginTop = 12
        .MarginBottom = 12
        .WordWrap = True
        .AutoSize = 2
        .TextRange.Text = text
        .TextRange.Font.Name = FONT_NAME
        .TextRange.Font.NameFarEast = FONT_NAME
        .TextRange.Font.Size = 11
        If status = "完了" Then
            .TextRange.Font.Fill.ForeColor.RGB = RGB(104, 112, 122)
        Else
            .TextRange.Font.Fill.ForeColor.RGB = RGB(36, 49, 64)
        End If
        .TextRange.Characters(Len(status & "  /  " & taskId) + 2, Len(title)).Font.Bold = True
        .TextRange.Characters(Len(status & "  /  " & taskId) + 2, Len(title)).Font.Size = 15
    End With
End Sub
