Attribute VB_Name = "TaskRules"
Option Explicit

Public Const FONT_NAME As String = "Meiryo UI"
Public Const FIRST_TASK_ROW As Long = 7
Public Const MAX_TITLE As Long = 120
Public Const MAX_TEXT As Long = 8000

Public Function StatusIndex(ByVal status As String) As Long
    Select Case status
        Case "起票済み": StatusIndex = 0
        Case "対応中": StatusIndex = 1
        Case "対応済み": StatusIndex = 2
        Case "連携済み": StatusIndex = 3
        Case "完了": StatusIndex = 4
        Case Else: StatusIndex = -1
    End Select
End Function

Public Function StatusColor(ByVal status As String) As Long
    Select Case StatusIndex(status)
        Case 0: StatusColor = RGB(255, 239, 166)
        Case 1: StatusColor = RGB(200, 230, 255)
        Case 2: StatusColor = RGB(210, 237, 214)
        Case 3: StatusColor = RGB(231, 218, 249)
        Case 4: StatusColor = RGB(226, 229, 233)
        Case Else: Err.Raise vbObjectError + 101, "StatusColor", "ステータスが不正です。"
    End Select
End Function

Public Function CleanTitle(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    value = Replace(value, ChrW(&H3000), " ")
    CleanTitle = Trim$(value)
End Function

Public Function ValidationMessage(ByVal title As String, ByVal body As String, _
                                  ByVal notes As String, ByVal status As String) As String
    If Len(CleanTitle(title)) = 0 Then
        ValidationMessage = "タイトルを入力してください。"
    ElseIf Len(CleanTitle(title)) > MAX_TITLE Then
        ValidationMessage = "タイトルは120文字以内にしてください。"
    ElseIf Len(body) > MAX_TEXT Or Len(notes) > MAX_TEXT Then
        ValidationMessage = "内容・備考はそれぞれ8000文字以内にしてください。"
    ElseIf StatusIndex(status) < 0 Then
        ValidationMessage = "ステータスを一覧から選択してください。"
    End If
End Function

Public Function PreviewText(ByVal value As String, ByVal limit As Long) As String
    Dim lastCode As Long
    value = Replace(Replace(value, vbCrLf, " / "), vbLf, " / ")
    value = Replace(value, vbCr, " / ")
    If Len(value) = 0 Then
        PreviewText = "（未記入）"
    ElseIf Len(value) <= limit Then
        PreviewText = value
    Else
        ' Do not cut a UTF-16 surrogate pair (for example, an emoji) in half.
        lastCode = CLng(AscW(Mid$(value, limit, 1))) And &HFFFF&
        If lastCode >= &HD800& And lastCode <= &HDBFF& Then limit = limit - 1
        PreviewText = Left$(value, limit) & "…"
    End If
End Function

Public Function CompletionValue(ByVal oldStatus As String, ByVal newStatus As String, _
                                ByVal oldCompleted As Variant, ByVal changedAt As Date) As Variant
    If newStatus <> "完了" Then
        CompletionValue = Empty
    ElseIf oldStatus = "完了" And IsDate(oldCompleted) Then
        CompletionValue = oldCompleted
    ElseIf oldStatus = "完了" And IsNumeric(oldCompleted) Then
        If CDbl(oldCompleted) > 0 Then
            CompletionValue = oldCompleted
        Else
            CompletionValue = changedAt
        End If
    Else
        CompletionValue = changedAt
    End If
End Function
