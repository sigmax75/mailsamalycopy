Attribute VB_Name = "MailSamalyCopy"
Option Explicit

' ============================================================
' MailSamalyCopy - Outlook選択メール情報をExcelに一括出力
' ============================================================
' 使い方:
'   1. Outlookでメールを複数選択
'   2. Alt+F11 -> このモジュールを貼り付け
'   3. ExportSelectedMails を実行（またはリボンに登録）
' ============================================================

Private Const ITEM_COUNT As Long = 15

' 項目定義
Private Function GetFieldName(ByVal idx As Long) As String
    Select Case idx
        Case 1:  GetFieldName = "件名"
        Case 2:  GetFieldName = "送信者名"
        Case 3:  GetFieldName = "送信者メールアドレス"
        Case 4:  GetFieldName = "受信日時"
        Case 5:  GetFieldName = "送信日時"
        Case 6:  GetFieldName = "宛先"
        Case 7:  GetFieldName = "CC"
        Case 8:  GetFieldName = "BCC"
        Case 9:  GetFieldName = "本文（テキスト）"
        Case 10: GetFieldName = "本文（HTML）"
        Case 11: GetFieldName = "添付ファイル名"
        Case 12: GetFieldName = "重要度"
        Case 13: GetFieldName = "カテゴリ"
        Case 14: GetFieldName = "メッセージサイズ"
        Case 15: GetFieldName = "会話トピック"
        Case Else: GetFieldName = ""
    End Select
End Function

' メールオブジェクトから指定項目の値を取得
Private Function GetFieldValue(ByVal oMail As Object, ByVal idx As Long) As String
    Dim i As Long
    Dim sAttach As String

    On Error Resume Next

    Select Case idx
        Case 1
            GetFieldValue = oMail.Subject
        Case 2
            GetFieldValue = oMail.SenderName
        Case 3
            ' ExchangeユーザーはSMTPアドレスを別途取得
            If oMail.SenderEmailType = "EX" Then
                Dim oSender As Object
                Set oSender = oMail.Sender
                If Not oSender Is Nothing Then
                    GetFieldValue = oSender.GetExchangeUser().PrimarySmtpAddress
                    If Err.Number <> 0 Then
                        Err.Clear
                        GetFieldValue = oMail.SenderEmailAddress
                    End If
                Else
                    GetFieldValue = oMail.SenderEmailAddress
                End If
                Set oSender = Nothing
            Else
                GetFieldValue = oMail.SenderEmailAddress
            End If
        Case 4
            If IsDate(oMail.ReceivedTime) Then
                GetFieldValue = Format(oMail.ReceivedTime, "yyyy/mm/dd hh:nn:ss")
            End If
        Case 5
            If IsDate(oMail.SentOn) Then
                GetFieldValue = Format(oMail.SentOn, "yyyy/mm/dd hh:nn:ss")
            End If
        Case 6
            GetFieldValue = oMail.To
        Case 7
            GetFieldValue = oMail.CC
        Case 8
            GetFieldValue = oMail.BCC
        Case 9
            GetFieldValue = oMail.Body
        Case 10
            GetFieldValue = oMail.HTMLBody
        Case 11
            sAttach = ""
            If oMail.Attachments.Count > 0 Then
                For i = 1 To oMail.Attachments.Count
                    If i > 1 Then sAttach = sAttach & ", "
                    sAttach = sAttach & oMail.Attachments.Item(i).FileName
                Next i
            End If
            GetFieldValue = sAttach
        Case 12
            Select Case oMail.Importance
                Case 0: GetFieldValue = "低"
                Case 1: GetFieldValue = "標準"
                Case 2: GetFieldValue = "高"
                Case Else: GetFieldValue = CStr(oMail.Importance)
            End Select
        Case 13
            GetFieldValue = oMail.Categories
        Case 14
            GetFieldValue = CStr(oMail.Size)
        Case 15
            GetFieldValue = oMail.ConversationTopic
        Case Else
            GetFieldValue = ""
    End Select

    On Error GoTo 0
End Function

' ============================================================
' メインプロシージャ
' ============================================================
Public Sub ExportSelectedMails()

    Dim olApp As Object
    Dim olSel As Object
    Dim oMail As Object
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    Dim selectedFields() As Long
    Dim fieldCount As Long
    Dim sInput As String
    Dim sPrompt As String
    Dim arrTokens() As String
    Dim i As Long, j As Long
    Dim nVal As Long
    Dim colIdx As Long
    Dim rowIdx As Long

    ' --- 選択メールの取得 ---
    Set olApp = Application
    Set olSel = olApp.ActiveExplorer.Selection

    If olSel.Count = 0 Then
        MsgBox "メールを選択してください。" & vbCrLf & _
               "受信トレイ等でメールを1つ以上選択してから実行してください。", _
               vbExclamation, "MailSamalyCopy"
        GoTo CleanUp
    End If

    ' --- 項目選択（InputBox） ---
    sPrompt = "取得する項目の番号をカンマ区切りで入力してください。" & vbCrLf & vbCrLf
    For i = 1 To ITEM_COUNT
        sPrompt = sPrompt & Format(i, "00") & ". " & GetFieldName(i) & vbCrLf
    Next i
    sPrompt = sPrompt & vbCrLf
    sPrompt = sPrompt & "例: 1,2,4,6" & vbCrLf
    sPrompt = sPrompt & "ALL: 全項目を取得" & vbCrLf & vbCrLf
    sPrompt = sPrompt & "選択メール数: " & olSel.Count & " 件"

    sInput = InputBox(sPrompt, "MailSamalyCopy - 項目選択", "ALL")

    ' キャンセル押下（空文字が返る）
    If Len(Trim(sInput)) = 0 Then
        GoTo CleanUp
    End If

    ' --- 入力解析 ---
    sInput = Trim(UCase(sInput))

    If sInput = "ALL" Then
        ' 全項目
        fieldCount = ITEM_COUNT
        ReDim selectedFields(1 To fieldCount)
        For i = 1 To ITEM_COUNT
            selectedFields(i) = i
        Next i
    Else
        ' カンマ区切り解析
        sInput = Replace(sInput, " ", "")
        arrTokens = Split(sInput, ",")

        ' まず有効な値をカウント
        fieldCount = 0
        For i = LBound(arrTokens) To UBound(arrTokens)
            If IsNumeric(arrTokens(i)) Then
                nVal = CLng(arrTokens(i))
                If nVal >= 1 And nVal <= ITEM_COUNT Then
                    fieldCount = fieldCount + 1
                End If
            End If
        Next i

        If fieldCount = 0 Then
            MsgBox "項目を1つ以上選択してください。" & vbCrLf & _
                   "1～" & ITEM_COUNT & " の番号をカンマ区切りで入力してください。", _
                   vbExclamation, "MailSamalyCopy"
            GoTo CleanUp
        End If

        ' 配列に格納（重複除去付き）
        ReDim selectedFields(1 To fieldCount)
        Dim usedFlags(1 To 15) As Boolean
        j = 0
        For i = LBound(arrTokens) To UBound(arrTokens)
            If IsNumeric(arrTokens(i)) Then
                nVal = CLng(arrTokens(i))
                If nVal >= 1 And nVal <= ITEM_COUNT Then
                    If Not usedFlags(nVal) Then
                        usedFlags(nVal) = True
                        j = j + 1
                        selectedFields(j) = nVal
                    End If
                End If
            End If
        Next i

        ' 重複除去後の実数で再調整
        If j < fieldCount Then
            fieldCount = j
            ReDim Preserve selectedFields(1 To fieldCount)
        End If
    End If

    ' --- Excel起動（Late Binding） ---
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

    If xlApp Is Nothing Then
        MsgBox "Excelを起動できませんでした。" & vbCrLf & _
               "Microsoft Excelがインストールされているか確認してください。", _
               vbCritical, "MailSamalyCopy"
        GoTo CleanUp
    End If

    On Error GoTo ErrHandler

    xlApp.Visible = True
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    xlWs.Name = "MailData"

    ' --- ヘッダー出力 ---
    For colIdx = 1 To fieldCount
        xlWs.Cells(1, colIdx).Value = GetFieldName(selectedFields(colIdx))
    Next colIdx

    ' ヘッダー書式（青背景・白文字・太字）
    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, fieldCount))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- データ出力 ---
    rowIdx = 2
    For i = 1 To olSel.Count
        ' MailItemのみ処理（会議出席依頼等はスキップ）
        If TypeName(olSel.Item(i)) = "MailItem" Then
            Set oMail = olSel.Item(i)
            For colIdx = 1 To fieldCount
                xlWs.Cells(rowIdx, colIdx).Value = GetFieldValue(oMail, selectedFields(colIdx))
            Next colIdx
            rowIdx = rowIdx + 1
        End If
    Next i

    ' メール0件チェック（全て非MailItemだった場合）
    If rowIdx = 2 Then
        MsgBox "選択されたアイテムにメール（MailItem）が含まれていませんでした。" & vbCrLf & _
               "会議出席依頼等はスキップされます。", _
               vbInformation, "MailSamalyCopy"
        xlWb.Close False
        xlApp.Quit
        GoTo CleanUp
    End If

    ' --- 列幅自動調整 ---
    xlWs.Columns.AutoFit

    ' 本文列が広がりすぎないよう最大幅を制限
    For colIdx = 1 To fieldCount
        If selectedFields(colIdx) = 9 Or selectedFields(colIdx) = 10 Then
            If xlWs.Columns(colIdx).ColumnWidth > 80 Then
                xlWs.Columns(colIdx).ColumnWidth = 80
            End If
        End If
    Next colIdx

    ' フィルター設定
    xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(rowIdx - 1, fieldCount)).AutoFilter

    ' カーソルをA1に
    xlWs.Range("A1").Select

    MsgBox "完了しました。" & vbCrLf & _
           "出力件数: " & (rowIdx - 2) & " 件" & vbCrLf & _
           "出力項目: " & fieldCount & " 項目", _
           vbInformation, "MailSamalyCopy"

    GoTo CleanUp

ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "MailSamalyCopy"

    ' Excelが中途半端に開いていたら閉じる
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0

CleanUp:
    Set oMail = Nothing
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set olSel = Nothing
    Set olApp = Nothing
End Sub
