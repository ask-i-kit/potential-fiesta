Option Explicit

' --- 右クリックメニュー登録・削除マクロ ---

' ▼▼▼ 設定項目 ▼▼▼
' 右クリックメニューに表示したい名前
Private Const MENU_CAPTION As String = "方眼紙テーブルをMarkdown変換"
' 右クリックで実行したいマクロの名前
Private Const TARGET_MACRO_NAME As String = "ExcelToMarkdownTable"
' ▲▲▲ 設定はここまで ▲▲▲


Sub AddCustomRightClickMenu()
    ' 機能: セルの右クリックメニューに、指定したマクロを実行するカスタム項目を追加します。

    Dim cellCommandBar As CommandBar
    Dim newButton As CommandBarButton

    ' 最初に、古いメニューが残っていれば削除して重複を防ぎます
    ' (メッセージは表示しないように引数を渡します)
    Call RemoveCustomRightClickMenu(showMsg:=False)

    ' セルの右クリックメニューオブジェクトを取得
    Set cellCommandBar = Application.CommandBars("Cell")

    ' メニューの末尾に新しいボタンを追加
    Set newButton = cellCommandBar.Controls.Add(Type:=msoControlButton)
    
    ' 追加したボタンのプロパティを設定
    With newButton
        .Caption = MENU_CAPTION
        ' OnActionには、このマクロが保存されているブック名を付けて指定すると確実です
        .OnAction = "'" & ThisWorkbook.Name & "'!" & TARGET_MACRO_NAME
        .BeginGroup = True ' 項目の上に区切り線を表示
        .Style = msoButtonCaption ' アイコンは表示せず、テキストのみにする
    End With

'    MsgBox "セルの右クリックメニューに「" & MENU_CAPTION & "」を追加しました。", vbInformation

    Set newButton = Nothing
    Set cellCommandBar = Nothing

End Sub

Sub RemoveCustomRightClickMenu(Optional showMsg As Boolean = True)
    ' 機能: 右クリックメニューに追加したカスタム項目を削除します。

    ' メニュー項目が存在しない場合にエラーにならないようにします
    On Error Resume Next
    
    ' 指定したキャプションを持つメニュー項目を削除
    Application.CommandBars("Cell").Controls(MENU_CAPTION).Delete
    
    ' エラーハンドリングを通常の状態に戻します
    On Error GoTo 0

    ' showMsgがTrueの場合のみメッセージを表示します
    If showMsg Then
        MsgBox "右クリックメニューから「" & MENU_CAPTION & "」を削除しました。", vbInformation
    End If

End Sub


Sub ExcelToMarkdownTable()
    ' 機能: 選択されたExcelの範囲をMarkdownテーブル形式に変換する。
    '       列の上から下まで同じ幅で結合されている場合、それを1列として扱う。
    ' 作成者: Gemini

    Dim selectionRange As Range
    Dim markdownOutput As String
    Dim columnMap As Object ' Collectionを使用
    Dim r As Long, c As Long, i As Long
    Dim c_index As Variant
    Dim colspan As Long
    Dim isUniform As Boolean
    Dim targetCell As Range, checkCell As Range
    Dim clipboard As Object

    ' エラーが発生した場合は処理を中断
    On Error GoTo ErrorHandler

    ' 現在選択されている範囲を取得
    Set selectionRange = Selection

    ' 何も選択されていない場合はメッセージを表示して終了
    If selectionRange Is Nothing Or selectionRange.Cells.Count = 0 Then
'        MsgBox "変換したい範囲を選択してください。", vbInformation
        Exit Sub
    End If

    ' --- 事前解析フェーズ: 出力すべき列を特定する ---
    Set columnMap = CreateObject("System.Collections.ArrayList") ' 順番を保持するリスト

    c = 1
    While c <= selectionRange.Columns.Count
        Set targetCell = selectionRange.Cells(1, c)

        ' この列が結合セルの左端でない場合、スキップして単独列として扱う
        If targetCell.MergeArea.Cells(1, 1).Address <> targetCell.Address Then
            columnMap.Add c
            c = c + 1
        Else
            ' この列が均一な結合列かを判定する
            isUniform = True
            colspan = targetCell.MergeArea.Columns.Count
            
            ' 選択範囲の2行目以降も同じ結合パターンかチェック
            If selectionRange.Rows.Count > 1 Then
                For r = 2 To selectionRange.Rows.Count
                    Set checkCell = selectionRange.Cells(r, c)
                    If checkCell.MergeArea.Columns.Count <> colspan Or _
                       checkCell.MergeArea.Cells(1, 1).Address <> checkCell.Address Then
                        ' パターンが異なる場合は均一ではない
                        isUniform = False
                        Exit For
                    End If
                Next r
            End If

            ' 判定結果に基づいて処理
            If isUniform Then
                ' 均一な結合列の場合：開始列を登録し、結合分だけカウンタを進める
                columnMap.Add c
                c = c + colspan
            Else
                ' 均一でない場合：この列を単独列として登録し、1つ進める
                columnMap.Add c
                c = c + 1
            End If
        End If
    Wend

    ' --- Markdown生成フェーズ ---
    markdownOutput = ""

    ' 行のループ
    For r = 1 To selectionRange.Rows.Count
        Dim currentRowOutput As String
        currentRowOutput = "|"

        ' 列マップに基づいてセルを取得
        For Each c_index In columnMap
            Set targetCell = selectionRange.Cells(r, c_index)
            currentRowOutput = currentRowOutput & " " & Replace(targetCell.Value, vbLf, "<br>") & " |"
        Next c_index
        
        markdownOutput = markdownOutput & currentRowOutput & vbCrLf

        ' 1行目の下にヘッダー区切り線を追加
        If r = 1 Then
            Dim separatorLine As String
            separatorLine = "|"
            For i = 0 To columnMap.Count - 1
                separatorLine = separatorLine & "---|"
            Next i
            markdownOutput = markdownOutput & separatorLine & vbCrLf
        End If
    Next r

    ' --- クリップボードへのコピー ---
    Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.SetText markdownOutput
    clipboard.PutInClipboard

    ' 完了メッセージ
'    MsgBox "Markdown形式のテーブルをクリップボードにコピーしました。", vbInformation

    ' オブジェクトを解放
    Set clipboard = Nothing
    Set columnMap = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Set clipboard = Nothing
    Set columnMap = Nothing
End Sub

