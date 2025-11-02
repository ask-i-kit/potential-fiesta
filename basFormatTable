Option Explicit


Sub SelectionFormatTable()
    ' 機能: 選択されたExcelの範囲をテーブル形式に罫線とヘッダ行を作成する。
    
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
    End With
    With Selection.Interior
        .ColorIndex = xlNone
    End With
    With Selection.Borders
        ' まず、すべての罫線を「細い実線」に設定
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' その後、内側の横線だけ「極細線」に変更
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
        
    With Selection.Rows(1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Rows(1).Interior
        .Color = RGB(153, 204, 255)
    End With

End Sub

