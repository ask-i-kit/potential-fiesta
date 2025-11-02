Option Explicit

Public Function SelectFolder(ByRef FolderPath As String, _
                                ByRef Caption As String) _
                                              As String
    ' FolderPath: 最初に開くフォルダ
    ' Caption   : ダイアログのキャプションに表示する文字列
    
    Dim Output As String
    Dim FileDialog As FileDialog
    Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FileDialog
        .Filters.Clear
        .Title = Caption
        
        .InitialFileName = FolderPath & "\"
        If .Show = True Then
            Output = .SelectedItems(1)
        Else
            Output = ""
        End If
    End With
    
    SelectFolder = Output
            
End Function




Private Sub test()
    Dim test As String
    test = SelectFolder("C:\", "フォルダを選択してください。")
    
    Debug.Print (test)
End Sub
