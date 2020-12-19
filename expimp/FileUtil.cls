VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
' フォルダを選択するダイアログを表示
'
' 初期はマクロファイルと同じフォルダを開く
'
' SelectDir     OUT     フォルダ名フルパス
'
Public Function SelectDir()
    ' 新しいファイルを開く
    Dim FileSysObj As Object
    Dim sPathName As String
    Set FileSysObj = CreateObject("Scripting.FileSystemObject")
    sPathName = FileSysObj.GetParentFolderName(ActiveWorkbook.FullName)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = sPathName + "\"
        .AllowMultiSelect = False
        If .Show = True Then
            SelectDir = .SelectedItems(1)
        Else
            MsgBox "処理を中止します。"
            End
        End If
    End With

End Function

'
' 特定のフォルダのファイルを開く
'
' sPathName     IN      フォルダパス
' sExt          IN      拡張子の指定
' cFileList     OUT     ファイルリスト
'
Public Function GetFiles(sPathName As String, sExt As String) As Collection
    Dim sFile As String
    Dim cFileList As Collection
    Set cFileList = New Collection
    sFile = Dir(sPathName & sExt)
    Do While sFile <> ""
        cFileList.Add Item:=sFile
        sFile = Dir()
    Loop
    Set GetFiles = cFileList
End Function
