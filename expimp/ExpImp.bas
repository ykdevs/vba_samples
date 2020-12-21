Attribute VB_Name = "ExpImp"

'
' モジュール読込み
'
Public Sub モジュール読込み()
    Static oFileUtil As New FileUtil
    Dim sPathName As String
    sPathName = oFileUtil.SelectDir()
    Call ImportAll(sPathName)
End Sub

'
' モジュールExport
'
Public Sub モジュール出力()
    Static oFileUtil As New FileUtil
    Dim sPathName As String
    sPathName = oFileUtil.SelectDir()
    Call ExportAll(sPathName)
End Sub

'
' モジュールの読込み
'
' sPath      IN      フォルダパス
'
Private Sub ImportAll(sPath As String)
    On Error Resume Next
    
    Static oCodeUtil As New CodeUtil
    Static oFileUtil As New FileUtil
    
    Dim oFso        As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sArModule() As String                   '// モジュールファイル配列
    Dim sModule                                 '// モジュールファイル
    Dim sExt        As String                   '// 拡張子
    Dim iMsg                                    '// MsgBox関数戻り値
    
    iMsg = MsgBox("同名のモジュールは上書きします。よろしいですか？", vbOKCancel, "上書き確認")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// 全モジュールのファイルパスを取得
    Dim FileList As Collection
    Set FileList = oFileUtil.GetFiles(sPath, "\*")
        
    '// 全モジュールをループ
    With ActiveWorkbook.VBProject
        For Each sModule In FileList
            '// 拡張子を小文字で取得
            sExt = LCase(oFso.GetExtensionName(sModule))
            
            '// 拡張子がcls、frm、basのいずれかの場合
            If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
                sFileName = sPath + "\\" + sModule
                sTempName = sFileName + ".tmp"
                Call oCodeUtil.Utf8ToSjis(sFileName, sTempName)
                
                '// 同名モジュールを削除
                Call .VBComponents.Remove(.VBComponents(oFso.GetBaseName(sModule)))
                '// モジュールを追加
                Call .VBComponents.Import(sTempName)
                '// Import確認用ログ出力
                Debug.Print sModule
            
                Kill sTempName
            End If
        Next
    End With
End Sub

'
' モジュールの出力
'
' sPath      IN      フォルダパス
'
Private Sub ExportAll(sPath As String)
    On Error Resume Next
    
    Static oCodeUtil As New CodeUtil
    
    Dim sFileName As String
    With ActiveWorkbook.VBProject
        Dim i As Integer
        For i = 1 To .VBComponents.Count
            Debug.Print "Type: " & .VBComponents(i).Type
            Debug.Print "Name: " & .VBComponents(i).Name
            If .VBComponents(i).Type = 1 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".bas"
            ElseIf .VBComponents(i).Type = 2 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".cls"
            Else
                sFileName = ""
            End If
            
            If sFileName <> "" Then
                sTempName = sFileName + ".tmp"
                .VBComponents(i).Export sTempName
                
                Call oCodeUtil.SjisToUtf8NoBOM(sTempName, sFileName)
                Kill sTempName
            End If
        Next i
    End With
End Sub



