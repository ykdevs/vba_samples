VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SjisToUtf8NoBOM(a_sFrom, a_sTo)
    Dim streamRead  As New ADODB.Stream '// 読み込みデータ
    Dim streamWrite As New ADODB.Stream '// 書き込みデータ
    Dim sText       As Variant          '// ファイルデータ
    
    '// ファイル読み込み
    streamRead.Type = adTypeText
    streamRead.Charset = "Shift-JIS"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// 改行コードCRLFをLFに変換
    sText = streamRead.ReadText
    sText = Replace(sText, vbCrLf, vbLf)
    
    '// ファイル書き込み
    streamWrite.Type = adTypeText
    streamWrite.Charset = "UTF-8"
    streamWrite.Open
    
    '// Shift-JISファイルのデータをUTF-8ファイルにコピー
    Call streamWrite.WriteText(sText)
    
    '// バイナリモードで書き込み済みデータ開始位置をBOM分の3バイトずらす
    streamWrite.Position = 0
    streamWrite.Type = adTypeBinary
    streamWrite.Position = 3
    
    '// 3バイトずらした状態でのデータを取得
    sText = streamWrite.Read
    
    '// ずらした開始位置を元に戻す
    streamWrite.Position = 0
    
    '// BOMが除去されたデータを先頭から書き込み直す
    Call streamWrite.Write(sText)
    
    '// 現時点の末尾を終端とし、直前に書き込まれていた3バイトをデータ対象外とする
    streamWrite.SetEOS
    
    '// 保存
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub

Public Sub Utf8ToSjis(a_sFrom, a_sTo)
    Dim streamRead  As New ADODB.Stream '// 読み込みデータ
    Dim streamWrite As New ADODB.Stream '// 書き込みデータ
    Dim sText                           '// ファイルデータ
    
    '// ファイル読み込み
    streamRead.Type = adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// 改行コードLFをCRLFに変換
    sText = streamRead.ReadText
    sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    '// ファイル書き込み
    streamWrite.Type = adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    '// データ書き込み
    Call streamWrite.WriteText(sText)
    
    '// 保存
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub

