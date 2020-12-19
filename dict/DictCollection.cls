VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Collection型でDictionary型のメソッドを用意する
'
'
Option Explicit

' メンバ変数
Private oKeys As Collection
Private oItems As Collection

'
' クラスの初期化
'
Private Sub Class_Initialize()
    Set oKeys = New Collection
    Set oItems = New Collection
End Sub

'
' Iterator
' https://docs.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/using-for-eachnext-statements
'
Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = oKeys.[_NewEnum]
End Property

'
' Count :   件数を返す
'
Public Function Count() As Long

    Count = oKeys.Count

End Function

'
' Add   :   キーを追加
'
'
Public Function Add(vKey As Variant, vItem As Variant) As Object

    oKeys.Add vKey, vKey
    oItems.Add vItem, vKey
    Set Add = Me

End Function

'
' Remove    :   キーを削除
'
'
Public Function Remove(vKey As Variant) As Object

    oKeys.Remove vKey
    oItems.Remove vKey
    Set Remove = Me

End Function

'
' RemoveAll :   キーを削除
'
'
Public Function RemoveAll() As Object

    Dim vVars As Variant
    For Each vVars In oKeys
        Call Remove(vVars)
    Next vVars
    Set RemoveAll = Me

End Function

'
' Item  :   値を返す
'
'
Public Function Item(vKey As Variant) As Variant

    Item = oItems.Item(vKey)

End Function

'
' Key   :   キーを置き換える
' http://www.niji.or.jp/home/toru/notes/17.html
'
Public Sub Key(vOldKey As Variant, vNewKey As Variant, vValue As Variant)

    Call Remove(vKey)
    Call Add(vNewKey, vValue)

End Sub

'
' Index    :   キーのIndex番号を返す
'
'
Public Function Index(vKey As Variant) As Long

    Dim i As Long
    i = 0
    Dim vVars As Variant
    For Each vVars In oKeys
        If vVars = vKey Then
            Index = i
            Exit Function
        End If
        i = i + 1
    Next vVars
    
    Index = -1

End Function

'
' Exists    :   存在チェック
'
'
Public Function Exists(vKey As Variant) As Boolean

    If Index(vKey) >= 0 Then
        Exists = True
    Else
        Exists = False
    End If

End Function

'
' Keys     :    キー配列を返す
'
Public Function Keys() As Variant()

    Dim InKeys As Variant
    ReDim InKeys(oKeys.Count() - 1) As Variant
    
    Dim i As Long
    i = 0
    Dim vVars As Variant
    For Each vVars In oKeys
        InKeys(i) = vVars
        i = i + 1
    Next vVars

    Keys = InKeys

End Function

'
' Items     :    アイテム配列を返す
'
Public Function Items() As Variant()

    Dim InItems As Variant
    ReDim InItems(oItems.Count() - 1) As Variant
   
    Dim i As Long
    i = 0
    Dim vVars As Variant
    For Each vVars In oItems
        InItems(i) = vVars
        i = i + 1
    Next vVars

    Items = InItems

End Function

'
' Sort  :    キーの値でソートして並び替える
'
Public Sub Sort()

    Dim vKeys As Variant
    vKeys = Keys()
    Call QuickSort(vKeys)

    Dim vItems As Variant
    ReDim vItems(UBound(vKeys)) As Variant
    
    Dim i As Long
    i = 0
    Dim vVars As Variant
    For Each vVars In vKeys
        vItems(i) = Item(vVars)
        i = i + 1
    Next vVars

    Call RemoveAll
    
    i = 0
    For i = 0 To UBound(vKeys)
        Call Add(vKeys(i), vItems(i))
    Next i

End Sub

'
' QuickSort :   クイックソート
'
Private Sub QuickSort(vList As Variant, Optional iFirst As Integer = 0, Optional iLast As Integer = -1)
    Dim iLeft                   As Integer      '// 左ループカウンタ
    Dim iRight                  As Integer      '// 右ループカウンタ
    Dim vMedian                 As Variant      '// 中央値
    Dim vTmp                    As Variant      '// 配列移動用バッファ
    
    '// ソート終了位置省略時は配列要素数を設定
    If (iLast = -1) Then
        iLast = UBound(vList)
    End If
    
    '// 中央値を取得
    vMedian = vList(Int((iFirst + iLast) / 2))
    
    iLeft = iFirst
    iRight = iLast
    
    Do
        '// 中央値の左側をループ
        Do
            '// 配列の左側から中央値より大きい値を探す
            If (vList(iLeft) >= vMedian) Then
                Exit Do
            End If
            
            '// 左側を１つ右にずらす
            iLeft = iLeft + 1
        Loop
        
        '// 中央値の右側をループ
        Do
            '// 配列の右側から中央値より大きい値を探す
            If (vMedian >= vList(iRight)) Then
                Exit Do
            End If
            
            '// 右側を１つ左にずらす
            iRight = iRight - 1
        Loop
        
        '// 左側の方が大きければここで処理終了
        If (iLeft >= iRight) Then
            Exit Do
        End If
        
        '// 右側の方が大きい場合は、左右を入れ替える
        vTmp = vList(iLeft)
        vList(iLeft) = vList(iRight)
        vList(iRight) = vTmp
        
        '// 左側を１つ右にずらす
        iLeft = iLeft + 1
        '// 右側を１つ左にずらす
        iRight = iRight - 1
    Loop
    
    '// 中央値の左側を再帰でクイックソート
    If (iFirst < iLeft - 1) Then
        Call QuickSort(vList, iFirst, iLeft - 1)
    End If
    
    '// 中央値の右側を再帰でクイックソート
    If (iRight + 1 < iLast) Then
        Call QuickSort(vList, iRight + 1, iLast)
    End If
    
End Sub

