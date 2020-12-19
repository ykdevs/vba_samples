Attribute VB_Name = "DictCollectionObjectTest"
'
' Collection型でDictionary型と同じようなことができるように
'
Sub DictCollectionObjectTest()

    ' Dictionary型の作成
    Dim oDict As DictCollection
    Set oDict = New DictCollection

    ' アイテムの追加
    '
    ' Dictionary.Add Key, Item
    '
    oDict.Add "Key1", "Item1"
    oDict.Add "Key2", "Item2"
    oDict.Add "Key3", "Item3"

    ' Item/Key
    Debug.Print oDict.Item("Key1")

    ' 件数
    Dim nCount As Integer
    nCount = oDict.Count
    Debug.Print "件数"
    Debug.Print "Count : " + str(nCount)

    ' キー配列の取得
    Dim vKeys As Variant
    vKeys = oDict.Keys
    
    ' キーのサイズ
    Dim nLen As Integer
    nLen = UBound(vKeys)
    
    ' For文でインデックスでアイテムを取得
    Debug.Print "For文でインデックスでアイテムを取得"
    For i = 0 To nLen
        Debug.Print str(i) + " : " + oDict.Items()(i)
    Next i
    
    ' For文でキー名でアイテムを取得
    Debug.Print "For文でキー名でアイテムを取得"
    For i = 0 To nLen
        Debug.Print vKeys(i) + " : " + oDict.Item(vKeys(i))
    Next i

    ' バリュー配列の取得
    Dim vItems As Variant
    vItems = oDict.Items
    ' For文でバリューを取得
    Debug.Print "For文でバリューを取得"
    For i = 0 To nLen
        Debug.Print vItems(i)
    Next i

    ' Existsで存在確認して削除
    If oDict.Exists("Key2") Then
        Debug.Print "Existsで存在確認して削除"
        oDict.Remove ("Key2")
    End If

    ' ForEach文でアイテムを取得
    Debug.Print "ForEach文でアイテムを取得"
    For Each vVars In oDict
        Debug.Print vVars + " : " + oDict.Item(vVars)
    Next

    oDict.Add "Key6", "Item6"
    oDict.Add "Key4", "Item4"
    oDict.Add "Key5", "Item5"
    oDict.Sort
    
    ' ForEach文でアイテムを取得
    Debug.Print "ForEach文でアイテムを取得"
    Dim nIdx As Long
    For Each vVars In oDict
        nIdx = oDict.Index(vVars)
        Debug.Print str(nIdx) + " : " + vVars + " : " + oDict.Items()(nIdx) + " : " + oDict.Item(vVars)
    Next

End Sub


