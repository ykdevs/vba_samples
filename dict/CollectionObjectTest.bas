Attribute VB_Name = "CollectionObjectTest"
'
' Collection型の扱い
'
' https://tonari-it.com/excel-vba-dictionary-keys-items/
'

Sub CollectionObjectTest()

    ' Collection型の作成
    Dim oDict As Collection
    Set oDict = New Collection

    ' アイテムの追加
    '
    ' Collection.Add Item, Key, Index1, Index2
    '
    oDict.Add "Item1", "Key1"
    oDict.Add "Item2", "Key2"
    oDict.Add "Item3", "Key3"

    ' 件数
    '
    ' Collection.Count
    '
    Dim nCount As Integer
    nCount = oDict.Count
    Debug.Print "件数"
    Debug.Print "Count : " + str(nCount)

    ' For文でアイテムを取得
    Debug.Print "For文でアイテムを取得"
    Dim i As Integer
    For i = 1 To nCount
        Debug.Print str(i) + " : " + oDict.Item(i)
    Next i

    ' ForEach文でアイテムを取得
    Debug.Print "ForEach文でアイテムを取得"
    i = 1
    For Each vVars In oDict
        Debug.Print str(i) + " : " + vVars
        i = i + 1
    Next

End Sub

