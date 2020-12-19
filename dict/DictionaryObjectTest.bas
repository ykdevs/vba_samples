Attribute VB_Name = "DictionaryObjectTest"
'
' Dictionary蝙九・謇ｱ縺・
'
' https://www.sejuku.net/blog/29736
'

Sub DictionaryObjectTest()

    ' Dictionary蝙九・菴懈・
    Dim oDict As Object
    Set oDict = CreateObject("Scripting.Dictionary")

    ' 繧｢繧､繝・Β縺ｮ霑ｽ蜉
    '
    ' Dictionary.Add Key, Item
    '
    oDict.Add "Key1", "Item1"
    oDict.Add "Key2", "Item2"
    oDict.Add "Key3", "Item3"

    ' Item/Key
    Debug.Print oDict.Item("Key1")


    ' 莉ｶ謨ｰ
    Dim nCount As Integer
    nCount = oDict.Count
    Debug.Print "莉ｶ謨ｰ"
    Debug.Print "Count : " + str(nCount)

    ' 繧ｭ繝ｼ驟榊・縺ｮ蜿門ｾ・
    Dim vKeys As Variant
    vKeys = oDict.Keys
    
    ' 繧ｭ繝ｼ縺ｮ繧ｵ繧､繧ｺ
    Dim nLen As Integer
    nLen = UBound(vKeys)
    
    
    ' For譁・〒繧､繝ｳ繝・ャ繧ｯ繧ｹ縺ｧ繧｢繧､繝・Β繧貞叙蠕・
    Debug.Print "For譁・〒繧､繝ｳ繝・ャ繧ｯ繧ｹ縺ｧ繧｢繧､繝・Β繧貞叙蠕・"
    For i = 0 To nLen
        Debug.Print str(i) + " : " + oDict.Items()(i)
    Next i
    
    ' For譁・〒繧ｭ繝ｼ蜷阪〒繧｢繧､繝・Β繧貞叙蠕・
    Debug.Print "For譁・〒繧ｭ繝ｼ蜷阪〒繧｢繧､繝・Β繧貞叙蠕・"
    For i = 0 To nLen
        Debug.Print vKeys(i) + " : " + oDict.Item(vKeys(i))
    Next i

    ' 繝舌Μ繝･繝ｼ驟榊・縺ｮ蜿門ｾ・
    Dim vItems As Variant
    vItems = oDict.Items
    ' For譁・〒繝舌Μ繝･繝ｼ繧貞叙蠕・
    Debug.Print "For譁・〒繝舌Μ繝･繝ｼ繧貞叙蠕・"
    For i = 0 To nLen
        Debug.Print vItems(i)
    Next i

    ' Exists縺ｧ蟄伜惠遒ｺ隱阪＠縺ｦ蜑企勁
    If oDict.Exists("Key2") Then
        Debug.Print "Exists縺ｧ蟄伜惠遒ｺ隱阪＠縺ｦ蜑企勁"
        oDict.Remove ("Key2")
    End If

    ' ForEach譁・〒繧｢繧､繝・Β繧貞叙蠕・
    Debug.Print "ForEach譁・〒繧｢繧､繝・Β繧貞叙蠕・"
    For Each vVars In oDict
        Debug.Print vVars + " : " + oDict.Item(vVars)
    Next

End Sub


'' Dictionary繧貞盾辣ｧ蠑墓焚縺ｫ縺励√％繧後ｒ繧ｽ繝ｼ繝医☆繧狗ｴ螢顔噪繝励Ο繧ｷ繝ｼ繧ｸ繝｣縲・
'' Variant縺ｮ莠梧ｬ｡蜈・・蛻励・縺ｾ縺ｾ謇ｱ縺・◆縺九▲縺溘′縲∝ｼ墓焚縺ｫ縺ｧ縺阪↑縺・ｈ縺・↑縺ｮ縺ｧ驟榊・繧剃ｽｿ逕ｨ縺励※縺・ｋ
'' Key繧７alue繧・ong蝙九・縺ｿ繧定ｨｱ螳ｹ
Public Sub DictQuickSort(ByRef oDict As Object, Optional sIndex As String = "Key")
  
    Dim nIndex As Integer
    If sIndex = "Key" Then
        nIndex = 0
    Else
        nIndex = 1
    End If
  
    Dim i As Long
    Dim j As Long
    Dim dicSize As Long
    Dim varTmp(ARRAYSIZE, 2) As Long
  
    dicSize = dic.Count
  
    ' Dictionary縺檎ｩｺ縺九√し繧､繧ｺ縺・莉･荳九〒縺ゅｌ縺ｰ繧ｽ繝ｼ繝井ｸ崎ｦ・
    If dic Is Nothing Or dicSize < 2 Then
        Exit Sub
    End If
  
    ' Dictionary縺九ｉ莠悟・驟榊・縺ｫ霆｢蜀・
    i = 0
    Dim Key As Variant
    For Each Key In dic
        varTmp(i, 0) = Key
        varTmp(i, 1) = dic(Key)
        i = i + 1
    Next
  
    '繧ｯ繧､繝・け繧ｽ繝ｼ繝・
    Call QuickSort(varTmp, 0, dicSize - 1, nIndex)
  
    dic.RemoveAll
  
    For i = 0 To dicSize - 1
        dic(varTmp(i, 0)) = varTmp(i, 1)
    Next
End Sub


'
' Long蝙九・莠梧ｬ｡蜈・・蛻励ｒ蜿励￠蜿悶ｊ縲√％繧後・・貞・逶ｮ・・alue・峨〒繧ｯ繧､繝・け繧ｽ繝ｼ繝医☆繧・
'
Private Sub QuickSort(ByRef targetVar() As Long, ByVal min As Long, ByVal max As Long, nIndex As Integer)
    Dim i, j As Long
    Dim tmp As Long
    Dim pivot As Long
    
    If min < max Then
        i = min
        j = max
        pivot = med3(targetVar(i, nIndex), targetVar(Int(i + j / 2), nIndex), targetVar(j, nIndex))
        Do
            Do While targetVar(i, nIndex) < pivot
                i = i + 1
            Loop
            Do While pivot < targetVar(j, nIndex)
                j = j - 1
            Loop
            If i >= j Then Exit Do
            
            tmp = targetVar(i, 0)
            targetVar(i, 0) = targetVar(j, 0)
            targetVar(j, 0) = tmp
        
            tmp = targetVar(i, 1)
            targetVar(i, 1) = targetVar(j, 1)
            targetVar(j, 1) = tmp
        
            i = i + 1
            j = j - 1
        
        Loop
        Call QuickSort(targetVar, min, i - 1, nIndex)
        Call QuickSort(targetVar, j + 1, max, nIndex)
        
    End If
End Sub


'' Long, y, z 繧定ｾ樊嶌鬆・ｯ碑ｼ・＠莠檎分逶ｮ縺ｮ繧ゅ・繧定ｿ斐☆
Private Function med3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    If x < y Then
        If y < z Then
            med3 = y
        ElseIf z < x Then
            med3 = x
        Else
            med3 = z
        End If
    Else
        If z < y Then
            med3 = y
        ElseIf x < z Then
            med3 = x
        Else
            med3 = z
        End If
    End If
End Function

