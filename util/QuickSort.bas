Attribute VB_Name = "QuickSort"
'
'
'
Sub QuickSortTest()
    Dim vKeys As Variant
    vKeys = Array(3, 7, 6, 9, 2, 4, 1, 5, 8)
    
    For Each vVars In vKeys
        Debug.Print str(vVars)
    Next vVars

    Call QuickSort(vKeys)

    For Each vVars In vKeys
        Debug.Print str(vVars)
    Next vVars

End Sub


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
