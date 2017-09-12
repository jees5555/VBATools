Attribute VB_Name = "VbArrayTool"
Option Explicit
Function ArrayAdd(arr As Variant, element As Variant, Optional index) As Variant
    If IsMissing(index) Then
        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
        arr(UBound(arr)) = element
        ArrayAdd = arr
    ElseIf Not IsNumeric(index) Then
         Err.Raise 513, , "index is not a number"
    Else
        Dim i As Long
        Dim isAdd As Boolean
        isAdd = False
        ReDim newArr(LBound(arr) To UBound(arr) + 1)
        For i = LBound(arr) To UBound(arr) + 1
            If isAdd Then
                newArr(i) = arr(i - 1)
            ElseIf i <> index Then
                newArr(i) = arr(i)
            Else
                newArr(i) = element
                isAdd = True
            End If
        Next i
        ArrayAdd = newArr
    End If
End Function
Function TwoDArrayAdd(TwoDArr As Variant, addArr As Variant, Optional index) As Variant
    Dim i As Long
    Dim j As Long
    j = LBound(addArr)
    Dim k As Long
    If IsMissing(index) Then index = UBound(TwoDArr, 1) + 1
    
    If Not IsNumeric(index) Then
         Err.Raise 513, , "index is not a number"
    Else
        Dim isAdd As Boolean
        isAdd = False
        ReDim newArr(LBound(TwoDArr, 1) To UBound(TwoDArr, 1) + 1, LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
        For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1) + 1
            If isAdd Then
                For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                    newArr(i, j) = TwoDArr(i - 1, j)
                Next j
            ElseIf i <> index Then
                For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                    newArr(i, j) = TwoDArr(i, j)
                Next j
            Else
                k = LBound(addArr)
                For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                    newArr(i, j) = addArr(k)
                    k = k + 1
                    If k > UBound(addArr) Then Exit For
                Next j
                isAdd = True
            End If
        Next i
        TwoDArrayAdd = newArr
    End If
End Function
Function TwoDArrayReplace(TwoDArr As Variant, replaceArr As Variant, index As Long) As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim isAdd As Boolean
    isAdd = False
    ReDim newArr(LBound(TwoDArr, 1) To UBound(TwoDArr, 1), LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        If i <> index Then
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i, j) = TwoDArr(i, j)
            Next j
        Else
            k = LBound(replaceArr)
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i, j) = replaceArr(k)
                k = k + 1
                If k > UBound(replaceArr) Then Exit For
            Next j
        End If
    Next i
    TwoDArrayReplace = newArr
End Function

Function ArrayContains(arr As Variant, element As Variant) As Boolean
    ArrayContains = False
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = element Then
            ArrayContains = True
            Exit Function
        End If
    Next i
End Function
Function TwoDArrayContains(TwoDArr As Variant, element As Variant) As Boolean
    TwoDArrayContains = False
    Dim i As Long
    Dim j As Long
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
            If TwoDArr(i, j) = element Then
                TwoDArrayContains = True
                Exit Function
            End If
        Next j
    Next i
End Function
Function ArrayIndexOf(arr As Variant, element As Variant) As Variant
    ArrayIndexOf = False
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = element Then
            ArrayIndexOf = i
            Exit Function
        End If
    Next i
End Function
Function ArrayLastIndexOf(arr As Variant, element As Variant) As Variant
    ArrayLastIndexOf = False
    Dim i As Long
    For i = UBound(arr) To LBound(arr) Step -1
        If arr(i) = element Then
            ArrayLastIndexOf = i
            Exit Function
        End If
    Next i
End Function
Function ArrayRemoveByIndex(arr As Variant, index As Long) As Variant
    Dim i As Long
    Dim isRemoved As Boolean
    isRemoved = False
    ReDim newArr(LBound(arr) To UBound(arr) - 1)
    For i = LBound(arr) To UBound(arr)
        If isRemoved Then
            newArr(i - 1) = arr(i)
        ElseIf i <> index Then
            newArr(i) = arr(i)
        Else
            isRemoved = True
        End If
    Next i
    ArrayRemoveByIndex = newArr
End Function
Function TwoDArrayRemoveByIndex(TwoDArr As Variant, index As Long) As Variant
    Dim i As Long
    Dim j As Long
    Dim isRemove As Boolean
    isRemove = False
    ReDim newArr(LBound(TwoDArr, 1) To UBound(TwoDArr, 1) - 1, LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        If isRemove Then
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i - 1, j) = TwoDArr(i, j)
            Next j
        ElseIf i <> index Then
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i, j) = TwoDArr(i, j)
            Next j
        Else
            isRemove = True
        End If
    Next i
    TwoDArrayRemoveByIndex = newArr
End Function
Function ArrayRemoveByElement(arr As Variant, element As Variant) As Variant
    Dim i As Long
    Dim j As Long
    Dim isRemoved As Boolean
    isRemoved = False
    ReDim newArr(LBound(arr) To UBound(arr) - 1)
    For i = LBound(arr) To UBound(arr)
        If isRemoved Then
            newArr(i - 1) = arr(i)
        ElseIf arr(i) <> element Then
            If i = UBound(arr) Then
                 ArrayRemoveByElement = arr
                 Exit Function
            End If
            newArr(i) = arr(i)
        Else
            isRemoved = True
        End If
    Next i
        ArrayRemoveByElement = newArr
End Function
Function TwoDArrayRemoveByFirstElement(TwoDArr As Variant, firstElement As Variant) As Variant
    Dim i As Long
    Dim j As Long
    Dim isRemoved As Boolean
    isRemoved = False
    ReDim newArr(LBound(TwoDArr, 1) To UBound(TwoDArr, 1) - 1, LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        If isRemoved Then
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i - 1, j) = TwoDArr(i, j)
            Next j
        ElseIf TwoDArr(i, 1) <> firstElement Then
            If i = UBound(TwoDArr, 1) Then
                 TwoDArrayRemoveByFirstElement = TwoDArr
                 Exit Function
            End If
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i, j) = TwoDArr(i, j)
            Next j
        Else
            isRemoved = True
        End If
    Next i
        TwoDArrayRemoveByFirstElement = newArr
End Function
Function ArrayRemoveRange(arr As Variant, fromIndex As Long, toIndex As Long) As Variant
    If fromIndex >= toIndex Then
        Err.Raise 514, , "fromIndex equal or greater than toIndex"
    End If
    Dim i As Long
    Dim removeNum As Long
    removeNum = toIndex - fromIndex
    
    ReDim newArr(LBound(arr) To UBound(arr) - removeNum)
    For i = LBound(arr) To UBound(arr)
        If i < fromIndex Then
            newArr(i) = arr(i)
        ElseIf i >= fromIndex And i < toIndex Then
            
        Else
            newArr(i - removeNum) = arr(i)
        End If
    Next i
    ArrayRemoveRange = newArr
End Function
Function TwoDArrayRemoveRange(TwoDArr As Variant, fromIndex As Long, toIndex As Long) As Variant
    If fromIndex >= toIndex Then
        Err.Raise 514, , "fromIndex equal or greater than toIndex"
    End If
    Dim i As Long
    Dim j As Long
    Dim removeNum As Long
    removeNum = toIndex - fromIndex

    ReDim newArr(LBound(TwoDArr, 1) To UBound(TwoDArr, 1) - removeNum, LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        If i < fromIndex Then
            For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i, j) = TwoDArr(i, j)
            Next j
        ElseIf i >= fromIndex And i < toIndex Then

        Else
             For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
                newArr(i - removeNum, j) = TwoDArr(i, j)
            Next j
        End If
    Next i
    TwoDArrayRemoveRange = newArr
End Function
Function TwoDArrayToOneDArray(TwoDArr As Variant, row As Long)
ReDim OneDArr(LBound(TwoDArr, 2) To UBound(TwoDArr, 2))
Dim i%
For i = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
    OneDArr(i) = TwoDArr(row, i)
Next i
TwoDArrayToOneDArray = OneDArr
End Function
Function UnionArray(arr1 As Variant, arr2 As Variant) As Variant
    Dim i%, j%
    j = LBound(arr2)
    Dim arr2Num&: arr2Num = UBound(arr2) - LBound(arr2) + 1
    ReDim UnionedArray(LBound(arr1) To UBound(arr1) + arr2Num)
    
    For i = LBound(arr1) To UBound(arr1) + arr2Num
        If i <= UBound(arr1) Then
            UnionedArray(i) = arr1(i)
        Else
            UnionedArray(i) = arr2(j)
            j = j + 1
        End If
    Next i
    UnionArray = UnionedArray
End Function
Sub TestArray(arr As Variant)
Dim i&
Dim str$
For i = LBound(arr) To UBound(arr)
    If arr(i) <> "" Then
        str = str & arr(i) & "  "
    Else
        str = str & "X" & "  "
    End If
Next i
MsgBox str
End Sub
Sub TestTwoDArray(TwoDArr As Variant)
    Dim i&
    Dim j&
    Dim str$
    MsgBox ""
    For i = LBound(TwoDArr, 1) To UBound(TwoDArr, 1)
        For j = LBound(TwoDArr, 2) To UBound(TwoDArr, 2)
            If TwoDArr(i, j) <> "" Then
                str = str & TwoDArr(i, j) & "  "
            Else
                str = str & "X" & "  "
            End If
            Next j
        str = str & vbCrLf
    Next i
    MsgBox str
End Sub
Sub BubbleSort(ByRef arr As Variant, Optional OrderBy As String)

    Dim First As Integer, Last As Long
    Dim i As Long, j As Long
    Dim Temp
    
    First = LBound(arr)
    Last = UBound(arr)
    For i = First To Last - 1
        For j = i + 1 To Last
            If OrderBy = "DESC" Then
                    If arr(i) < arr(j) Then
                    Temp = arr(j)
                    arr(j) = arr(i)
                    arr(i) = Temp
                End If
            Else
                If arr(i) > arr(j) Then
                    Temp = arr(j)
                    arr(j) = arr(i)
                    arr(i) = Temp
                End If
            End If
        Next j
    Next i
End Sub

Sub SelectionSort(ByRef arr As Variant, Optional OrderBy As String)
    'OrderBy="ASC"(Defult) or "DESC"
    Dim i&, j&
    Dim Temp As Double
    Dim m As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If OrderBy = "DESC" Then
                If arr(j) > arr(m) Then m = j
            Else
                If arr(j) < arr(m) Then m = j
            End If
        Next j
        Temp = arr(i)
        arr(i) = arr(m)
        arr(m) = Temp
    Next i
End Sub
Sub Quicksort(ByRef arr As Variant, Optional OrderBy As String)
    Call RunningQuicksort(arr, LBound(arr), UBound(arr))
    If OrderBy = "DESC" Then
        Call ReverseSort(arr)
    End If
End Sub
Sub RunningQuicksort(ByRef list As Variant, ByVal Min As Long, ByVal Max As Long)
    Dim med_value As Long
    Dim hi As Long
    Dim lo As Long
    Dim i As Long

    ' If min >= max, the list contains 0 or 1 items so it
    ' is sorted.
    If Min >= Max Then Exit Sub

    ' Pick the dividing value.
    i = Int((Max - Min + 1) * Rnd + Min)
    med_value = list(i)

    ' Swap it to the front.
    list(i) = list(Min)

    lo = Min
    hi = Max
    Do
        ' Look down from hi for a value < med_value.
        Do While list(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            list(lo) = med_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        list(lo) = list(hi)
        
        ' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While list(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            list(hi) = med_value
            Exit Do
        End If
        
        ' Swap the lo and hi values.
        list(hi) = list(lo)
    Loop
    
    ' Sort the two sublists.
    RunningQuicksort list, Min, lo - 1
    RunningQuicksort list, lo + 1, Max
End Sub
Sub ReverseSort(ByRef arr As Variant)
    Dim i As Long
    Dim Temp
    
    Dim Min As Long
    Dim Max As Long
    Dim Length As Long
    
    Min = LBound(arr)
    Max = UBound(arr)
    Length = Max - Min + 1
    
    For i = Min To (Min + Length) / 2
        Temp = arr(i)
        arr(i) = arr(Max - (i - Min))
        arr(Max - (i - Min)) = Temp
    Next i
End Sub
