VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim size As Integer
Dim length As Integer
Dim keyArr() As LongLong
Dim valueArr() As String

Private Sub Class_Initialize()
    length = 0
End Sub

Public Function loadItems(ByVal loadSelect As Integer)
    Dim sheetReference As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    
    Select Case loadSelect
        Case 0      'Get a list of vendor conversions (VENDOR ITEM# -> SBT ITEM#)
            Set sheetReference = ThisWorkbook.Sheets("Item Conversion")
            lastRow = sheetReference.Cells(Rows.Count, 1).End(xlUp).row
            size = lastRow - 3
            ReDim keyArr(size)
            ReDim valueArr(size)
            For i = 3 To lastRow + 1
                If sheetReference.Cells(i, 1).value <> "" Then
                    Call addItem(CLngLng(sheetReference.Cells(i, 1).value), CStr(sheetReference.Cells(i, 2).value))
                End If
            Next i
        Case 1      'Get a list of row locations on order guide (SBT ITEM# -> ORDER GUIDE ROW#)
            Set sheetReference = ThisWorkbook.Sheets("Order Guide")
            lastRow = sheetReference.Cells(Rows.Count, 1).End(xlUp).row
            size = sheetReference.Cells(Rows.Count, 1).End(xlUp).row - 6
            ReDim keyArr(size)
            ReDim valueArr(size)
            For i = 5 To sheetReference.Cells(Rows.Count, 1).End(xlUp).row
                If sheetReference.Cells(i, 1).value <> "" Then
                    Call addItem(getAscii(CStr(sheetReference.Cells(i, 1).value)), CStr(sheetReference.Cells(i, 2).row))
                End If
            Next i
        Case 2      'Get a list of columns locations on order guide (STORE# -> ORDER GUIDE COLUMN#)
            Set sheetReference = ThisWorkbook.Sheets("Order Guide")
            Dim lastColumn As Integer: lastColumn = sheetReference.Cells(4, Columns.Count).End(xlToLeft).column
            size = lastColumn - 4
            ReDim keyArr(size)
            ReDim valueArr(size)
            Dim pointer
            For i = 3 To lastColumn - 1
                pointer = sheetReference.Cells(4, i).value
                Call addItem(CLngLng(pointer), CStr(i))
            Next i
    End Select
    
    Call Quicksort(0, CLng(size))
    
End Function

Public Function getValue(ByVal key) As Variant
    If VarType(key) = 20 Then
        getValue = CStr(recursiveBinSearch(0, size, key))
    ElseIf VarType(key) = 8 Then
        getValue = recursiveBinSearch(0, size, getAscii(CStr(key)))
    End If
End Function

Public Function printDictionary()
    Dim i As Integer
    Debug.Print ("Printing Dictionary")
    Debug.Print ("")
    For i = 0 To length - 1
        Debug.Print ("  Index: " & CStr(i))
        Debug.Print ("      Key: " & CStr(keyArr(i)))
        Debug.Print ("      Value: " & CStr(valueArr(i)))
        Debug.Print ("")
    Next i
End Function

Private Function addItem(ByVal key, ByVal value)
    If (length < size + 1) Then
        keyArr(length) = key
        valueArr(length) = value
        length = length + 1
    End If
End Function

Private Function recursiveBinSearch(ByVal low As Integer, ByVal high As Integer, ByVal key As Variant) As Variant
    Dim mid As Integer
    
    If (low <= high) Then
        mid = (low + high) / 2
        If (key = keyArr(mid)) Then
            recursiveBinSearch = valueArr(mid)
            Exit Function
        ElseIf (key < keyArr(mid)) Then
            recursiveBinSearch = recursiveBinSearch(low, mid - 1, key)
            Exit Function
        Else
            recursiveBinSearch = recursiveBinSearch(mid + 1, high, key)
            Exit Function
        End If
    End If
    
    recursiveBinSearch = "N\A"
End Function

Private Function Quicksort(arrLbound As Long, arrUbound As Long)
    Dim pivotVal As LongLong
    Dim vSwap1   As LongLong
    Dim vSwap2   As String
    Dim tmpLow   As Long
    Dim tmpHi    As Long
     
    tmpLow = arrLbound
    tmpHi = arrUbound
    pivotVal = keyArr((arrLbound + arrUbound) \ 2)
 
    While (tmpLow <= tmpHi) 'divide
       While (keyArr(tmpLow) < pivotVal And tmpLow < arrUbound)
          tmpLow = tmpLow + 1
       Wend
      
       While (pivotVal < keyArr(tmpHi) And tmpHi > arrLbound)
          tmpHi = tmpHi - 1
       Wend
     
       If (tmpLow <= tmpHi) Then
          vSwap1 = keyArr(tmpLow)
          keyArr(tmpLow) = keyArr(tmpHi)
          keyArr(tmpHi) = vSwap1
          vSwap2 = valueArr(tmpLow)
          valueArr(tmpLow) = valueArr(tmpHi)
          valueArr(tmpHi) = vSwap2
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend
 
  If (arrLbound < tmpHi) Then Quicksort arrLbound, tmpHi
  If (tmpLow < arrUbound) Then Quicksort tmpLow, arrUbound
End Function

Private Function getAscii(ByRef str As String) As LongLong
    Dim i As Integer
 
    For i = 1 To Len(str)
        getAscii = getAscii & CStr(Asc(mid$(str, i, 1)))
    Next i
End Function


