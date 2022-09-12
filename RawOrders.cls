VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RawOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim filePath As String
Dim storeOrders() As String

Private Sub Class_Initialize()
    ReDim storeOrders(50, 10, 2)
    filePath = "/Users/anthonyvanegas/Desktop/"
End Sub

Public Function getStoreOrders(ByVal itemDict As Dictionary)

    Dim lineArray() As String: lineArray = getRawLineArray()
    
    Dim i As Integer
    Dim itemCount As Integer: itemCount = 0
    For i = 0 To UBound(lineArray)
        If lineArray(i) <> "" Then
            Dim pointerArray() As String: pointerArray = Split(lineArray(i), " ")
            Dim pointer As String: pointer = Trim(pointerArray(0))
            If Len(pointer) = 8 And IsNumeric(pointer) Then             'Find line that contains item number
                If itemDict.getValue(CLngLng(pointer)) <> "N\A" Then
                    Dim itemNumber As LongLong: itemNumber = CLngLng(pointer)
                    Dim j As Integer
                    
                    Dim storeNumberArray() As Integer                   'Get Store Numbers for orders
                    If (InStr(lineArray(i + 2), "P")) Then              'Skip over the chance that it is not a kings store
                        pointerArray = Split(lineArray(i + 2), "P")
                        ReDim storeNumberArray(UBound(pointerArray) - 1) As Integer
                        For j = 0 To UBound(pointerArray)
                            pointer = pointerArray(j)
                            If (IsNumeric(pointer)) Then
                                storeNumberArray(j) = CInt(pointer)
                            End If
                        Next j
                        
                        Dim orderNumberArray() As Integer                   'Get order qtys
                        pointerArray = Split(lineArray(i + 3), " ")
                        ReDim orderNumberArray(UBound(storeNumberArray)) As Integer
                        Dim orderCount As Integer: orderCount = 0
                        For j = 0 To UBound(pointerArray)
                            pointer = pointerArray(j)
                            If (IsNumeric(pointer)) Then
                                orderNumberArray(orderCount) = CInt(pointer)
                                orderCount = orderCount + 1
                            End If
                        Next j
                        
                        For j = 0 To UBound(storeNumberArray)               'Add items to the array
                            storeOrders(itemCount, j, 0) = itemNumber
                            storeOrders(itemCount, j, 1) = storeNumberArray(j)
                            storeOrders(itemCount, j, 2) = orderNumberArray(j)
                        Next j
                        
                        itemCount = itemCount + 1
                        i = i + 4
                    End If
                Else
                    Dim sheetReference As Worksheet: Set sheetReference = ThisWorkbook.Sheets("Item Conversion")
                    Dim lastRow As Integer: lastRow = sheetReference.Cells(Rows.Count, 3).End(xlUp).row
                    If sheetReference.Cells(lastRow, 3).value <> pointer Then
                        sheetReference.Cells(lastRow + 1, 3).value = pointer
                    End If
                End If
            End If
        End If
    Next i
    
    getStoreOrders = storeOrders
    
End Function

Private Function getRawLineArray() As String()
    Dim textFile As Integer
    Dim fileContent As String
    
    'Get file name from user
    On Error Resume Next
        'filePath = filePath & Application.InputBox( _
            Title:="Export Assistant", _
            prompt:="Enter File Name: ", _
            Type:=2) & ".txt"
        filePath = filePath & "REPORT.txt"
    On Error GoTo 0
    
    textFile = FreeFile
    Open filePath For Input As textFile
    fileContent = Input(LOF(textFile), textFile)
    Close textFile
    
    getRawLineArray = Split(fileContent, vbCr)
End Function

