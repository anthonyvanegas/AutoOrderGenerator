Attribute VB_Name = "Main"
Option Explicit

Sub enterOrders()
    'Clear worksheet
    Call clearSheet
    
    'Get item dictionary                    (VENDOR# -> ITEM#)
    Dim itemDict: Set itemDict = New Dictionary
    itemDict.loadItems (0)
    
    'Get item row dictionary                (ITEM# -> ROW#)
    Dim rowDict: Set rowDict = New Dictionary
    rowDict.loadItems (1)
    
    'Get store column dictionary            (STORE# -> COLUMN#)
    Dim columnDict: Set columnDict = New Dictionary
        columnDict.loadItems (2)

    'Get store orders                       (ITEM# -> STORE# -> QTY#)
    Dim order: Set order = New RawOrders
    Dim storeOrders() As String: storeOrders = order.getStoreOrders(itemDict)
    Set order = Nothing
    
    'Loop through each item num#
    Dim itemCount As Integer: itemCount = 0
    Dim itemPointer As String: itemPointer = storeOrders(itemCount, 0, 0)
    Dim storeCount As Integer
    Dim storePointer As String
    Dim sheetRef As Worksheet: Set sheetRef = ThisWorkbook.Sheets("Order Guide")
    Dim validFlag As Boolean: validFlag = True
    While (itemPointer <> "" And validFlag)
        storeCount = 0
        storePointer = storeOrders(itemCount, storeCount, 1)
        While (storePointer <> "" And validFlag)
            Dim row As String: row = rowDict.getValue(CStr(itemDict.getValue(CLngLng(itemPointer))))
            Dim column As String: column = columnDict.getValue(CLngLng(storePointer))
            If (row <> "N\A" And column <> "N\A") Then
                sheetRef.Cells(CInt(row), CInt(column)).value = CInt(storeOrders(itemCount, storeCount, 2))
                storeCount = storeCount + 1
                storePointer = storeOrders(itemCount, storeCount, 1)
            ElseIf (row = "N\A") Then
                MsgBox ("No Location on order guide for: " & itemPointer)
                MsgBox ("Re-run the Macro once item is added into the order guide")
                validFlag = False
            ElseIf (column = "N\A") Then
                MsgBox ("Store #" & storePointer & " is not located on the order guide")
                MsgBox ("Re-run the Macro once store is added into the order guide")
                validFlag = False
            End If
        Wend
        itemCount = itemCount + 1
        itemPointer = storeOrders(itemCount, 0, 0)
    Wend
    
    'Clean up
    Set itemDict = Nothing
    Set rowDict = Nothing
    Set columnDict = Nothing
    
    'Report if we have fully entered the orders
    If (validFlag) Then
        MsgBox ("Orders entered!")
    End If
End Sub

Private Function clearSheet()
    Dim sheetReference As Worksheet: Set sheetReference = ThisWorkbook.Worksheets("Order Guide")
    Dim lastColumn As Integer: lastColumn = sheetReference.Cells(4, Columns.Count).End(xlToLeft).column
    Dim lastRow As Integer: lastRow = sheetReference.Cells(Rows.Count, 1).End(xlUp).row
    
    sheetReference.Range(sheetReference.Cells(5, 3).Address & ":" & sheetReference.Cells(lastRow - 1, lastColumn - 1).Address).ClearContents
End Function

