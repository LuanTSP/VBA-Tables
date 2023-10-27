VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PROPERTIES
Public origin As range

' INIT FUNCTION
Public Function Init(origin As range)
    Set Me.origin = origin
End Function

' UTILITY FUNCTIONS
Public Function RowsCount() As Integer
    RowsCount = Me.origin.CurrentRegion.Rows.Count
End Function

Public Function ColumnsCount() As Integer
    ColumnsCount = Me.origin.CurrentRegion.columns.Count
End Function

Public Function headerRange() As range
    Set headerRange = Me.origin.Worksheet.range(Me.origin.CurrentRegion.Rows.item(1).Address)
End Function

Public Function header() As Collection
    Dim headerRange As range
    Set headerRange = Me.headerRange()
    
    Dim data As New Collection, item As Variant
    For Each item In headerRange.Cells
        Call data.Add(item.value)
    Next item
    
    Set header = data
End Function

Public Sub CreateIDColumn(Optional index As Integer = 1)
    Dim data As New Collection, lines As Integer
    lines = Me.RowsCount() - 1
    
    Dim header As Collection
    Set header = Me.header()
    
    Dim item As Variant
    For Each item In header
        If item = "ID" Then
            GoTo Sair
        End If
    Next item
    
    For i = 1 To lines
        Call data.Add(i)
    Next i
    
    Call Me.AddColumn("ID", data, index)
Sair:
    Exit Sub
End Sub

Public Function AllIndexWhere(header As String, value) As Collection
    Dim column As Collection, indexes As New Collection
    Set column = Me.ColumnByName(header)
    
    Dim item As Variant, i As Integer
    i = 1
    For Each item In column
        If item = value Then
            Call indexes.Add(i)
        End If
        i = i + 1
    Next item
    
    Set AllIndexWhere = indexes
End Function

Public Function IndexWhere(header As String, value) As Integer
    Dim column As Collection, indexes As New Collection
    Set column = Me.ColumnByName(header)
    
    Dim item As Variant, i As Integer
    i = 1
    For Each item In column
        If item = value Then
            IndexWhere = item
            Exit For
        End If
        i = i + 1
    Next item
End Function

' LINE FUNCTIONS
Public Function RowAddress(index As Integer) As String
    RowAddress = Me.origin.CurrentRegion.Rows.item(index + 1).Address
End Function

Public Function RowRange(Optional index As Integer) As range
    Set RowRange = Me.origin.Worksheet.range(Me.RowAddress(index))
End Function

Public Function Row(index As Integer) As Collection
    Dim dataRange As range, headerRange As range
    Set dataRange = Me.RowRange(index)
    Set headerRange = Me.headerRange()
    
    Dim data As New Collection, item As Variant, i As Integer
    i = 1
    For Each item In dataRange.Cells
        Call data.Add(item.value, CStr(headerRange.Cells(1, i)))
        i = i + 1
    Next item
        
    Set Row = data
End Function

Public Function LastRowRange() As range
    Dim corr As Integer

    Set LastRowRange = Me.RowRange(Me.RowsCount() - 1)
End Function

Public Function LastRow() As Collection
    Dim dataRange As range
    Set dataRange = Me.LastRowRange()
    
    Dim data As New Collection, item As Variant
    For Each item In dataRange.Cells
        Call data.Add(item.value)
    Next item
    
    Set LastRow = data
End Function

Public Sub AddRow(rowData As Collection, Optional index As Integer)
    Dim last As Integer, item As Variant, blank As range, i As Integer
    
    If index = 0 Then
        last = Me.RowsCount()
        Set blank = Me.RowRange(last)
        
        For i = 1 To rowData.Count
            blank.Cells(1, i).value = rowData.item(i)
        Next i
    Else
        Dim cutRegion As range
        Set cutRegion = ExpandDown(Me.RowRange(index))
        
        Call cutRegion.Copy(Me.RowRange(index))
        
        Set blank = Me.RowRange(index - 2) 'this 2 is magic. don't know why its needed
        
        For i = 1 To rowData.Count
            blank.Cells(1, i).value = rowData.item(i)
        Next i
    End If
End Sub

Public Sub EditRow(rowData As Collection, index As Integer)
    Dim editLine As range
    Set editLine = Me.RowRange(index) 'this 2 is magic. don't know why its needed
        
    For i = 1 To rowData.Count
        editLine.Cells(1, i).value = rowData.item(i)
    Next i
End Sub

' COLUMN FUNCTIONS
Public Function ColumnAddress(index As Integer) As String
    ColumnAddress = Me.origin.CurrentRegion.columns.item(index).Address
End Function

Public Function ColumnRange(index As Integer) As range
    Set ColumnRange = Me.origin.Worksheet.range(Me.ColumnAddress(index))
End Function

Public Function column(index As Integer) As Collection
    Dim dataRange As range
    Set dataRange = Me.ColumnRange(index)
    
    Dim data As New Collection, item As Variant
    For Each item In dataRange.Cells
        Call data.Add(item.value)
    Next item
    
    Call data.remove(1)
    
    Set column = data
End Function

Public Function ColumnByName(name As String) As Collection
    Dim dataRange As range
    Set dataRange = Me.ColumnRangeByName(name)
    
    Dim data As New Collection, item As Variant
    For Each item In dataRange.Cells
        Call data.Add(item.value)
    Next item
    
    Call data.remove(1)
    
    Set ColumnByName = data
End Function

Public Function ColumnRangeByName(name As String) As range
    Dim item As Variant, i As Integer
    i = 1
    For Each item In Me.header()
        If name = item Then
            Set ColumnRangeByName = Me.ColumnRange(i)
            GoTo Error
        End If
        i = i + 1
    Next item
    
    Call MsgBox("Column with name " & name & " not found")
Error:
    Exit Function
End Function

Public Sub AddColumn(header As String, colData As Collection, Optional index As Integer)
    Dim last As Integer, item As Variant, blank As range, i As Integer
    
    If index = 0 Then
        last = Me.ColumnsCount()
        Set blank = Me.ColumnRange(last + 1)
        
        
        blank.Cells(1, 1).value = header
        blank.Worksheet.range(blank.Cells(1, 1).Address).Font.Bold = True
        For i = 1 To colData.Count
            blank.Cells(i + 1, 1).value = colData.item(i)
        Next i
    Else
        Dim cutRegion As range
        Set cutRegion = ExpandRight(Me.ColumnRange(index))
        
        Call cutRegion.Copy(Me.ColumnRange(index + 1))
        
        Set blank = Me.ColumnRange(index)
        
        
        blank.Cells(1, 1).value = header
        blank.Worksheet.range(blank.Cells(1, 1).Address).Font.Bold = True
        For i = 1 To colData.Count
            blank.Cells(i + 1, 1).value = colData.item(i)
        Next i
    End If
End Sub

Public Sub DeleteColumn(index As Integer)
    Dim cutRegion As range
    Set cutRegion = ExpandRight(Me.ColumnRange(index + 1))
    
    Call cutRegion.Copy(Me.ColumnRange(index))
    
    Me.ColumnRange(Me.ColumnsCount()).Clear
End Sub

Public Sub DeleteColumnByName(name As String)
    Dim item As Variant, i As Integer
    i = 1
    For Each item In Me.header()
        If name = item Then
            Call Me.DeleteColumn(i)
            GoTo Sair
        End If
        i = i + 1
    Next item
Sair:
    Exit Sub
End Sub


