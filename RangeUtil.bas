Attribute VB_Name = "RangeUtil"
'@Folder "D:\vba\purchase_details.VBAProject"


Option Explicit

Public Function toEnd(startNumber As Long, startColumn As String, endColumn As String, s As Worksheet)
    Dim arr()
    Dim endRow As Long
    
    endRow = s.Cells(Rows.Count, startColumn).End(xlUp).Row
    
    '特殊情况
    If endRow = startNumber Then
        Debug.Print "只有一行"
        If startColumn = endColumn Then
            Debug.Print "一行一列"
            ReDim arr(1 To 1, 1 To 1)
            arr(1, 1) = s.range(startColumn & endRow).value
            toEnd = arr: Exit Function
        End If
    End If
    
    If endRow < startNumber Then
        'MsgBox "输入的开始行号超过了结尾"
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = ""
    Else
        arr = s.range(startColumn + CStr(startNumber) + ":" + endColumn + CStr(endRow))
    End If
    toEnd = arr
End Function

Public Function toEndOneColumnThisSheet(startNumber As Long, column As String)
    toEndOneColumnThisSheet = toEnd(startNumber, column, column, ActiveSheet)
End Function

Public Function toEndOneColumn(startNumber As Long, column As String, s As Worksheet)
    toEndOneColumn = toEnd(startNumber, column, column, s)
End Function

Public Function toEndThisSheet(startNumber As Long, startColumn As String, endColumn As String)
    toEndThisSheet = toEnd(startNumber, startColumn, endColumn, ActiveSheet)
End Function

Public Function toEndSheetOneColumn(startNumber As Long, column As String, s As Worksheet)
    toEndSheetOnColumn = toEnd(startNumber, column, column, s)
End Function

Public Function toEndSheet(startNumber As Long, startColumn As String, endColumns As String, s As Worksheet)
    toEndSheet = toEnd(startNumber, startColumn, endColumns, s)
End Function

Public Function isFill(startNumber As Long, column As String, s As Worksheet) As Boolean
    Dim endRow As Long
    
    
    endRow = s.Cells(Rows.Count, column).End(xlUp).Row
    
    If endRow = startNumber Then isFill = True: Exit Function
    
    isFill = False
End Function

