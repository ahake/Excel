' The way this works is very simple. Use Column A of your spreadsheet to set a True or False value. 
' You can do this manually or via formula, 
' DeleteBlankRows() hides every row with True in column A
' ShowAllRows() just unhides all hidden rows

Sub DeleteBlankRows()
    For Each c In Range("A1:A70")
        If c.Value Then
            c.EntireRow.Hidden = True
        Else
            c.EntireRow.Hidden = False
        End If
    Next
End Sub
Sub ShowAllRows()
    Rows.EntireRow.Hidden = False
End Sub
