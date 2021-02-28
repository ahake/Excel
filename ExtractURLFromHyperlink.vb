' Extract all hyperlinks from a spreadsheet (pastes them in the cell to the right of the hyperlink)

Sub ExtractHL()
    Dim HL As Hyperlink
    For Each HL In ActiveSheet.Hyperlinks
        HL.Range.Offset(0, 1).Value = HL.Address
    Next
End Sub

' Get the URL of a cell via function

Function GetURL(rng As Range) As String
    On Error Resume Next
    GetURL = rng.Hyperlinks(1).Address
End Function