Sub removeEmptyCellsInColumn()
' remove all empty cells in a column; transform sub to take in a cell # as parameter to be called in
    Dim NextFilled As Integer
    itStartRow = 1
    itEndRow = 1
    constCol = 7
    While (itEndRow <> -1)
        'itEndRow = Range(Cells(itStartRow, constCol)).EntireColumn.Find(What:="?*", After:=itStartRow, LookIn:=xlValues).Row
        itStartRow = Cells(itStartRow, constCol).EntireColumn.Find(What:="", After:=Cells(itStartRow, constCol), LookIn:=xlValues).Row
        itEndRow = Cells(itStartRow, constCol).EntireColumn.Find(What:="?*", After:=Cells(itStartRow, constCol), LookIn:=xlValues).Row
        If itEndRow <= itStartRow Then
            itEndRow = -1
        Else
            Range(Cells(itStartRow, constCol), Cells(itEndRow - 1, constCol)).Delete Shift:=xlUp
        End If
    Wend
End Sub
