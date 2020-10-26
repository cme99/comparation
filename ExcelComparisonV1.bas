Sub findWrongInfoBetween2Workbook()
    Dim my_Filename As Variant
    Dim my_newFile As Workbook
    Dim my_oldFile As Workbook
    Dim isMatch As Boolean

    my_Filename = Application.GetOpenFilename(FileFilter:="Excel Files, *.xl*;*.xm*")
    If my_Filename <> False Then
        Set my_oldFile = Workbooks.Open(my_Filename)
    End If
    my_Filename = Application.GetOpenFilename(FileFilter:="Excel Files, *.xl*;*.xm*")
    If my_Filename <> False Then
        Set my_newFile = Workbooks.Open(my_Filename)
    End If
    
    compareSheets my_oldFile, my_newFile
    Dim roi As Variant
    roi = InputBox("Input your interested area, e.g: A1:D10")
    compareRows my_oldFile, my_newFile, roi
End Sub
Sub compareRows(my_oldFile As Workbook, my_newFile As Workbook, roi As Variant)
    numSheetNew = my_newFile.Sheets.Count
    For c = 1 To numSheetNew
        my_newFile.Sheets(c).Activate
        newFileSheetColor = my_newFile.Sheets(c).Tab.Color
        If newFileSheetColor <> 11851260 Then
            Range(roi).Select
            RowCount = Range(roi).Rows.Count
            ColCount = Range(roi).Columns.Count
            Dim rng As Range, cell As Range
            Set rng = Range(roi)
            For Each cell In rng
                If cell <> my_oldFile.Sheets(my_newFile.Sheets(c).Name).Range(cell.Address).Value Then
                    cell.Interior.Color = RGB(252, 213, 180)
                    cell.Font.FontStyle = "Bold Italic"
                End If
            Next cell
        End If
    Next c
End Sub
Sub compareSheets(my_oldFile As Workbook, my_newFile As Workbook)
    numSheetOld = my_oldFile.Sheets.Count
    numSheetNew = my_newFile.Sheets.Count
    If numSheetOld > numSheetNew Then
        For i = 1 To numSheetOld
            isMatch = False
            For j = 1 To numSheetNew
                If my_oldFile.Sheets(i).Name = my_newFile.Sheets(j).Name Then
                    isMatch = True
                    Exit For
                End If
            Next j
            If isMatch = False Then
                For Each WS In my_newFile.Sheets
                  If my_oldFile.Sheets(i).Name = WS.Name Then
                    CheckIfSheetExists = True
                  End If
                Next WS
                If CheckIfSheetExists = 0 Then
                    my_newFile.Sheets.Add.Name = my_oldFile.Sheets(i).Name
                    my_newFile.Sheets(my_oldFile.Sheets(i).Name).Tab.Color = RGB(252, 213, 180)
                End If
            End If
        Next i
    ElseIf numSheetOld < numSheetNew Then
        For i = 1 To numSheetNew
            isMatch = False
            For j = 1 To numSheetOld
                If my_newFile.Sheets(i).Name = my_oldFile.Sheets(j).Name Then
                    isMatch = 1
                    Exit For
                End If
            Next j
            If isMatch = 0 Then
                my_newFile.Sheets(i).Tab.Color = RGB(252, 213, 180)
            End If
        Next i
    End If
End Sub


