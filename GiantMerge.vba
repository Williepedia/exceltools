Option Explicit
Const NUMBER_OF_SHEETS = 1

Public Sub GiantMerge()
    Dim externWorkbookFilepath As Variant
    Dim externWorkbook As Workbook
    Dim targetBook As Workbook
    Dim i As Long
    Dim mainLastEnd(1 To NUMBER_OF_SHEETS) As Range
    Dim mainCurEnd As Range

    Application.ScreenUpdating = False

    ' Initialise
    Set targetBook = Application.Workbooks.Add
    
    ' Correct number of sheets
    Application.DisplayAlerts = False
    If targetBook.Sheets.Count < NUMBER_OF_SHEETS Then
        targetBook.Sheets.Add Count:=NUMBER_OF_SHEETS - targetBook.Sheets.Count
    ElseIf targetBook.Sheets.Count > NUMBER_OF_SHEETS Then
        For i = targetBook.Sheets.Count To NUMBER_OF_SHEETS + 1 Step -1
            targetBook.Sheets(i).Delete
        Next i
    End If
    Application.DisplayAlerts = True

    For i = 1 To NUMBER_OF_SHEETS
        Set mainLastEnd(i) = GetTrueEnd(targetBook.Sheets(i))
    Next i


    ' Load the data
    Application.DisplayAlerts = False
    For Each externWorkbookFilepath In GetWorkbooks()
        Set externWorkbook = Application.Workbooks.Open(externWorkbookFilepath, , True)

        For i = 1 To NUMBER_OF_SHEETS

            If mainLastEnd(i).Row > 1 Then
                ' There is data in the sheet

                ' Copy new data (skip headings)
                externWorkbook.Sheets(i).Range("A2:" & GetTrueEnd(externWorkbook.Sheets(i)).Address).Copy targetBook.Sheets(i).Cells(mainLastEnd(i).Row + 1, 1)

                ' Find the end column and row
                Set mainCurEnd = GetTrueEnd(targetBook.Sheets(i))
            Else
                ' No nata in sheet yet (prob very first run)

                ' Get correct sheet name from first file we check
                targetBook.Sheets(i).Name = externWorkbook.Sheets(i).Name

                ' Copy new data (with headings)
                externWorkbook.Sheets(i).Range("A1:" & GetTrueEnd(externWorkbook.Sheets(i)).Address).Copy targetBook.Sheets(i).Cells(mainLastEnd(i).Row, 1)

                ' Find the end column and row
                Set mainCurEnd = GetTrueEnd(targetBook.Sheets(i)).Offset(, 1)

                ' Add file name heading
                targetBook.Sheets(i).Cells(1, mainCurEnd.Column).Value = "File Name"
            End If

            ' Add file name into extra column
            targetBook.Sheets(i).Range(targetBook.Sheets(i).Cells(mainLastEnd(i).Row + 1, mainCurEnd.Column), mainCurEnd).Value = externWorkbook.Name

            Set mainLastEnd(i) = mainCurEnd
        Next i

        externWorkbook.Close
        Application.DisplayAlerts = False
    Next externWorkbookFilepath
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' Returns a collection of file paths, or an empty collection if the user selects cancel
Private Function GetWorkbooks() As Collection
    Dim fileNames As Variant
    Dim xlFile As Variant

    Set GetWorkbooks = New Collection

    fileNames = Application.GetOpenFilename(Title:="Please choose the files to merge", _
                                               FileFilter:="Text Files (*.txt), *.txt, Excel Files (*.xls*), *.xls*", _
                                               MultiSelect:=True, _
                                               FilterIndex:=2)
    If TypeName(fileNames) = "Variant()" Then
        For Each xlFile In fileNames
            GetWorkbooks.Add xlFile
        Next xlFile
    End If
End Function

' Finds the true end of the table (excluding unused columns/rows and rows filled with 0's)
Private Function GetTrueEnd(ws As Worksheet) As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long

    On Error Resume Next
    lastCol = ws.UsedRange.Find("*", , , xlPart, xlByColumns, xlPrevious).Column
    lastRow = ws.UsedRange.Find("*", , , xlPart, xlByRows, xlPrevious).Row
    On Error GoTo 0

    If lastCol <> 0 And lastRow <> 0 Then

        ' look back through the last rows of the table, looking for a non-zero value
        For r = lastRow To 1 Step -1
            For c = 1 To lastCol
                If ws.Cells(r, c).Text <> "" Then
                    If ws.Cells(r, c).Text <> 0 Then
                        Set GetTrueEnd = ws.Cells(r, lastCol)
                        Exit Function
                    End If
                End If
            Next c
        Next r
    End If

    Set GetTrueEnd = ws.Cells(1, 1)
End Function






