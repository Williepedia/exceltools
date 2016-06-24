

Private Function autofitRange() As Range

Set autofitRange = Application.InputBox("Select the Autosize Data Range", Type:=8)

End Function

Sub quickFormat()

Dim width As Integer
Dim ws As Worksheet
Dim subschedule As String
Dim wb As Workbook
Dim side_margin As Long
Dim myform As New grf_style_form

Set wb = ActiveWorkbook

myform.Show

If myform.boolUserCancelled = True Then
    GoTo userCancelled
End If


For Each ws In wb.Worksheets:
    
    top_of_header = 0
    bottom_of_header = 0
    
     '------------------------------
     'Set Universal Font to Arial 10
     '------------------------------
     With ws.Cells.Font
         .Name = "Arial"
         .Size = 10
         .OutlineFont = False
         .Shadow = False
     End With
    
    If myform.CheckBox_RenameWS Then
        
        On Error GoTo error_setSheetName:
        ws.Name = setSheetName(ws)
         
        On Error GoTo 0
    End If
    
    '-------------------------------------------------------
    'Autofit columns + a buffer for when numbers are cut off
    '-------------------------------------------------------
    
    
    Maximum_Column_Width = 100
    Column_Buffer = 2
    If myform.CheckBox_AutosizeColumns = True Then
        autofitRange().SpecialCells(xlCellTypeVisible).Columns.AutoFit
    End If
    
    'Resets End of Sheet (why this only works on the ActiveSheet, I have no idea)
    ws.Activate
    a = activesheet.UsedRange.Columns.Count
    b = activesheet.UsedRange.Rows.Count
     
    'determine ws width
    width = sheetWidth(ws)
         
    '-----------------------------------
    'set to landscape if ws is too wide
    '-----------------------------------
    maxPortraitWidth = 800
    If width > maxPortraitWidth Then
        ws.PageSetup.Orientation = xlLandscape
    ElseIf width < maxPortraitWidth Then
        ws.PageSetup.Orientation = xlPortrait
    End If
    
    '-----------
    'Set Margins
    '-----------
    If ws.PageSetup.Orientation = xlLandscape Then
        side_margin = 0.25
    Else
        side_margin = 0.75
    End If
    
    '-------------------
    'Initial Page Setup
    '-------------------
     With ws.PageSetup
         .LeftHeader = ""
         .CenterHeader = ""
         .RightHeader = "&""Arial,Regular""&10&A"
         .LeftFooter = ""
         .CenterFooter = ""
         .RightFooter = ""
         .LeftMargin = Application.InchesToPoints(side_margin)
         .RightMargin = Application.InchesToPoints(side_margin)
         .TopMargin = Application.InchesToPoints(1)
         .BottomMargin = Application.InchesToPoints(1)
         .HeaderMargin = Application.InchesToPoints(0.5)
         .FooterMargin = Application.InchesToPoints(0.5)
         .PrintHeadings = False
         .PrintGridlines = False
         .PrintComments = xlPrintNoComments
         .PrintQuality = 600
         .CenterHorizontally = True
         .CenterVertically = False
         .Draft = False
         .PaperSize = xlPaperLetter
         .FirstPageNumber = xlAutomatic
         .Order = xlDownThenOver
         .BlackAndWhite = False
         .Zoom = False
         .FitToPagesWide = 1
         .FitToPagesTall = False
         .PrintErrors = xlPrintErrorsDisplayed
         .OddAndEvenPagesHeaderFooter = False
         .DifferentFirstPageHeaderFooter = False
         .ScaleWithDocHeaderFooter = True
         .AlignMarginsHeaderFooter = True
         .EvenPage.LeftHeader.Text = ""
         .EvenPage.CenterHeader.Text = ""
         .EvenPage.RightHeader.Text = ""
         .EvenPage.LeftFooter.Text = ""
         .EvenPage.CenterFooter.Text = ""
         .EvenPage.RightFooter.Text = ""
         .FirstPage.LeftHeader.Text = ""
         .FirstPage.CenterHeader.Text = ""
         .FirstPage.RightHeader.Text = ""
         .FirstPage.LeftFooter.Text = ""
         .FirstPage.CenterFooter.Text = ""
         .FirstPage.RightFooter.Text = ""
     End With
     

     
     '--------------
     'Remove Filters
     '--------------
     ws.AutoFilterMode = False
     
     '------------------------------------------------------
     'Add page numbers to footer if Worksheet Exceeds 1 Page
     '------------------------------------------------------
    If myform.CheckBox_RepeatingHeaders And (ws.PageSetup.Pages.Count) > 1 Then
        ws.Activate
        ws.Cells(1, 1).Activate
        headerRng = Application.InputBox("Select Header Rows", Type:=8)
        If headerRng = vbNullString Then
            GoTo userCancelled
        End If
        Set Rng = headerRng
        top_of_header = (Rng.Row)
        bottom_of_header = (Rng.Rows.Count + Rng.Row - 1)
    End If

     If (ws.PageSetup.Pages.Count) > 1 Then
         Debug.Print ("Adding Page #'s... to Worksheet " & ws.Index)
         ws.PageSetup.CenterFooter = "&""Arial,Regular""Page &P of &N"
         ws.PageSetup.PrintTitleRows = "$" & top_of_header & ":$" & bottom_of_header
     End If
 Next ws
 

 Exit Sub
    
'--------------
'ERROR HANDLING
'--------------
error_setSheetName:
    newname = InputBox("Error Renaming Sheet. Enter Sheet Name", "Rename Sheet", "Sheet " & ws.Index)
    If newname = vbNullString Then
        GoTo userCancelled
    End If
    ws.Name = newname
    Resume Next

userCancelled:
    MsgBox ("User Cancelled!")
    Exit Sub

End Sub


Private Function sheetWidth(ws As Worksheet) As Integer

Dim width As Integer

For Each cell In ws.UsedRange.Rows(1)
    width = width + cell.width
Next cell

sheetWidth = width

End Function

Private Sub resetUsedRange()
    a = activesheet.UsedRange.Columns.Count
    b = activesheet.UsedRange.Rows.Count
End Sub

Private Function setSheetName(ws As Worksheet) As String

        'Set lettering scheme for subschedules - does not work for wb's exceeding 27 sheets
        subschedule = ""
        If ws.Index <> 1 Then
            subschedule = Chr(ws.Index + 63)
        End If
        
        'Name worksheets according to filename, stopping after the second space + subschedule reference
        setSheetName = Left(ws.Parent.Name, InStr(InStr(1, ws.Parent.Name, " ") + 1, ws.Parent.Name, " ") - 1) & subschedule
        
End Function


Private Sub MeasureSelection()
'PURPOSE: Provide Height and Width of Currently Selected Cell Range
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim cell As Range
Dim width As Long
Dim Height As Long

'Measure Selection Height
  For Each cell In Selection.Cells.Columns(1)
    Height = Height + cell.Height
  Next cell

'Measure Selection Width
  For Each cell In Selection.Cells.Rows(1)
    width = width + cell.width
  Next cell

'Report Results
  MsgBox "Height:  " & Height & "px" & vbCr & "Width:   " _
   & width & "px", , "Dimensions"

End Sub

