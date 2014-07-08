Option Explicit

Sub onlyDigits_Descriptor()
    Dim argDescription(1) As String
    argDescription(0) = "A numeric argument"
   
    Application.MacroOptions Macro:="onlyDigits", _
        Description:="Recieved a single argument and returns its half value", _
        Category:="My Category", _
        StatusBar:="Returns the half value of a numeric argument", _
        ArgumentDescriptions:=argDescription
End Sub

Function onlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim I As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For I = 1 To Len(s)
        If Mid(s, I, 1) >= "0" And Mid(s, I, 1) <= "9" Then
            retval = retval + Mid(s, I, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function

Function leadingZeros(target As String, length As Integer)

    While Len(target) < length
        target = "0" & target
    Wend
    
    leadingZeros = target

End Function

Function dhLastDayInQuarter(Optional dtmDate As Date = 0) As Date
    ' Returns the last day in the quarter specified
    ' by the date in dtmDate.
    Const dhcMonthsInQuarter As Integer = 3
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInQuarter = DateSerial( _
     year(dtmDate), _
     Int((month(dtmDate) - 1) / dhcMonthsInQuarter) _
      * dhcMonthsInQuarter + (dhcMonthsInQuarter + 1), _
     0)
End Function

Function period2date(target As String)

Dim year As Integer
Dim month As Integer
Dim day As Integer

year = Left(target, 4)


Select Case Right(target, 2)
    Case "Q1"
        month = 3
        day = 31
    Case "Q2"
        month = 6
        day = 30
    Case "Q3"
        month = 9
        day = 30
    Case "Q4"
        month = 12
        day = 31
        
End Select
period2date = DateSerial(year, month, day)
End Function


Function csvRange(myrange As Range)
    Dim csvRangeOutput
    For Each entry In myrange
        If Not IsEmpty(entry.Value) Then
            csvRangeOutput = csvRangeOutput & entry.Value & ", "
        End If
    Next
    csvRange = Left(csvRangeOutput, Len(csvRangeOutput) - 1)
End Function

Function reverseCSV(target As String, delimiter As String)

    Dim rCSV As Variant
    
    rCSV = Split(target, delimiter)
    
    
    reverseCSV = rCSV


End Function
Private Function FindLastRow() As Long
Dim LastRow As Long
If WorksheetFunction.CountA(Cells) > 0 Then
    'Search for any entry, by searching backwards by Rows.
    LastRow = Cells.find(What:="*", After:=[A1], _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).Row
    FindLastRow = LastRow
End If
End Function

Private Function FindLastColumn() As Integer
Dim lastcolumn As Integer
If WorksheetFunction.CountA(Cells) > 0 Then
    'Search for any entry, by searching backwards by Columns.
    lastcolumn = Cells.find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious).Column
    FindLastColumn = lastcolumn
End If
End Function

