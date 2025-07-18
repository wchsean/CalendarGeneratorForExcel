Sub CopyWorkSheets()

Dim xNumber As Integer
Dim xWsName As String
Dim startday As String
Dim not_open_day As New Collection
Dim check As String

On Error Resume Next
xTitleId = "KutoolsforExcel"
xWsName = Application.InputBox("Copy worksheet name", xTitleId, , Type:=2)


input_startday = InputBox("which day to start", "which day to start", "yyyy/mm/dd")
input_endday = InputBox("which day to end", "which day to end", "yyyy/mm/dd")
startday = CDbl(CDate(input_startday))
endday = CDbl(CDate(input_endday))
number_of_day = startaday + endday - startday
xNumber = number_of_day + 1


Do While Not check = "False"
    Dim user_input As Integer
    
    user_input = InputBox("which day need be hidden", "which day need be hidden", "1-7 1=monday 7=sunday enter 0 to exit")

    If user_input < 1 Or user_input > 7 Then
    
    check = "False"
    
    Else
        not_open_day.Add (user_input)

    
    End If
    
Loop

For i = 1 To xNumber
    Dim day As String
    Dim textday As String

    day = startday
    
    Application.ActiveWorkbook.Sheets(xWsName).Copy _
    before:=Application.ActiveWorkbook.Sheets(xWsName)
    textday = WorksheetFunction.Text(day, "dddd,  d  mmmm  yyyy")
    
    Range("b1").Value = textday
    ActiveSheet.Name = WorksheetFunction.Text(day, "d ddd mmm")
    Range("e2").Value = WorksheetFunction.Text(day, "yyyy/m/") + "1"
    startday = startday + 1
    
    If Weekday(day, 2) = 6 Or Weekday(day, 2) = 7 Then
        Debug.Print Weekday(day)
        ActiveSheet.Select
        With ActiveSheet.Tab
        .Color = 16751103
        .TintAndShade = 0
        End With
     End If
     For x = 1 To not_open_day.Count
        If Weekday(day, 2) = not_open_day(x) Then
            Debug.Print Weekday(day)
            ActiveSheet.Visible = False
        End If
    Next
ActiveSheet.Range("a1").Select
Next
End Sub



