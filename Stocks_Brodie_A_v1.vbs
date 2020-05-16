
Sub stocks1()

Application.ScreenUpdating = False

Dim sheetsarray() As String
Dim lastrow As Long
Dim lastrow2 As Long



Dim init_sheets As Integer

    init_sheets = 0
' find number of sheets
For Each ws In Worksheets

    init_sheets = init_sheets + 1

    Next ws
'put sheets name in array

Dim sheets_array() As String

    ReDim sheets_array(init_sheets)

Dim counter As Integer

    counter = 1

For Each ws In Worksheets


    sheets_array(counter) = ws.Name
    counter = counter + 1


Next ws

Dim vol As Double


  
'Sheets(sheets_array(1)).Select

For e = 1 To init_sheets

  Dim column As Integer
  column = 1
  counter = 2
Sheets(sheets_array(e)).Select

   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through rows in the column
  For i = 2 To lastrow
If i = 2 Then
Cells(i, 11).Value = Cells(2, 4).Value
End If

    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
       Cells(counter, 10).Value = Cells(i, column).Value
       Cells(counter + 1, 11).Value = Cells(i + 2, 3).Value
       Cells(counter, 12).Value = Cells(i, 6).Value
       vol = vol + Cells(i, 7).Value
      Cells(counter, 13).Value = vol
      ' Message Box the value of the current cell and value of the next cell
     ' MsgBox (Cells(i, column).Value & " and then " & Cells(i + 1, column).Value)
      counter = counter + 1
      vol = 0
      
Else

vol = vol + Cells(i, 7).Value

    End If

  Next i
  
     lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row
  
  For i = 2 To lastrow2
   Cells(i, 14).Value = Cells(i, 12).Value - Cells(i, 11).Value
   If Cells(i, 11).Value = 0 Then
   Else
   Cells(i, 15).Value = Cells(i, 14).Value / Cells(i, 11).Value
   End If
   
  Next i
  
 Columns("N:O").Select
    Application.CutCopyMode = False
    Selection.Cut
    ActiveWindow.SmallScroll Down:=-9
    Range("K1").Select
    ActiveSheet.Paste
'cell formating

    Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Columns("L:L").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
 Cells(1, 10).Value = "Ticker"
 Cells(1, 11).Value = "Yearly Change"
 Cells(1, 12).Value = "Percent Change"
 Cells(1, 13).Value = "Total Stock Volume"
   Range("J1:M1").Select
  
    Selection.FormatConditions.Delete
    Next e
    
End Sub

