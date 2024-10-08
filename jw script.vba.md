
Sub vbachallenge()


' set to run for each worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    
    'Set fields - always need to do first
    
    Dim total As Single
    Dim J As Long
    Dim change As Single
    Dim K As Integer
    Dim Start As Long
    Dim LastRow As Long
    Dim PercentChange As Single
    Dim DailyChange As Single
    Dim AverageChange As Double
    
'Set Locations for titles of new columns and fields - max,greatest change, etc.; add range/value first and then output

Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Set initial values for iteraitons

K = 0
total = 0
change = 0
Start = 2

LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For J = 2 To LastRow
' in IF statements, you want to find your first non-zero, then exit For, if its a 0 it will be handled differently
' look for change in ticker value, look for zero values

        If ws.Cells(J + 1, 1).Value <> ws.Cells(J, 1).Value Then

            total = total + Cells(J, 7).Value
            
            If total = 0 Then
           ws.Range("I" & 2 + K).Value = Cells(J, 1).Value
            ws.Range("J" & 2 + K).Value = 0
            ws.Range("K" & 2 + K).Value = "%" & 0
            ws.Range("L" & 2 + K).Value = 0
            
        Else
        
        If ws.Cells(Start, 3) = 0 Then
        For find_value = Start To J
            If ws.Cells(find_value, 3).Value <> 0 Then
                Start = find_value
            Exit For
        End If
    Next find_value
End If

' finally calculate the change, in 6th column


change = (ws.Cells(J, 6) - ws.Cells(Start, 3))
PercentChange = change / ws.Cells(Start, 3)

Start = J + 1

' Print Results with formatting, first prints ticker value, then change with formatting, then percent cahange with formatting

ws.Range("I" & 2 + K).Value = ws.Cells(J, 1).Value
ws.Range("J" & 2 + K).Value = change
ws.Range("J" & 2 + K).NumberFormat = "0.00"
ws.Range("K" & 2 + K).Value = PercentChange
ws.Range("K" & 2 + K).NumberFormat = "0.00%"
ws.Range("L" & 2 + K).Value = total

' adding color formatting to the above results
' inserting a case statement: Select Case Change ((^^ referring to the calculate statement you made above)
'Case is >0 range(....).Interior.ColorIndex = 4// end with end select

Select Case change
    Case Is < 0
    ws.Range("J" & 2 + K).Interior.ColorIndex = 3
        Case Is > 0
    ws.Range("J" & 2 + K).Interior.ColorIndex = 4
    Case Else
    ws.Range("J" & 2 + K).Interior.ColorIndex = 2
End Select


Select Case PercentChange
    Case Is < 0
    ws.Range("K" & 2 + K).Font.ColorIndex = 3
        Case Is > 0
    ws.Range("K" & 2 + K).Font.ColorIndex = 4
    Case Else
    ws.Range("K" & 2 + K).Font.ColorIndex = 1
End Select

'end if to reset for the new stock ticker below

End If

    total = 0
    change = 0
    K = K + 1
    
Else
        total = total + ws.Cells(J, 7).Value

End If

Next J

' max and other values in VBA using MATCH function

        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))


        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)

        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)



' need to end with a spinner on top of everything that applies to all worksheets
    Next ws

End Sub

