Sub Summary_DataTest()

'Declaring variables and arrays
Dim file_Name, max_incr_ticker, max_decr_ticker, max_vol_ticker, date_List(15000), ticker(15000) As String
Dim sheet, test_Sheet As Worksheet, wk As Workbook
Dim max_Increase, max_Decrease, max_Volume, last_Row, stock_Count As Integer
Dim ticker_Count, volume(15000) As Double
Dim open_Price(15000), close_Price(15000) As Double

'Open Source Workbook
Application.ScreenUpdating = False
Application.DisplayAlerts = False

file_Name = Range("E7")
Set wk = Workbooks.Open(file_Name, ReadOnly:=True)

'Iteration Source Workbook over all sheets
For Each sheet In wk.Worksheets

'Getting Last row number for data on active sheet
    sheet.Activate
    last_Row = Cells(Rows.Count, 1).End(xlUp).Row

'Import Dataset into Arrays from Active Sheet
        For i = 2 To last_Row

          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ticker(stock_Count) = Cells(i, 1)
                date_List(stock_Count) = Cells(i, 2)
                open_Price(stock_Count) = Cells(i - ticker_Count, 3)
                close_Price(stock_Count) = Cells(i, 6)
                volume(stock_Count) = Application.Sum(Range(Cells(i - ticker_Count, 7), Cells(i, 7)))
                ticker_Count = 0
                stock_Count = stock_Count + 1
          Else
                ticker_Count = ticker_Count + 1

          End If
        Next i
Next

'Closing  Source Workbook
ActiveWorkbook.Close False
Application.ScreenUpdating = True
Application.DisplayAlerts = True

'Creating Summary Page at Main Book
Set test_Sheet = ThisWorkbook.Worksheets.Add
test_Sheet.Name = "Summary_Analysis"
test_Sheet.Activate

Range("A1:D1500").HorizontalAlignment = xlCenter
Range("K1:K1500").HorizontalAlignment = xlCenter
Range("E1:E1500").HorizontalAlignment = xlHAlignRight
Range("D1:D1500").NumberFormat = "0.00%"
Range("E1:E1500").NumberFormat = "#,###"
Range("A1:K1").Interior.Color = RGB(204, 255, 204)
Range("A1:K1").Font.Size = 14
Range("A1:K1").Font.Bold = True

Cells(1, 1) = "Year"
Cells(1, 2) = "Ticker"
Cells(1, 3) = "Yearly_Change"
Cells(1, 4) = "Percent_Change"
Cells(1, 5) = "Total_Stock_Volume"
Range("H1") = "Challenge Output"
Range("I1") = "Ticker"
Range("J1") = "Output_Value"
Range("K1") = "Year"
Range("H2") = "Greatest % Increase"
Range("H3") = "Greatest % Decrease"
Range("H4") = "Greatest Total Volume"
Columns("B:E").AutoFit
Columns("H:J").AutoFit
Columns("A").ColumnWidth = 10
Columns("F").ColumnWidth = 2

'Extracting Data from Arrays into the Summary Table
For i = 1 To stock_Count
    Cells(i + 1, 2) = ticker(i - 1)
    Cells(i + 1, 3) = close_Price(i - 1) - open_Price(i - 1)
    Cells(i + 1, 5) = volume(i - 1)

    Cells(i + 1, 1) = Left(date_List(i - 1), 4)

    If open_Price(i - 1) <> 0 Then
            Cells(i + 1, 4) = (close_Price(i - 1) / open_Price(i - 1) - 1)
    End If

    If Cells(i + 1, 3) < 0 Then
        Cells(i + 1, 3).Interior.Color = RGB(255, 0, 0)
        Else
        Cells(i + 1, 3).Interior.Color = RGB(0, 255, 0)
    End If

Next i

'Calculating Challenges Outputs from Summary Table over all years
max_Increase = 0
max_Decrease = 0
max_Volume = 0
For i = 2 To stock_Count
    If Cells(i, 4) > max_Increase Then
        max_Increase = Cells(i, 3)
        max_incr_ticker = Cells(i, 2)
    End If
    If Cells(i, 4) < max_Decrease Then
        max_Decrease = Cells(i, 3)
        max_decr_ticker = Cells(i, 2)
    End If
    If Cells(i, 5) > max_Volume Then
        max_Volume = Cells(i, 5)
        max_vol_ticker = Cells(i, 2)
    End If
Next i

'Displaying the Results into the Challenges Outputs Table All years
Range("J2:J19").NumberFormat = "0.00%"
Range("J2") = max_Increase
Range("I2") = max_incr_ticker

Range("J3") = max_Decrease
Range("I3") = max_decr_ticker

Range("J4").NumberFormat = "#,###"
Range("J4") = max_Volume
Range("I4") = max_vol_ticker
Range("K2:K4") = "all years"

'Calculating Challenges Outputs from Summary Table over each year
max_Increase = 0
max_Decrease = 0
max_Volume = 0
ticker_Count = 0
For i = 2 To stock_Count + 1
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_Count = ticker_Count + 1

        Cells(6 * ticker_Count, 8).Interior.Color = RGB(204, 255, 204)
        Cells(6 * ticker_Count, 8).Font.Size = 14
        Cells(6 * ticker_Count, 8).Font.Bold = True
        Cells(6 * ticker_Count, 8) = Cells(i, 1)

        Cells(6 * ticker_Count + 1, 8) = "Greatest % Increase"
        Cells(6 * ticker_Count + 1, 9) = max_incr_ticker
        Cells(6 * ticker_Count + 1, 10) = max_Increase
        Cells(6 * ticker_Count + 1, 11) = Cells(i, 1)
        Cells(6 * ticker_Count + 2, 8) = "Greatest % Decrease"
        Cells(6 * ticker_Count + 2, 9) = max_decr_ticker
        Cells(6 * ticker_Count + 2, 10) = max_Decrease
        Cells(6 * ticker_Count + 2, 11) = Cells(i, 1)
        Cells(6 * ticker_Count + 3, 10).NumberFormat = "#,###"
        Cells(6 * ticker_Count + 3, 8) = "Greatest Total Volume"
        Cells(6 * ticker_Count + 3, 9) = max_vol_ticker
        Cells(6 * ticker_Count + 3, 10) = max_Volume
        Cells(6 * ticker_Count + 3, 11) = Cells(i, 1)
        max_Increase = 0
        max_Decrease = 0
        max_Volume = 0
   Else
        If Cells(i, 4) > max_Increase Then
            max_Increase = Cells(i, 3)
            max_incr_ticker = Cells(i, 2)
        End If
        If Cells(i, 4) < max_Decrease Then
            max_Decrease = Cells(i, 3)
            max_decr_ticker = Cells(i, 2)
        End If
        If Cells(i, 5) > max_Volume Then
            max_Volume = Cells(i, 5)
            max_vol_ticker = Cells(i, 2)
        End If
    End If

Next i

End Sub
