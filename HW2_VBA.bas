Attribute VB_Name = "Module1"

Public Sub run_ticker_analyzer()
    
    'start_time = Now()
    
    ' Run the ticker_analyzer for each sheet.
    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        Worksheets(i).Select
        ticker_analyzer
    Next i
    
    'end_time = Now
    'Text = " " & start_time & " to " & end_time
    'MsgBox (Text)
End Sub


Public Sub ticker_analyzer()
    ' since the ticker data is sorted
    ' Now row 2 is the current ticker start.
    ' To look for the last row of the current ticker
    ' The last row of current ticker is located just before ticker symbol change
    ' That is, the change feature of change is:
        'a row with current ticker symbol, but next row with different ticker symbol
    ' Use binary search to look for the last row of current ticker
    
    
    ' Create a variable to refer the row. initiate it as 2, the first row of ticker data
    Dim row_index As Long
    row_index = 2
    
    
    ' Find the last row. Assign the x coordinate to variable last_row
    Dim last_row_of_data As Long
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    last_row_of_data = ActiveCell.Row
    
    ' Find the last column in the worksheet. Assign the y coordinate to the variable last_col
    Dim last_col As Integer
    last_col = Cells(2, Columns.Count).End(xlToLeft).Column
   
    ' Set up the header of summary section
    Cells(1, last_col + 2).Value = "Ticker"
    Cells(1, last_col + 3).Value = "Yearly Change"
    Cells(1, last_col + 4).Value = "Percent Change"
    Cells(1, last_col + 5).Value = "Total Stock Volume"
    
    Cells(2, last_col + 8).Value = "Greatest % Increase"
    Cells(3, last_col + 8).Value = "Greatest % Decrease"
    Cells(4, last_col + 8).Value = "Greatest Total Volume"
    Cells(1, last_col + 9).Value = "Ticker"
    Cells(1, last_col + 10).Value = "Value"

    
    ' Create a variable "find_the_row" to store if the end row of current ticker is found or not
    Dim find_ticker_change As Boolean
    find_ticker_change = False
    
    Dim start_row As Long
    Dim end_row As Long
    Dim middle As Long
    Dim current_ticker_start_row As Long
    Dim current_ticker_last_row As Long
    Dim openning_price As Double
    Dim closing_price As Double
    Dim current_ticker As String
    Dim next_ticker As String
    Dim greatest_percent_increase As Double
    Dim ticker_greatest_percent_increase As String
    Dim greatest_percent_decrease As Double
    Dim ticker_greatest_percent_decrease As String
    Dim greatest_total_volume As Long
    Dim ticker_greatest_total_volume As String
    Dim percent_change As Double
    Dim total_volume As Long
    Dim summary_row As Integer
    Dim next_start_row As Long
    
    
    
    
    ' Initialize values
    start_row = 2
    end_row = last_row_of_data
    
    current_ticker = Cells(2, 1).Value
    next_ticker = Cells(2, 1).Value
    openning_price = Cells(2, 3).Value
    current_ticker_start_row = 2
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
    summary_row = 2
    
    ' Use binary search algorithm to search ticker change
    Do While start_row <= end_row
        middle = Fix((start_row + end_row) / 2)
        If (Cells(middle, 1).Value = current_ticker) Then
            If (Cells(middle + 1, 1).Value = current_ticker) Then
                start_row = middle + 1
            Else
                current_ticker_last_row = middle
                next_ticker = Cells(middle + 1, 1).Value
                next_start_row = middle + 1
                find_ticker_change = True
            End If
        Else
            If (Cells(middle - 1, 1).Value = current_ticker) Then
                current_ticker_last_row = middle - 1
                next_ticker = Cells(middle, 1).Value
                next_start_row = middle
                find_ticker_change = True
            Else
                end_row = middle - 1
            End If
        End If
        
        ' Write down the summary of the current ticker.
        If (find_ticker_change) Then
                ' Write down the summary for current ticker
                closing_price = Cells(current_ticker_last_row, 6).Value
                Cells(summary_row, last_col + 2).Value = current_ticker
                Cells(summary_row, last_col + 3).Value = closing_price - openning_price
                
                ' Check if openning price equals to zero.
                If (openning_price = 0) Then
                    Cells(summary_row, last_col + 4).Value = "Bad data"
                    Cells(summary_row, last_col + 4).Interior.ColorIndex = 7
                Else
                    percent_change = (closing_price - openning_price) / openning_price
                    Cells(summary_row, last_col + 4).Value = percent_change
                    Cells(summary_row, last_col + 4).NumberFormat = "0.00%"
                End If
                
                ' Highlight the positive value in green and negtive value in red
                If (Cells(summary_row, last_col + 3).Value >= 0) Then
                    Cells(summary_row, last_col + 3).Interior.ColorIndex = 4
                Else
                    Cells(summary_row, last_col + 3).Interior.ColorIndex = 3
                End If
                
                ' Find the greatest value and store them into variables
                If (percent_change > greatest_percent_increase) Then
                    greatest_percent_increase = percent_change
                    ticker_greatest_percent_increase = current_ticker
                End If
                If (percent_change < greatest_percent_decrease) Then
                    greatest_percent_decrease = percent_change
                    ticker_greatest_percent_decrease = current_ticker
                End If
                
                ' Use excel formula to calculate total stock volume
                Cells(summary_row, last_col + 5).Formula = "=SUM(G" & current_ticker_start_row & ":" & "G" & current_ticker_last_row & ")"
                If (Cells(summary_row, last_col + 5).Value > greatest_total_value) Then
                    greatest_total_value = Cells(summary_row, last_col + 5).Value
                    ticker_greatest_total_value = current_ticker
                End If
                
                ' Set the value for next loop
                current_ticker = next_ticker
                start_row = next_start_row
                current_ticker_start_row = start_row
                end_row = last_row_of_data
                openning_price = Cells(start_row, 3).Value
                find_ticker_change = False
                'middle = (start_row + end_row) / 2
                summary_row = summary_row + 1
        End If
              
    Loop
    
    ' Write down the summary of greatest increase/decrease/total volume
    
    Cells(2, last_col + 9).Value = ticker_greatest_percent_increase
    Cells(2, last_col + 10).Value = greatest_percent_increase
    Cells(2, last_col + 10).NumberFormat = "0.00%"
    
    Cells(3, last_col + 9).Value = ticker_greatest_percent_decrease
    Cells(3, last_col + 10).Value = greatest_percent_decrease
    Cells(3, last_col + 10).NumberFormat = "0.00%"
    
    Cells(4, last_col + 9).Value = ticker_greatest_total_value
    Cells(4, last_col + 10).Value = greatest_total_value
    
    ' Adjust the columns I:Q to be autofit.
    Columns("I:Q").AutoFit
    Cells(4, last_col + 10).Select
    
End Sub

Sub timer()
 MsgBox (format(Now, "HH:MM:SS")
End Sub


