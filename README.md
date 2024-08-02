# VBA-challenge

# Summary
The Multiple Stock Analysis VBA Script is for analyzing stock data for three different worksheets in an Excel workbook. The key calculations for each stock are Quarterly Change, Percent Change, and Total Stock Volume. The script locates the stock with the greatest percent increase, greatest percent decrease, and the greatest total volume, based on its ticker symbol.

# Features
-Ticker: Symbol of stock
-Quarterly Change: the difference of the open and close price for that quarter
-Percent Change: the change from the open price to the close price for that quarter using a percentage
-Total Stock Volume: the total volume of stock for that quarter that's been traded

-Greatest % Increase: the stock with the highest % increase
-Greatest % Decrease: the stock with highest % decrease
Greatest Total Volume: the stock with the greatest total trading volume

-the results are shown in columns 'I' through 'Q' on each worksheet
-the greatest values summary is shown in columns 'P' and 'Q'

# Formatting
Conditional formatting: positive changes are highlighted in green and negative changes are highlighted in red.

# How to use

To prepare, make sure the excel workbook has multiple stock data worksheets--each sheet should have columns for 'ticker', 'date', 'open', 'high', 'low', 'close' and 'volume'.

When running, open the VBA editor, copy and paste the VBA script into the "This Workbook" module for the workbook, and run the macro. After running the script, the results should be displayed in the appropriate columns on every worksheet. 

# Error Handling in Script
The script contains error handling to help with any unexpected problems that may occur during running the macro. If an error happens, a message box will appear with the error details.

# Notes/References 

1. Error code for troubleshooting

"  On Error GoTo ErrorHandler

    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Analysis complete!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in line " & Erl, vbCritical
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub"

Logan, F., (2024, July 31). Tutoring Zoom Session, Ln 2, 129-138. 

2. For loop troubleshooting/typo fix

"                ElseIf .Cells(x + 1, 1).Value <> .Cells(x, 1).Value Then
                    EndRow = x
                Else
                    TotalStockVolume = TotalStockVolume + .Cells(x, 7).Value
                    GoTo NextIteration"
                    
EdX. (2024). Xpert Learning Assistant (July 30 Version). [Large Language Model]. Ln 56-60.
