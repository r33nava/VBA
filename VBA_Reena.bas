Attribute VB_Name = "Module1"
Sub A_Click()

'This module calculates stocks from spreadsheet 'A'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume

     


End Sub
Sub B_Click()
'This module calculates stocks from spreadsheet 'B'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
Sub C_Click()
'This module calculates stocks from spreadsheet 'C'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
Sub D_Click()
'This module calculates stocks from spreadsheet 'D'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
Sub E_Click()
'This module calculates stocks from spreadsheet 'E'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
Sub F_Click()
'This module calculates stocks from spreadsheet 'A'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
Sub P_Click()
'This module calculates stocks from spreadsheet 'A'.
'It will loop through each year of stock data and grab the total amount of volume each stock had over the year.


'Variables for Symbolu

    Dim Symbol  As Double
    Dim Total_Volume As Double

 

'Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"


' Starts at Row #2 because the first row contains headers and 2nd row starts with ticker symbol.
' Total Volume starts at '0'.

    Symbol = 2
    Total_Volume = 0

' For the 1st symbol to be identified and added uder the 'Ticker' header

    Cells(Symbol, 9).Value = Cells(Symbol, 1).Value

' This allows module to determine the last row in the worksheet.

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'For the loop to run from row 2 to the last row in the spreadsheet

    For Row = 2 To LastRow


' If the ticker symbol matches what's the ticker symbol currently in row 1, then start the loop

    If Cells(Row, 1).Value = Cells(Symbol, 9) Then
    
        ' This adds up all the volume per ticker symbol and holds it in variable "Total_Volumne"
        Total_Volume = Total_Volume + Cells(Row, 7).Value

    Else

        'Add the total volumn under the 'Total Volume' column and looks to the 'Ticker' column to check the ticker symbol

        Cells(Symbol, 10).Value = Total_Volume
     
        Total_Volume = Cells(Row, 7).Value

        Symbol = Symbol + 1

        Cells(Symbol, 9).Value = Cells(Row, 1).Value

     End If

     
'End of current iteration and start of the next row
    Next Row


    Cells(Symbol, 10).Value = Total_Volume
End Sub
