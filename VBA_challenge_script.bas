Attribute VB_Name = "Module1"
Option Explicit

'Starting with code that will call the stocks code on all worksheets

Sub allsheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call stocks
        ws.Columns.AutoFit
    Next ws
    Application.ScreenUpdating = True
    
    MsgBox ("All years analyzed")
    
End Sub
Sub stocks()

    Dim i As Long
    Dim ticker As String
    Dim last As Long
    Dim counter As Long
    Dim opn As Double   'opening price for a ticker
    Dim cls As Double   'closing price for a ticker
    Dim vol As Double   'volume
    Dim g_lst As Long   'challenge question
    Dim g_chg As Double 'challenge question
    Dim high As Double  'challenge question
    Dim low As Double   'challenge question
    Dim g_vol As Double 'challenge question
    Dim tsv As Double   'challenge question
    
    'counting the number of rows filled with data in the column to determine iteration range
    last = Cells(Rows.Count, 1).End(xlUp).Row 'Application.CountA(Range("A:A"))
    
    'starting the counter on 2, as this will be the row we will start populating
    counter = 2

    'filling cells with new labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Beginning iteration through all tickers/clmn A until the last row
    For i = 2 To last

        'For each iteration, the ticker value can be extracted from the cell in first clmn
        ticker = Cells(i, 1).Value

        'If the current ticker is not the same as the previous ticker
        'And it is not the same as the next ticker
        'Then extract the ticker symbol, assign values to open,close,volume
        'Calculate change and percent change
        'Add a counter (move to the next row for next entry)
        'This will only run if there is only 1 data point for a stock
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(counter, 9).Value = ticker
            opn = Cells(i, 3).Value
            cls = Cells(i, 6).Value
            Cells(counter, 10).Value = cls - opn
            Cells(counter, 10).NumberFormat = "0.00"
            Cells(counter, 11).Value = (cls - opn) / opn
            Cells(counter, 11).NumberFormat = "0.00%"
            Cells(counter, 12).Value = vol
            counter = counter + 1

        'Otherwise, if the current ticker is only different from the previous ticker
        'Then extract the ticker symbol
        'And assign values to open and volume for later calculations
        
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Cells(counter, 9).Value = ticker
            opn = Cells(i, 3).Value
            vol = Cells(i, 7).Value

        'Else if the current ticker is only different from the next ticker
        'And open value assigned earlier was zero
        'Then extract close value
        'Calculate change and volume change
        'Percent change will be set to zero as change cannot be divided by open
        'Add a counter

        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value And opn = 0 Then
            cls = Cells(i, 6).Value
            Cells(counter, 10).Value = cls - opn
            Cells(counter, 10).NumberFormat = "0.00"
            Cells(counter, 11).Value = 0
            Cells(counter, 11).NumberFormat = "0.00%"
            vol = vol + Cells(i, 7).Value
            Cells(counter, 12).Value = vol
            counter = counter + 1
        
        'Else if the current ticker is not equal to the next ticker
        'Then assign close value
        'Calculate change, percent change and volume change
        'Conditional formatting applied
        'Add a counter
        
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            cls = Cells(i, 6).Value
            Cells(counter, 10).Value = cls - opn
            Cells(counter, 10).NumberFormat = "0.00"
            Cells(counter, 11).Value = (cls - opn) / opn
            Cells(counter, 11).NumberFormat = "0.00%"
            vol = vol + Cells(i, 7).Value
            Cells(counter, 12).Value = vol
            
            If Cells(counter, 10).Value > 0 Then
                Cells(counter, 10).Interior.ColorIndex = 4
            Else
                Cells(counter, 10).Interior.ColorIndex = 3
            End If
            
            counter = counter + 1

        'In all other situations (when the ticker is the same)
        'Keep adding to the volume assigned in the beginning
        
        Else
            vol = vol + Cells(i, 7).Value

        End If

    Next i
    
    'counting the number of rows filled with data in the column
    'To determine iteration range minus 1 so i+1 can be used
    g_lst = Cells(Rows.Count, 1).End(xlUp).Row 'Application.CountA(Range("K:K"))
    high = Cells(2, 11).Value   'initializing variable with first value
    low = Cells(2, 11).Value    'initializing variable with first value
    g_vol = Cells(2, 12).Value  'initializing variable with first value
    
    'filling cells with new labels
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Beginning iteration through all tickers/clmn A until the last row
    For i = 2 To g_lst

        'For each iteration, the initial values can be extracted from the following cells
        ticker = Cells(i, 9).Value
        g_chg = Cells(i, 11).Value  'g_chg is just defining the selected
        tsv = Cells(i, 12).Value    'total stock volume

        'If selected cell value is greater than high
        'Then extract the ticker symbol
        'Reassign high the value and extract it
        
        If g_chg > high Then
            Cells(2, 16).Value = ticker
            high = g_chg
            Cells(2, 17).Value = high
            Cells(2, 17).NumberFormat = "0.00%"

        'Else if selected cell value is less than low
        'Then extract the ticker symbol
        'Reassign low the value and extract it
        
        ElseIf g_chg < low Then
            Cells(3, 16).Value = ticker
            low = g_chg
            Cells(3, 17).Value = low
            Cells(3, 17).NumberFormat = "0.00%"

        End If

        'If selected cell value is greater than the assigned volume
        'Then extract ticker symbol
        'Reassign g_vol the new value and extract it
        
        If tsv > g_vol Then
            Cells(4, 16).Value = ticker
            g_vol = tsv
            Cells(4, 17).Value = g_vol
            Cells(4, 17).NumberFormat = "0.0000E+00"
            
        End If

    Next i

End Sub
