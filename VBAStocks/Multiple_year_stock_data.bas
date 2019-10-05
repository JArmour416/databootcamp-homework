Attribute VB_Name = "Module1"
Sub VBA_Stocks()

' Loop through all worksheets
Dim ws As Worksheet

'Set an initial variable for holding ticket symbol
Dim Ticker_Name, Ticker_Increase, Ticker_Decrease, Ticker_Total As String

' Set an initial variable for holding the total stock volume, open and close value
Dim Volume_Total, Open_Value, Close_Value As Double

' Set an initial variable for holding the yearly and percent change
Dim Yearly As Long
Dim Percent As Double

' Keep track of the location for each ticker symbol in the summary table
Dim Summary_Table_Row As Integer

' Set an initial variable for the three greatest values
Dim Greatest_Increase As Integer
Dim Greatest_Decrease As Integer
Dim Greatest_Total As Double

    'For Each ws In ActiveWorkbook.Worksheets
    For Each ws In Worksheets
    
    ' Set or reset variables
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total = 0
    Summary_Table_Row = 1
    Volume_Total = 0
    
    ' Find Last Row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Create headings for the four new columns
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volumes"
        
        ' Loop through all ticker values
        For i = 2 To LastRow
        
        ' Check if we are still within the same ticker symbol, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
             ' Add one to the summary table row
             Summary_Table_Row = Summary_Table_Row + 1
            
             ' Set and print the ticker symbol
            Ticker_Name = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    
            ' Set the start open value and then select the first next ticker value
            If Open_Value = 0 Then
                Open_Value = ws.Cells(2, 3).Value
            Else
                Open_Value = ws.Cells(i + 1, 3).Value
            End If
                               
            ' Set the close value
            Close_Value = ws.Cells(i, 6).Value
                                       
            ' Add to and Print the total stock volume
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
                
            ' Set Yearly Change to equal close value less open value
            Yearly = Close_Value - Open_Value
            ws.Range("J" & Summary_Table_Row).Value = Yearly
            ws.Range("J2:J" & Summary_Table_Row).NumberFormat = "0.00000000"
                            
            ' Conditional formatting for negative and positive value
            For j = 2 To Summary_Table_Row
                If (ws.Cells(j, 10)) >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            Next j
                
            ' Set percentage change to equal the change in value divided by the oridinal value
            If Open_Value > 0 Then
                Percent = (Yearly / Open_Value)
            Else
                Percent = 0
            End If
                ws.Range("K" & Summary_Table_Row).Value = Percent
                ws.Range("K2:K" & Summary_Table_Row).NumberFormat = "0.00%"
    
            ' Reset the total stock volume
            Volume_Total = 0
    
            ' If the cell immediately following a row is the same ticker symbol...
        Else
            ' Add to the total stock volume
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
        End If
            
        Next
        
        ' Enter headers and row labels for greatest values
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 17).NumberFormat = "0000000000"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
                                        
        ' Retrieve the greatest increase of percent changed
        For j = 2 To Summary_Table_Row
            If ws.Cells(j, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(j, 11).Value
            Ticker_Increase = ws.Cells(j, 9).Value
            ws.Cells(2, 16).Value = Ticker_Increase
            ws.Cells(2, 17).Value = Greatest_Increase
        End If
        Next j
             
        ' Retrieve the greatest decrease of percent changed
        For j = 2 To Summary_Table_Row
            If ws.Cells(j, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(j, 11).Value
            Ticker_Decrease = ws.Cells(j, 9).Value
            ws.Cells(3, 16).Value = Ticker_Decrease
            ws.Cells(3, 17).Value = Greatest_Decrease
            End If
        Next j
            
        ' Retrieve the highest value and enter it with the ticker symbol
        For j = 2 To Summary_Table_Row
            If ws.Cells(j, 12).Value > Greatest_Total Then
            Greatest_Total = ws.Cells(j, 12).Value
            Ticker_Total = ws.Cells(j, 9).Value
            ws.Cells(4, 16).Value = Ticker_Total
            ws.Cells(4, 17).Value = Greatest_Total
            End If
        Next j
             
        ' Format cells to display nicely
        ws.Columns.AutoFit
        
    Next ws
    
End Sub

