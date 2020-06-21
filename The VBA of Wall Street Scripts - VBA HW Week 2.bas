Attribute VB_Name = "Module1"

Sub All_Tabs()
    
    ' Set initial variable for tab
    Dim xSh As Worksheet
        
    Application.ScreenUpdating = False
    
    ' Loop through all tabs of workbook	
    For Each xSh In Worksheets
        
        xSh.Select
        Call Ticker_Challenge
        
	'format summary table
	Columns("K").NumberFormat = "0.00%"
        Columns("I:L").AutoFit
        Range("I:L").HorizontalAlignment = xlCenter
               
    Next
        
    Application.ScreenUpdating = True
    
End Sub

Sub Ticker_Challenge()

    ' Set an initial variable for holding the ticker symbol
    Dim Ticker_Symbol As String
        
    ' Set an initial variable for holding the total stock volume and starting value
    Dim Stock_Volume As Double
    Stock_Volume = 0
        
    ' Keep track of the location for each ticker symbol in the summary table and starting location
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
            
    ' Set an initial variable for holding the yearly change
    Dim Yearly_Change As Double
            
    ' Set an initial variable for holding ticker open value
    Dim Year_Open As Double
            
    ' Set an initial variable for holding ticker close value
    Dim Year_Close As Double
    
    ' Set an initial variable for holding the percent change
    Dim Percent_Change As Double
                
    ' Establish equation for last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
    ' Label the summary table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
          
    ' Loop through all ticker data
    For i = 2 To LastRow
            
        ' Check if cell immediately following a row is not the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set the ticker symbol
            Ticker_Symbol = Cells(i, 1).Value
            
            ' Add to the total stock volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
            ' Output ticker symbol to summary table
            Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            
            ' Output total stock volume to summary table
            Range("L" & Summary_Table_Row).Value = Stock_Volume
            
            ' Assign year close value
            Year_Close = Cells(i, 6).Value
                    
            ' Calculate yearly change
            Yearly_Change = Year_Close - Year_Open
                    
            ' Output yearly change to summary table
            Range("J" & Summary_Table_Row).Value = Yearly_Change
                           
            If Year_Open = 0 Then
            
                Range("K" & Summary_Table_Row).Value = 0
                
            Else
            
                ' Calculate percent change
            Percent_Change = (Year_Close - Year_Open) / Year_Open
                
            End If
            
            ' Output percent change to summary table
            Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Percent_Change = 0
                
                ' If percent change value is > 0
                If Range("K" & Summary_Table_Row).Value >= 0 Then
                    
                    ' Output conditional formatting
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                ' If percent change value is any other result other than > 0
                Else
                
                    ' Output conditional formatting
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
                  
            ' Reset the stock volume
            Stock_Volume = 0
                    
        ' If the cell immediately preceeding a row is not the same ticker
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
            ' Assign year open value
            Year_Open = Cells(i, 3).Value
            
        ' If the cell immediately following a row is the same ticker
        Else
            
            ' Add to the stock volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
        End If
                    
    Next i

End Sub

