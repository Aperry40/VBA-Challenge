Attribute VB_Name = "Module1"
Sub Worksheet()
  
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call StockData
    Next
    Application.ScreenUpdating = True
 
End Sub
Sub StockData()

    'Variable to hold the ticker symbol
    Dim Ticker As String
    
    Dim open2 As Double
    
  
    Dim close2 As Double
    
   
   'Variable to hold the yearly change
    Dim yearlyChange As Double
    yearlyChange = 0
   
   'Variable to hold the percent change
    Dim percentChange As Double
    percentChange = 0
   
   
    'Variable to hold the total stock volume
    Dim stockVolume As LongLong
    stockVolume = 0

     
    'Summary table row
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    'Variable to hold the last row
    Dim lastRow As Long
     
    'count the number of rows
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

         
    'loop through all of the ticker rows
    For Row = 2 To lastRow
    
    
    
        'check to see if we are within same ticker symbol. If not display new ticker symbol
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
        
        
            'Set (reset) the Ticker symbol
            Ticker = Range("A" & Row).Value
            close2 = Range("F" & Row).Value
            open2 = Range("C" & Row).Value
            yearlyChange = (close2 - open2)
            stockVolume = Range("G" & Row).Value
             
             
            'Add to summarytable values one last time before change
            yearlyChange = yearlyChange + (close2 - open2)
            stockVolume = Range("G" & Row).Value
            
        If open2 > 0 Then
            percentChange = (yearlyChange / open2) * 100
            
            
            'Add the values to the summarytable
            Range("I" & summaryTableRow).Value = Ticker
            Range("J" & summaryTableRow).Value = yearlyChange
            Range("K" & summaryTableRow).Value = percentChange
            Range("K" & summaryTableRow).NumberFormat = "0.00%"
            Range("L" & summaryTableRow).Value = stockVolume
            
            'once summary table is populated, then add to the summary row count
            summaryTableRow = summaryTableRow + 1
            
            ' then reset the variables within the summary table
            yearlyChange = 0
            percentChange = 0
            stockVolume = 0
                
            
            
        Else
            ' If we are in the same Ticker, add on to the running total
            stockVolume = stockVolume + Range("G" & Row).Value
            
   
            
        End If
        
      End If
        
    Next Row
        
    
End Sub





