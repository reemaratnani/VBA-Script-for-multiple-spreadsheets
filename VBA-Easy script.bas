Attribute VB_Name = "Module1"

Sub stockvolume_part1()
        
        
        'Set variable for ticker & total stock volume
    Dim Totalstock_volume As Double
    Dim Ticker As String
    Dim Summary_row As Long
    Dim LastRow As Long
    Dim i As Long 'Set i as varible to loop through all rows in the table
    Dim ws As Worksheet
    
    For Each ws In Worksheets 'To loop through all worksheet
    
         Totalstock_volume = 0
         Summary_row = 2 'Summary-row is to keep track of entry row for the tickers and total stock volume in Summary table
         LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       'Loop through all Tickers to define Total stock volume of each Ticker.
        For i = 2 To LastRow
        
            
            'Check if we are within same Ticker or not...
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                'Set the Ticker Symbol value
                Ticker = ws.Cells(i, 1).Value
                'Add to the total stock volume
                Totalstock_volume = Totalstock_volume + ws.Cells(i, 7).Value
                'Print the Ticker symbol in summary table
                ws.Range("I" & Summary_row).Value = Ticker
                'Print the Ticker total in summary table
                ws.Range("J" & Summary_row).Value = Totalstock_volume

               'Add one to summary_row for the next ticker data
                  Summary_row = Summary_row + 1
            'Reset the total volume to 0
               Totalstock_volume = 0
            Else
            
           'Add to the total stock volume
            Totalstock_volume = Totalstock_volume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
    Next ws
    


End Sub
