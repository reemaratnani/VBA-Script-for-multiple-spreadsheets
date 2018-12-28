Attribute VB_Name = "Module1"

Sub stockvolume_part1()
        
        
    Dim Totalstock_volume As Double
    Dim Ticker As String
    Dim Summary_row As Long
    Dim LastRow As Long
    Dim i As Long 
    Dim ws As Worksheet
    
    For Each ws In Worksheets 
    
        Totalstock_volume = 0
        Summary_row = 2 'Summary-row is to keep track of entry row for the tickers and total stock volume in Summary table
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       
        For i = 2 To LastRow
        
            
            
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                Ticker = ws.Cells(i, 1).Value
                Totalstock_volume = Totalstock_volume + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_row).Value = Ticker
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
