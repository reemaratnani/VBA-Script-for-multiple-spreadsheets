Attribute VB_Name = "Module12"

Sub stockdata_hard()

    Dim ws As Worksheet
    Dim Totalstock_volume As Double
    Dim Ticker As String
    Dim Summary_row As Long
    Dim LastRow As Long
    Dim i As Long
    Dim yearly_open As Single
    Dim yearly_close As Single
    Dim yearly_change As Double
    Dim percent_change As Double
    
    
    For Each ws In Worksheets 'To loop through all worksheets
    
    Totalstock_volume = 0
    Summary_row = 2 'To keep track of summary table in order to update values for each ticker
    j = 2 'To track the yearly-open value for ticker and initially set to 2 to get the value for first ticker from the second row

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        For i = 2 To LastRow
           
           'To check if next cell ticker is same or not...
           
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                Ticker = ws.Cells(i, 1).Value
                Totalstock_volume = Totalstock_volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_row).Value = Totalstock_volume
                yearly_close = ws.Cells(i, 6).Value
                yearly_open = ws.Cells(j, 3).Value
                yearly_change = yearly_close - yearly_open
                    'For percent_change if denomitor is 0 then division by 0 occurs
                      If yearly_open = 0 Then
                         percent_change = 0
                         ws.Range("K" & Summary_row).Value = Null
                      Else
                         percent_change = yearly_change / yearly_open
                         ws.Range("K" & Summary_row).Value = percent_change
                         ws.Range("K" & Summary_row).NumberFormat = "0.00%" 'to set format of percent_change in percent
                      End If
            
                ws.Range("I" & Summary_row).Value = Ticker
                ws.Range("J" & Summary_row).Value = yearly_change
                      
                       If yearly_change < 0 Then 'To set negative change in red
                         ws.Range("J" & Summary_row).Interior.ColorIndex = 3
                       Else 'To set positive change in green
                         ws.Range("J" & Summary_row).Interior.ColorIndex = 4
                       End If
                   
             
                Summary_row = Summary_row + 1
                'Increase the count of rows to get the yearly_open value for the next ticker
                j = i + 1
               
                Totalstock_volume = 0
            
            
            Else
                Totalstock_volume = Totalstock_volume + ws.Cells(i, 7).Value

                
            End If
            
           
        Next i
        'To get the Greatest % increase/decrease and Greatest total volume from summary table
        
        Dim min, max, max2  As Double 'To set initial variable for the required values
        min = 1000
        max = -1000
        max2 = -100
        
        For Summary_row = 2 To 3500
                    If max < ws.Cells(Summary_row, 11) Then
                       max = ws.Cells(Summary_row, 11)
                       ws.Range("Q2").Value = max
                       ws.Range("P2").Value = ws.Cells(Summary_row, 9)
                    End If
                    If min > ws.Cells(Summary_row, 11) Then
                       min = ws.Cells(Summary_row, 11)
                       ws.Range("Q3").Value = min
                       ws.Range("P3").Value = ws.Cells(Summary_row, 9)
                    End If
                    If max2 < ws.Cells(Summary_row, 12) Then
                       max2 = ws.Cells(Summary_row, 12)
                       ws.Range("Q4").Value = max2
                       ws.Range("P4").Value = ws.Cells(Summary_row, 9)
                    End If
                       ws.Range("Q2 :Q3").NumberFormat = "0.00%"
                       ws.Range("Q4").NumberFormat = "General"
        Next Summary_row
        
    Next ws
    

End Sub
