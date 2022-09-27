Attribute VB_Name = "Module1"
Sub Wall_Street():


'Define My Variables

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Declare My Variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Close_Price As Double
    Dim Open_Price As Double
    Dim Summary_Table_Row As Double
    
    
'This is the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
    
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("I" & Summary_Table_Row).Value = Total_Stock_Volume
        
        
        Close_Price = Cells(i, 6).Value
        
        If Open_Price = 0 Then
            Percent_Change = 0
        Else: Percent_Change = (Yearly_Change / Open_Price)
        
        End If
        
        'assign yearly change percent
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K" & Summary_Table_Row).NumberFormat = ".0%"
        
        
        'Assign ticker to columm I
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        'Assign Yearly change to column 3
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change



End Sub

