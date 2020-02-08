VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()


Dim total As Double
Dim Ticker As String
Dim LastRow As Long
Dim YearlyChange As Double
Dim PercentChange As Double

'Loop Through All Worksheets
For Each ws In Worksheets



' Header Names
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 3).Value = "Yearly Open"
 ws.Cells(1, 6).Value = "Yearly Close"

'Set Initial Total
 Dim j As Integer
 total = 0
 Summary_Table_Row = 2

' last Row
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 For i = 2 To LastRow
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set the ticker
                Ticker = Cells(i, 1).Value
                Total_Stock_Vol = Total_Stock_Vol + Cells(i, 7).Value
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
          ' To avoid zero results
            If YearOpen = 0 Then
        
            YearOpen = ws.Cells(i, 3).Value
            
            End If
            
           YearClose = ws.Cells(i, 6).Value
           YearlyChange = YearClose - YearOpen
           
' To avoid zero results
            If YearChange = 0 And YearOpen = 0 Then
        
            YearPercent = 0
            
             Else
           
           YearPercent = YearlyChange / YearOpen
           
            End If

      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
            
             ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
            Else
            
             ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
            End If
           
               YearOpen = ws.Cells(i + 1, 3).Value
           
           
           
           ws.Range("J" & Summary_Table_Row).Value = YearlyChange
           ws.Range("K" & Summary_Table_Row).Value = PercentChange
           ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
           ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Vol
           Summary_Table_Row = Summary_Table_Row + 1
           Total_Stock_Vol = 0
         
         Else
           
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                 
     End If

 Next i
 
  ' Greatest % Increase, % Decrease, and Total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
 
 
Next ws


End Sub
