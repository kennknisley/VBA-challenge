VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks():

 'variable for stock ticker
  Dim stock_ticker As String

 'variable for stock volue
  Dim stock_volume As Double
  stock_volume = 0
  
  Dim ws As Worksheet
  
  
  

 'location for each stock ticker
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Dim j As Integer
  Dim change As Double
  Dim start As Long
  Dim i As Long
  Dim percent_change As Double
  Dim daily_change As Double
  Dim averiage_change As Double
  
  For Each ws In Worksheets
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Stock Volume"
  ws.Range("N2").Value = "Greatest Increase"
  ws.Range("N3").Value = "Greatest Decrease"
  ws.Range("N4").Value = "Greatest Volume"
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"
  
  j = 0
  change = 0
  start = 2
  
 'find last row
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row

 'loop through all stock tickers
  For i = 2 To lastRow
  
    'open price?????????
     open_price = ws.Cells(i, 3).Value
  
    'checking stock ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'stock ticker
      stock_ticker = ws.Cells(i, 1).Value

      'stock volume
      stock_volume = stock_volume + ws.Cells(i, 7).Value
      
      If stock_volume = 0 Then
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = 0
        ws.Range("K" & 2 + j).Value = "%" & 0
        ws.Range("L" & 2 + j).Value = 0
      
      Else
        ' Find First non zero starting value
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
      change = ws.Cells(i, 6) - ws.Cells(start, 3)
      percent_change = Round((change / ws.Cells(start, 3) * 100), 2)
      start = i + 1
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = change
        ws.Range("K" & 2 + j).Value = "%" & percent_change
        ws.Range("L" & 2 + j).Value = stock_volume
    Select Case change
    Case Is > 0
    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
    Case Is < 0
    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
    Case Else
    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
    End Select
    End If
    
    stock_volume = 0
    change = 0
    j = j + 1
    

    'if the cell immediately following a row is the same ticker...
    Else

      'add to the stock volume
      stock_volume = stock_volume + ws.Cells(i, 7).Value

    End If

  Next i
  
   ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        
  increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
        
        ws.Range("O2") = ws.Cells(increase_number + 1, 9)
        ws.Range("O3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("O4") = ws.Cells(volume_number + 1, 9)
        
  Next ws
  
End Sub


