VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Compile_Stock_Data()


Dim ws As Worksheet

For Each ws In Worksheets

    Dim sheetname As String
    Dim i As Long
    Dim j As Long
    'count ticker row
    Dim tickercount As Long
    Dim lastrowA As Long
    Dim Percentchance As Double
    Dim totalvol As Double
    Dim maxchange As Double
    Dim minchange As Double
    Dim grtotalvol As Double
    Dim maxchangeticker As String
    Dim minchangeticker As String
    Dim grtotalvolticker As String
    Dim lastrowi As Long


    

'column names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
'Increase and decrease
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

totalvol = 0

'first row i think?
tickercount = 2
j = 2
'find last cell in column
lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'row loop
For i = 2 To lastrowA
    totalvol = totalvol + ws.Cells(i, 7).Value
    
    'name change
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'fill cells w tick name
    ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
    'yearly change
    ws.Cells(tickercount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    
    'color change
    If ws.Cells(tickercount, 10).Value < 0 Then
    
        ws.Cells(tickercount, 10).Interior.ColorIndex = 3
        
        Else
        
        ws.Cells(tickercount, 10).Interior.ColorIndex = 4
        
        
        End If
        
    'percent change
        
     If ws.Cells(j, 3).Value <> 0 Then
     percentchange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
     
     'Percent Format
     ws.Cells(tickercount, 11).Value = percentchange
     
     Else
     ws.Cells(tickercount, 11).Value = 0
     
     
     End If
     ws.Cells(tickercount, 11).NumberFormat = "0.00%"
     
     'Total volume
     ws.Cells(tickercount, 12).Value = totalvol
     
     totalvol = 0
     
    
'percent change = (close value-open value)/open value

    
    'ticker increase by 1
    tickercount = tickercount + 1
    'set new start point
    j = i + 1
    
    
    
    End If
    
    Next i
    
lastrowi = ws.Cells(Rows.Count, "i").End(xlUp).Row

'Set min max and total variables
maxchange = 0
minchange = 0
grtotalvol = 0

'loop through column i

For i = 2 To lastrowi
    
    'max change
    If ws.Cells(i, "k").Value > maxchange Then
        
    
        maxchange = ws.Cells(i, "k").Value
        maxchangeticker = ws.Cells(i, "i").Value
    End If
    
    'min change
    If ws.Cells(i, "k").Value < minchange Then
    
        minchange = ws.Cells(i, "k").Value
        minchangeticker = ws.Cells(i, "i").Value
    
    End If
    
    'greatest total
    If ws.Cells(i, "l").Value > grtotalvol Then
        
        grtotalvol = ws.Cells(i, "l").Value
        grtotalvolticker = ws.Cells(i, "i").Value
    End If
    
Next i

'fill cells with loop results
ws.Cells(2, "p").Value = maxchangeticker
ws.Cells(2, "q").Value = maxchange
ws.Cells(2, "q").NumberFormat = "0.00%"

ws.Cells(3, "p").Value = minchangeticker
ws.Cells(3, "q").Value = minchange
ws.Cells(3, "q").NumberFormat = "0.00%"
    
ws.Cells(4, "p").Value = grtotalvolticker
ws.Cells(4, "q").Value = grtotalvol

'continue loop through worksheets
Next ws


End Sub
