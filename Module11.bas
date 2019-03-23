Attribute VB_Name = "Module11"
Option Explicit

'Karen Alkoby
'Homework #2
'GWU Boot Camp   3/17/19

Sub Stock():

' Declaration variables
Dim total, openPrice, closePrice, yearlyChange, percentChange, maxPerct, minPerct, maxTotal As Double
Dim totalRow, maxIndex, minIndex, i, lastRow, lastRangeRow As Long
Dim tickr, totalMaxTickr, maxTickr, minTickr As String

Dim ws As Worksheet
Dim wsCount As Integer: wsCount = 0

Dim starting_ws As Worksheet


For Each ws In ThisWorkbook.Worksheets   'loop through worksheets
   wsCount = wsCount + 1
  ' Debug.Print wsCount
   
   Sheets(wsCount).Select
   
   'Initialization
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    openPrice = Cells(2, 3).Value
    total = 0: percentChange = 0: closePrice = 0: maxPerct = 0: minPerct = 0: maxTotal = 0
    totalRow = 1: minIndex = maxIndex = 0
            
    Call SetHeaders

    For i = 2 To lastRow
       
       If ((Cells(i, 1).Value) <> (Cells(i + 1, 1).Value)) Then
          total = total + Cells(i, 7).Value
          
          If maxTotal < total Then
             maxTotal = total
             totalMaxTickr = Cells(i, 1).Value
          End If
          
          totalRow = totalRow + 1
          Cells(totalRow, 12).Value = total
          Cells(totalRow, 9).Value = Cells(i, 1).Value
          closePrice = Cells(i, 6).Value
          yearlyChange = (closePrice - openPrice)
          
          
        '  Debug.Print "ClosePrice= " + Str(closePrice) + " OpenPrice= " + Str(openPrice) + " at " + Str(i)
          
          
         
          Cells(totalRow, 10).Value = yearlyChange
         
          
          If Cells(totalRow, 10).Value >= 0 Then
             Cells(totalRow, 10).Interior.Color = vbGreen
          Else
             Cells(totalRow, 10).Interior.Color = vbRed
             
          End If
          If Not openPrice = 0 Then
              'percentChange = ((closePrice - openPrice) / openPrice) * 100
              percentChange = ((closePrice - openPrice) / openPrice)
              Cells(totalRow, 11).Value = Format(percentChange, "Percent")
              
          End If
          
          If maxPerct < percentChange Then
            ' Debug.Print Str(maxPerct) + " < " + Str(percentChange) + " max now"
             maxPerct = percentChange
             maxTickr = Cells(i, 1).Value
          End If
          If minPerct > percentChange Then
            '  Debug.Print Str(minPerct) + " > " + Str(percentChange) + " min now"
             minPerct = percentChange
             minTickr = Cells(i, 1).Value
          End If
            
          
          'Reset the calculated prices
          openPrice = Cells(i + 1, 3).Value
          closePrice = 0
          
         
          total = 0
       Else
       
         total = total + Cells(i, 7).Value
       
       End If
       
     Next i

    ' Debug.Print totalRow
    'Finalize the advanced homework part
     Cells(2, 16).Value = maxTickr
     Cells(2, 17).Value = Format(maxPerct, "Percent")
     
     Cells(3, 16).Value = minTickr
     Cells(3, 17).Value = Format(minPerct, "Percent")
     
     Cells(4, 16).Value = totalMaxTickr
     Cells(4, 17).Value = maxTotal
  
  Next
  Call AutoFitAll

End Sub

Sub SetHeaders()
'Set headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
End Sub

Sub AutoFitAll()
 Application.ScreenUpdating = False
 Dim wkSt As String
 Dim wkBk As Worksheet
 wkSt = ActiveSheet.Name
 
 For Each wkBk In ActiveWorkbook.Worksheets
    On Error Resume Next
    wkBk.Activate
    Cells.EntireColumn.AutoFit
 Next wkBk
 Sheets(wkSt).Select
 Application.ScreenUpdating = True
    
 
 
End Sub

