Attribute VB_Name = "Module1"
Sub TickerCounter()

Dim Ticker As String

Dim YearOpen As Double
YearOpen = 0

Dim YearClose As Double
YearClose = 0

Dim YearChange As Double
YearChange = 0

Dim Volend As Long
Volend = 0

Dim rowcount As Long
rowcount = 2

Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim PercentChange As Double

Dim BegVol As Double
BegVol = 0


 
 For i = 2 To lastrow
   
    'Check if Ticker changes
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Set Year open price
      YearOpen = Cells(i, 3).Value
     
    ' Set the Ticker Symbol
      Ticker = Cells(i, 1).Value

    ' Set the BegVol
      BegVol = Cells(i, 7).Value
        
    End If
        
   'Check if next Ticker is different
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    ' Set Year close value
        YearClose = Cells(i, 6).Value
        
    ' Calculate Year Change
        YearChange = YearClose - YearOpen
        
    ' Send Year Change to table
        Cells(rowcount, 10).Value = YearChange
        
    ' Send Ticker to table
        Range("I" & rowcount).Value = Ticker
    
    ' Calc PerentChang and send to table
        PercentChange = YearChange / YearOpen
        Cells(rowcount, 11).Value = PercentChange
      
    ' Set the Volume sum and send to table
        Cells(rowcount, 7).Value = Volend
        
        Volume = WorksheetFunction.Sum(BegVol, Volend)
        
        
       
       
          
    ' Add to Row and reset values
        rowcount = rowcount + 1
        
        YearOpen = 0
        YearClose = 0
        YearChange = 0
        PercentChange = 0
        Volend = 0
        BegVol = 0
                
    End If
        
    
 Next i
 
               
    

End Sub
