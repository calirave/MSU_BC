VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub stock()

    'count the number of rows in the data
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Set an initial variable for holding the stock name
    Dim stock As String

    ' Set a variable for holding the opening price
    Dim openprice As Double
    openprice = 0
  
    ' Set a variable for holding the closing price
    Dim closeprice As Double
    closeprice = 0
    
    'Set a variable for the stock volume
    Dim volume As Double
    volume = 0

    'Set a variable for stockcount
    Dim stockcount As Integer
    stockcount = 0
    
    'set variable to track price change
     Dim change As Double
     
     'set variable to track percentage change
      Dim perc_change As Double
    
    ' Keep track of the stocks in summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Loop through all records
    For i = 2 To lastrow

    ' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the stock name
      stock = Cells(i, 1).Value

      ' Add to the volume
      'volume = volume + Cells(i, 7).Value
    
      ' Print the stock name in the Summary Table
      Range("I" & Summary_Table_Row).Value = stock

      ' Print the volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = volume
      
      'Pring stock count to summary table
      Range("M" & Summary_Table_Row).Value = stockcount
      
      'Print stock open price
      Range("N" & Summary_Table_Row).Value = openprice
      'Print stock closing price
      Range("O" & Summary_Table_Row).Value = closeprice
      
      'calculate and print the price change
      change = closeprice - openprice
    
      
      'Print stock change price
      Range("J" & Summary_Table_Row).Value = change
      
      'calculate percentage change
      perc_change = (change / openprice) * 100
      
      'Print stock percentage change
      Range("K" & Summary_Table_Row).NumberFormat = "0.00"
      Range("K" & Summary_Table_Row).Value = perc_change
      
      'apply color to cell based on positive or negative percentage
      If Range("K" & Summary_Table_Row).Value < 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
     Else
      Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
     End If
     
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
     
     
      
      
      ' Reset the volume
      volume = 0

        'reset stock count
        stockcount = 0
        
        openprice = 0
        closeprice = 0
        'change = 0
    
    Else

      ' Add to the volume
        volume = volume + Cells(i, 7).Value
        stockcount = stockcount + 1
        'get closing price for the year
        closeprice = Cells(i + 1, 6).Value
        
    'find open price on first entry for stock
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
      openprice = Cells(i, 3).Value
 
     
       
      End If
      
    
       
    End If
    
  Next i
 

 
   
End Sub


Sub newcount()

 'create new counter for looping through the new table
 lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
 'Cells(2, 18).Value = lastrow2

    
    Set myrange = Worksheets("A").Range("K2:K" & lastrow2)
    highest = Application.WorksheetFunction.Max(myrange)
    lowest = Application.WorksheetFunction.Min(myrange)
    Cells(3, 18).Value = lowest
    Cells(2, 18).Value = highest
    
    Set myrange2 = Worksheets("A").Range("L2:L" & lastrow2)
    highestvolume = Application.WorksheetFunction.Max(myrange2)
    Cells(4, 18).Value = highestvolume
    
    
 
End Sub

