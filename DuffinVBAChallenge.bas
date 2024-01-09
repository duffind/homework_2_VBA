Attribute VB_Name = "DuffinVBAChallenge"
Sub WorksheetLoop()
    
    'Loops through all worksheets in workbook and calls subroutines
    'Citation: Microsoft. Microsoft Support. (n.d.). https://support.microsoft.com/en-gb/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

    Dim WS_Count As Long
    Dim WorksheetNumber As Long

    ' Set WS_Count equal to the number of worksheets in the active workbook.
         
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    
    For WorksheetNumber = 1 To WS_Count
    
        Call GroupTickers(WorksheetNumber)
        Call YearlyChange(WorksheetNumber)
        Call PercentChange(WorksheetNumber)
        Call TotalStockVolume(WorksheetNumber)
        Call GreatestTotalVolume(WorksheetNumber)
        Call Color(WorksheetNumber)

    Next WorksheetNumber

End Sub

Sub GroupTickers(WorksheetNumber As Long)

'This subroutine sorts through ticker symbols in column A and places each unique symbol in column I

    Worksheets(WorksheetNumber).Range("I1").Value = "Ticker"
    
    Dim Ticker As String
    Dim EvaluationCell As Variant
    Dim TickerCounter As Long
    Dim PlacementCounter As Long
    Dim BlankCell As Boolean
    
    Ticker = Worksheets(WorksheetNumber).Range("A2").Value
    EvaluationCell = Worksheets(WorksheetNumber).Range("A2").Value
    BlankCell = IsEmpty(EvaluationCell)
    TickerCounter = 2
    PlacementCounter = 2
    
    Do While BlankCell = False
    
        'Main loop continues until the end of column A ticker symbol list reached("BlankCell" is true)
        
        Do Until Ticker <> EvaluationCell
        
            'Nested loop finds where ticker symbols transition then exits to main loop
            
            Ticker = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
        
            TickerCounter = TickerCounter + 1
    
            EvaluationCell = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
        
        Loop
        
        Worksheets(WorksheetNumber).Cells(PlacementCounter, "I").Value = Ticker
        
        PlacementCounter = PlacementCounter + 1
        
        Ticker = EvaluationCell
        
        BlankCell = IsEmpty(EvaluationCell)
        
    Loop
    
    Worksheets(WorksheetNumber).Columns("I:I").EntireColumn.AutoFit
     
End Sub

Sub YearlyChange(WorksheetNumber As Long)

    'This subroutine sorts through ticker symbols in column A and calculates change from yearly opening price and closing per ticker.  Places results in column J.

    Worksheets(WorksheetNumber).Range("J1").Value = "Yearly Change"
    
    Dim Ticker As String
    Dim EvaluationCell As Variant
    Dim TickerCounter As Double
    Dim PlacementCounter As Double
    Dim BlankCell As Boolean
    Dim OpenValueC2 As Double
    Dim RestOpenValue As Double
    Dim CloseValue As Double
    Dim ConserveRestOpenValue As Double
    Dim Count As Integer
    
    Ticker = Worksheets(WorksheetNumber).Range("A2").Value
    EvaluationCell = Worksheets(WorksheetNumber).Range("A2").Value
    BlankCell = IsEmpty(EvaluationCell)
    TickerCounter = 2
    PlacementCounter = 2
    OpenValueC2 = Worksheets(WorksheetNumber).Cells(2, "C").Value

    ConserveRestOpenValue = 0

    Do While BlankCell = False
    
        'Main loop continues until the end of column A ticker symbol list reached("BlankCell" is true)
        
        ConserveRestOpenValue = RestOpenValue
        
        Do Until Ticker <> EvaluationCell
        
            'Nested loop finds where ticker symbols transition and closing and opening values then exits to main loop
            
            Ticker = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
            
            CloseValue = Worksheets(WorksheetNumber).Cells(TickerCounter, "F").Value
            
            TickerCounter = TickerCounter + 1
    
            EvaluationCell = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
            
            RestOpenValue = Worksheets(WorksheetNumber).Cells(TickerCounter, "C").Value
                    
        Loop
        
            If ConserveRestOpenValue = 0 Then
            
            'Nested if calculates the difference between close value and open value. Places the calculation result in column J
            
                Worksheets(WorksheetNumber).Cells(2, "J").Value = CloseValue - OpenValueC2
            
            Else
                
                PlacementCounter = PlacementCounter + 1
        
                Worksheets(WorksheetNumber).Cells(PlacementCounter, "J").Value = CloseValue - ConserveRestOpenValue
        
                Ticker = EvaluationCell
        
                BlankCell = IsEmpty(EvaluationCell)
                
            End If
      
    Loop
    
    'Loops above result in extra erronous calculation in J2.  This code takes all values below J2 and copies them starting J2.  Deletes value in cell at end of resulting column.
    
    Worksheets(WorksheetNumber).Activate
    
    Count = WorksheetFunction.Count("J2", Range("J2", Range("J2").End(xlDown)))
    
    Count = Count + 1
 
    Range("J4", Range("J4").End(xlDown)).Select
    
    Selection.Copy
    
    Range("J3").Select
    
    Worksheets(WorksheetNumber).Paste
    
    Worksheets(WorksheetNumber).Cells(Count, 10).Select
    
    Selection.ClearContents
    
    Worksheets(WorksheetNumber).Columns("J:J").EntireColumn.AutoFit
    
End Sub

Sub PercentChange(WorksheetNumber As Long)

'This subroutine sorts through ticker symbols in column A and calculates percent change from yearly opening price and closing per ticker. Places results in column K.

    Worksheets(WorksheetNumber).Range("K1").Value = "Percent Change"
    
    Dim Ticker As String
    Dim EvaluationCell As Variant
    Dim TickerCounter As Double
    Dim PlacementCounter As Double
    Dim BlankCell As Boolean
    Dim OpenValueC2 As Double
    Dim RestOpenValue As Double
    Dim CloseValue As Double
    Dim ConserveRestOpenValue As Double
    Dim Count As Integer
    
    Ticker = Worksheets(WorksheetNumber).Range("A2").Value
    EvaluationCell = Worksheets(WorksheetNumber).Range("A2").Value
    BlankCell = IsEmpty(EvaluationCell)
    TickerCounter = 2
    PlacementCounter = 2
    OpenValueC2 = Worksheets(WorksheetNumber).Cells(2, "C").Value

    ConserveRestOpenValue = 0

    Do While BlankCell = False
    
        'Main loop continues until the end of column A ticker symbol list reached("BlankCell" is true)
        
        ConserveRestOpenValue = RestOpenValue
        
        Do Until Ticker <> EvaluationCell
        
            'Nested loop finds where ticker symbols transition and closing and opening values then exits to main loop
            
            Ticker = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
            
            CloseValue = Worksheets(WorksheetNumber).Cells(TickerCounter, "F").Value
            
            TickerCounter = TickerCounter + 1
    
            EvaluationCell = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
            
            RestOpenValue = Worksheets(WorksheetNumber).Cells(TickerCounter, "C").Value
                    
        Loop
        
            If ConserveRestOpenValue = 0 Then
            
            'Nested if calculates the difference between close value and open value as percentage. Places the calculation result in column K
            
                Worksheets(WorksheetNumber).Cells(2, "K").Value = FormatPercent((CloseValue - OpenValueC2) / OpenValueC2)
            
            Else
                
                PlacementCounter = PlacementCounter + 1
        
                Worksheets(WorksheetNumber).Cells(PlacementCounter, "K").Value = FormatPercent((CloseValue - ConserveRestOpenValue) / ConserveRestOpenValue)
        
                Ticker = EvaluationCell
        
                BlankCell = IsEmpty(EvaluationCell)
                
            End If
      
    Loop
    
    'Loops above result in extra erroneous calculation in K2.  This code takes all values below K2 and copies them starting K2.  Deletes value in cell at end of resulting column.
    
    Worksheets(WorksheetNumber).Activate
    
    Count = WorksheetFunction.Count("K2", Range("K2", Range("K2").End(xlDown)))
    
    Count = Count + 1
 
    Range("K4", Range("K4").End(xlDown)).Select
    
    Selection.Copy
    
    Range("K3").Select
    
    Worksheets(WorksheetNumber).Paste
    
    Worksheets(WorksheetNumber).Cells(Count, 11).Select
    
    Selection.ClearContents
    
    Worksheets(WorksheetNumber).Columns("K:K").EntireColumn.AutoFit
    
End Sub

Sub TotalStockVolume(WorksheetNumber As Long)

'This subroutine sorts through ticker symbols in column A and calculates total stock volume for each ticker.  Places the result in column L.

    Worksheets(WorksheetNumber).Range("L1").Value = "Total Stock Volume"
    
    Dim Ticker As String
    Dim EvaluationCell As Variant
    Dim TickerCounter As Long
    Dim PlacementCounter As Long
    Dim BlankCell As Boolean
    
    Ticker = Worksheets(WorksheetNumber).Range("A2").Value
    EvaluationCell = Worksheets(WorksheetNumber).Range("A2").Value
    BlankCell = IsEmpty(EvaluationCell)
    TickerCounter = 2
    PlacementCounter = 2
    
    Do While BlankCell = False
    
        'Main loop continues until the end of column A ticker symbol list reached("BlankCell" is true)
        
        Do Until Ticker <> EvaluationCell
        
            'Nested loop finds where ticker symbols transition then exits to main loop
            
            Ticker = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
        
            TickerCounter = TickerCounter + 1
    
            EvaluationCell = Worksheets(WorksheetNumber).Cells(TickerCounter, "A").Value
        
        Loop
        
        Worksheets(WorksheetNumber).Cells(PlacementCounter, "L").Value = WorksheetFunction.SumIf(Range("A:A"), Ticker, Range("G:G"))
        
        PlacementCounter = PlacementCounter + 1
        
        Ticker = EvaluationCell
        
        BlankCell = IsEmpty(EvaluationCell)
        
    Loop
    
    Worksheets(WorksheetNumber).Columns("L:L").EntireColumn.AutoFit


End Sub

Sub GreatestTotalVolume(WorksheetNumber As Long)

'This subroutine finds the minimum percent change, maximum percent change, and greatest total volume in columns K and L.  It finds the corresponding ticker in column I. Publishes resuls in columns Q and P.
'Citation: 7.6 the set keyword in Excel VBA. Excel VBA Programming - using Set. (n.d.). https://www.homeandlearn.org/the_set_keyword.html

    Worksheets(WorksheetNumber).Range("P1").Value = "Ticker"
    Worksheets(WorksheetNumber).Range("Q1").Value = "Value"
    Worksheets(WorksheetNumber).Range("O2").Value = "Greatest % Increase"
    Worksheets(WorksheetNumber).Range("O3").Value = "Greatest % Decrease"
    Worksheets(WorksheetNumber).Range("O4").Value = "Greatest Total Volume"
    
    Dim MaxValue As Variant
    Dim MaxValue2 As Range
    Dim MinValue As Variant
    Dim MinValue2 As Range
    Dim MaxVolume As Variant
    Dim MaxVolume2 As Range
    
    Worksheets(WorksheetNumber).Activate
    
    MaxValue = WorksheetFunction.Max(Range("K2", Range("K2").End(xlDown)))
    MinValue = WorksheetFunction.Min(Range("K2", Range("K2").End(xlDown)))
    MaxVolume = WorksheetFunction.Max(Range("L2", Range("L2").End(xlDown)))
    
    Worksheets(WorksheetNumber).Cells(2, "Q") = FormatPercent(MaxValue)
    Worksheets(WorksheetNumber).Cells(3, "Q") = FormatPercent(MinValue)
    Worksheets(WorksheetNumber).Cells(4, "Q") = MaxVolume
    
    Set MaxValue2 = Range("K2", Range("K2").End(xlDown)).Find(What:=FormatPercent(MaxValue), LookIn:=xlValues)
    Set MinValue2 = Range("K2", Range("K2").End(xlDown)).Find(What:=FormatPercent(MinValue), LookIn:=xlValues)
    Set MaxVolume2 = Range("L2", Range("L2").End(xlDown)).Find(What:=MaxVolume, LookIn:=xlFormulas)

    Range("P2").Value = MaxValue2.Offset(0, -2).Value
    Range("P3").Value = MinValue2.Offset(0, -2).Value
    Range("P4").Value = MaxVolume2.Offset(0, -3).Value
    
    Worksheets(WorksheetNumber).Columns("O:O").EntireColumn.AutoFit
    Worksheets(WorksheetNumber).Columns("P:P").EntireColumn.AutoFit
    Worksheets(WorksheetNumber).Columns("Q:Q").EntireColumn.AutoFit
    
End Sub

Sub Color(WorksheetNumber As Long)

'This subroutine finds the negative and positive values in column J.  It sets negatives to red and positives to green.

    Dim CellCount As Integer
    Dim StartNumber As Integer
    Dim TestValue As Double
    
    Worksheets(WorksheetNumber).Activate
    
    CellCount = WorksheetFunction.Count(Range("J2", Range("J2").End(xlDown)))
    
    
    For StartNumber = 2 To CellCount + 1
    
        TestValue = Worksheets(WorksheetNumber).Cells(StartNumber, "J").Value
        
        If TestValue >= 0 Then
            
            Worksheets(WorksheetNumber).Cells(StartNumber, "J").Interior.Color = RGB(0, 255, 0)
        
        Else
        
            Worksheets(WorksheetNumber).Cells(StartNumber, "J").Interior.Color = RGB(255, 0, 0)
            
        End If
        
        Next StartNumber
        
End Sub

