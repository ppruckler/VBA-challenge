Attribute VB_Name = "Module1"
Sub sheetsearch()

    Dim ticker As String

    Dim Brand_Total As Double
    Brand_Total = 0
  
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim percent As Double
    percent = 0
    Dim opening As Double
    Dim closing As Double
    Dim diff As Double
    diff = 0
    
   

    For Each ws In Worksheets
        Dim WorksheetName As String
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        WorksheetName = ws.Name
        ws.Cells(1, 10).Value = "ticker"
        ws.Cells(1, 11).Value = "yearly change"
        ws.Cells(1, 12).Value = "percent change"
        ws.Cells(1, 13).Value = "total stock volume"
       
  'for loop to find ticker and total stock
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    Brand_Total = Brand_Total + ws.Cells(i, 7).Value
                    ws.Range("J" & Summary_Table_Row).Value = ticker
                    ws.Range("M" & Summary_Table_Row).Value = Brand_Total
                    Summary_Table_Row = Summary_Table_Row + 1
                    Brand_Total = 0
                    
                Else
                    Brand_Total = Brand_Total + ws.Cells(i, 7).Value
                    
                End If
            Next i
        Summary_Table_Row = 2
        
'for loop to find yearly change
          For i = 2 To LastRow
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    
                    opening = ws.Cells(i, 3).Value
                    closing = ws.Cells(i, 6).Value
                    
                    ws.Range("k" & Summary_Table_Row).Value = diff
                    Summary_Table_Row = Summary_Table_Row + 1
                    diff = 0
                    
                Else
                    diff = closing - opening
                End If
            Next i
          Summary_Table_Row = 2
          
          
   'for loop to find percent
            For i = 2 To LastRow
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                     percent = diff / opening * 100
                    
                    ws.Range("L" & Summary_Table_Row).Value = percent
                    Summary_Table_Row = Summary_Table_Row + 1
                    percent = 0
                Else
                    percent = diff / opening * 100
            End If
            Next i
          Summary_Table_Row = 2
          
          
          
          
          
            Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
          
          
              
    Next ws

End Sub


