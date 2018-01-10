'Playing with worksheet objects just for fun

Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim I, b1 As Integer

         
         'add more worksheets
         'Sheets.Add After:=Sheets(1), Count:=5
         
         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count
         ActiveWorkbook.Worksheets(1).cells(4, 4) = WS_Count
         
         b1 = 65

         ' Begin the loop.
         For I = 1 To WS_Count
         
           'name worksheets alphabetically. if it produces error, zero out titles ("") and try again
            ActiveWorkbook.Worksheets(I).Name = "bug" & Chr$(b1)
            
            'ActiveWorkbook.Worksheets(I).cells(5, 5).Value = cells(1, 1) + cells(2, 1)
            
            'delete range
            'Worksheets(I).Range("E5:F6").Value = ""
            
            'ActiveWorkbook.Worksheets(I).cells(6, 6).Value = WorksheetFunction.Sum(Range("A1:A6"))
            
            
            'only works in first worksheet
            'cells(7, 7) = WorksheetFunction.Sum(Range("A1:A6"))
                        
            'works globally:
            'Worksheets(I).cells(7, 7) = WorksheetFunction.Sum(Range("A1:A6"))
            
            b1 = b1 + 1
         Next I

      End Sub
