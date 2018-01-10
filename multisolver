'Excel VBA macro to do solver for lots of cells. Made by someone else and commented by me
'May need to copy text into excel in new macro tab

'limit = ActiveSheet.Range("Q4")

Sub Solver_big()

'

'Worksheets("conc_solve").XML '- Note - This is the name of the spreadsheet where the Solver needs to Run

limit = ActiveSheet.Range("P2")
' This is the cell that has the number of loops corresponding to number of rows where the Solver needs to run
i = 0
' Reset the Loop to zero

Application.ScreenUpdating = True
Do Until i = limit
SolverReset

'F=target col to reproduce, K=constant to change
SolverOk SetCell:="$F$" & 4 + i, MaxMinVal:=1, ValueOf:=0, ByChange:="$K$" & 4 + i, Engine:= _
1, EngineDesc:="GRG Nonlinear"

'Starting at row four. Col K must be <=1.0
SolverAdd CellRef:="$K$" & 4 + i, Relation:=1, FormulaText:="1.0"

'Col L (eqn col) must equal value in F, starting with row 4 for both. F=target, L=eqn
SolverAdd CellRef:="$F$" & 4 + i, Relation:=2, FormulaText:="$L$" & 4 + i
SolverSolve UserFinish:=True
SolverFinish KeepFinal:=1
i = i + 1
Application.DisplayAlerts = False
Loop
End Sub


'relation 1: <=
'relation 2: =
