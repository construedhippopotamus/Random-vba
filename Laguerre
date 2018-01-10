'Laguerre equation

Sub laguerre()
Dim x, g, V, sumk, sumj As Long
Dim i, k, n As Integer


'number of terms
n = 11
k = 2
i = 2     'outer loop variable
sumk = 0
sumj = 0
'

For i = 2 To 15

For k = 2 To n

w = Cells(k, 8)
x = Cells(k, 7)
r1 = Cells(i, 4)
r2 = Cells(i, 3)
V = Cells(2, 11)
mult = Cells(i, 13)

g = ((x / r1) ^ r1) / r1 * Exp(-r1 * r2 * V / x)
sumk = w * g + sumk

'problem is with f(xj, xk) eqn.. What is difference btwn xj and xk?

f = x / (r2 ^ 2) * (x / r2 + r1 * V / x) ^ (r2 - 1)
sumj = w * f + sumj

Next
Cells(3, 14) = mult
Cells(4, 14) = sumk
Cells(5, 14) = sumj
Cells(i, 14) = mult * sumk * sumj

'Solver to find correct V so that fvcalc=0.2

SolverReset
'SolverOptions Precision:=0.001
SolverOK SetCell:=Cells(i, 14), ValueOf:="0.2", ByChange:=Cells(i, 11)
    
'SolverAdd CellRef:=Range("F4:F6"), _
    Relation:=1, _
    FormulaText:=100
'SolverAdd CellRef:=Range("C4:E6"), _
    Relation:=3, _
    FormulaText:=0
'SolverAdd CellRef:=Range("C4:E6"), _
    Relation:=4
SolverSolve UserFinish:=False

SolverFinish KeepFinal:=1

Next

End Sub




