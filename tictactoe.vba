'Tic tac toe. Works in excel interface. Slow.

Sub AItictac()

Dim win, xin, yin, B As Integer

win = 1
Line = 1
xin = 1
yin = 1

Range("A1:R3") = ""


Do While win = 1

'player 1 is player
'--------------------------------------------------------------------------------------------------

MsgBox "Player 1 enter coordinates of square"

Call getxy(xin, yin)

If Cells(xin, yin) = 1 Or Cells(xin, yin) = -1 Then
    
    MsgBox "Sorry - no cheating allowed without bribes. Select empty cell."
    
    Call getxy(xin, yin)
    
    Cells(xin, yin) = 1

End If

Cells(xin, yin) = 1

'check win condition
Call checkwin(win)

If win = 0 Then
Exit Sub
End If


'player 2 = AI
'--------------------------------------------------------------------------------------------------

MsgBox "Player 2 enter coordinates of new square"

Call getxy(xin, yin)
  
'ensure cell is currently empty

If Cells(xin, yin) = 1 Or Cells(xin, yin) = -1 Then

MsgBox "Sorry - no cheating allowed without bribes. Select empty cell."

Call getxy(xin, yin)

End If

Cells(xin, yin) = -1

'check win condition
Call checkwin(win)

Loop

End Sub


Sub getxy(xin, yin)

xin = Application.InputBox(prompt:="x coordinate (1-3)=", Type:=1)

'need to make sure they can't erase current selection


'error checking (doesn't handle letters characters etc
If xin > 3 Or xin < 1 Then

xin = Application.InputBox(prompt:="Input x coordinate between 1 and 3", Type:=1)

    If xin > 3 Or xin < 1 Then
    
    MsgBox "Follow the directions. :P"
    
    Exit Sub
    End If
    
End If

yin = Application.InputBox(prompt:="y coordinate (1-3)=", Type:=1)

'error checking (doesn't handle letters characters etc

If yin > 3 Or yin < 1 Then

yin = Application.InputBox(prompt:="Input y coordinate between 1 and 3", Type:=1)

    If yin > 3 Or yin < 1 Then
    
    MsgBox ":P  :P  :P Resounding raspberry noises"
    
    Exit Sub
    End If
    
End If


End Sub

Sub checkwin(win)

'define sets
Dim r1, r2, r3, c1, c2, c3, d1, d2, x, y, i, j, maxi, mini As Double
Dim full As Integer
Dim Array1
Dim Destination As Range
Set Destination = Range("K1")

c1 = WorksheetFunction.Sum(Range("A1:A3"))
c2 = WorksheetFunction.Sum(Range("B1:B3"))
c3 = WorksheetFunction.Sum(Range("C1:C3"))
r1 = WorksheetFunction.Sum(Range("A1:C1"))
r2 = WorksheetFunction.Sum(Range("A2:C2"))
r3 = WorksheetFunction.Sum(Range("A3:C3"))
d1 = Cells(1, 1) + Cells(2, 2) + Cells(3, 3)
d2 = Cells(3, 1) + Cells(2, 2) + Cells(1, 3)
i = 1
j = 1
full = 0
Array1 = Array(c1, c2, c3, r1, r2, r3, d1, d2)

Set Destination = Destination.Resize(1, UBound(Array1) + 1)
Destination.Value = Array1

maxi = WorksheetFunction.max(Array1)
Debug.Print all
mini = WorksheetFunction.min(Array1)

If maxi = 3 Then
MsgBox "Player one wins!"
win = 0

End If

If mini = -3 Then
MsgBox "Player two wins!"
win = 0

End If


For i = 1 To 3
    For j = 1 To 3
        If Abs(Cells(i, j)) > 0 Then   ' if there's a value in the cell, count it
            full = full + 1
        End If
        Debug.Print "full"; full; "i"; i; "j"; j
    Next j
                  
Next i


Debug.Print "i" & i
Debug.Print "j" & j

If full >= 9 And maxi <> 3 And min <> -3 Then
MsgBox "The cat wins!"
win = 0

End If



End Sub





