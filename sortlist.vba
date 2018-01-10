'sort list

Sub sort()

Dim i, j, iter, rn, rn1, outx, outxm1 As Integer

rn = 6
rn1 = rn + 1
iter = 14
i = 1
j = 1
outx = 2
outxm1 = outx - 1
Cells(outxm1, "f") = 0
Range(Cells(2, "f"), Cells(15, "f")) = 100

Do While j <= iter

    Do While i <= iter
       
            'If Cells(outx, "f") > Cells(rn, 3) And Cells(outx, "f") > Cells(outxm1, "f") Then
            If Cells(rn, 3) < Cells(outx, "f") And Cells(rn, 3) > Cells(outxm1, "f") Then
            
            Cells(outx, "F") = Cells(rn, 3)
            
            'ElseIf Cells(rn, 3) < Cells(outx, "f") And Cells(rn, 3) = Cells(outxm1, "f") Then
            
            End If
    
    rn = rn + 1
    i = i + 1
    
    Loop

outxm1 = outx
outx = outx + 1
j = j + 1
i = 1
rn = 6

Loop
   
End Sub
