'Create new column with no duplicate values from original column

Sub deleteduplicates()

Dim i, j, count1 As Integer
Dim word1, word2 As String

j = 2
i = 2

count1 = Range("B" & Rows.Count).End(xlUp).Row

Cells(1, 1) = count1


Do While i < count1 + 1

word1 = Cells(i, 2)
word2 = Cells(i - 1, 2)

If (StrComp(word1, word2, vbTextCompare) <> 0) Then
    
    Cells(j, 4) = word1
    j = j + 1

End If

i = i + 1
Loop


End Sub
