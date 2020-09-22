Attribute VB_Name = "RC4_Calc"
Public Sub main()
Dim key As String
For i = 1 To 16
    Randomize
    key = key & Chr(Rnd * 255)
Next i
MsgBox RC4(RC4("Hello World!!", key), key)
End Sub
Public Function RC4(inp As String, key As String) As String
Dim S(0 To 255) As Byte, K(0 To 255) As Byte, i As Long
Dim j As Long, temp As Byte, Y As Byte, t As Long, x As Long
Dim Outp As String

For i = 0 To 255
    S(i) = i
Next

j = 1
For i = 0 To 255
    If j > Len(key) Then j = 1
    K(i) = Asc(Mid(key, j, 1))
    j = j + 1
Next i

j = 0
For i = 0 To 255
    j = (j + S(i) + K(i)) Mod 256
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
Next i

i = 0
j = 0
For x = 1 To Len(inp)
    i = (i + 1) Mod 256
    j = (j + S(i)) Mod 256
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
    t = (S(i) + (S(j) Mod 256)) Mod 256
    Y = S(t)
    
    Outp = Outp & Chr(Asc(Mid(inp, x, 1)) Xor Y)
Next
RC4 = Outp
End Function
