Attribute VB_Name = "Gost"
'-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-
'Gosudarstvennyi Standard Soyuza SSR 28147-89
'              (GOST 28147-89)
'-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-
Private S1, S2, S3, S4, S5, S6, S7, S8
Public Function f(R As String, k As String) As String
x = BigMod32Add(R, k)
a = Val("&H" & Mid(x, 1, 1))
b = Val("&H" & Mid(x, 2, 1))
c = Val("&H" & Mid(x, 3, 1))
d = Val("&H" & Mid(x, 4, 1))
e = Val("&H" & Mid(x, 5, 1))
L = Val("&H" & Mid(x, 6, 1))
g = Val("&H" & Mid(x, 7, 1))
h = Val("&H" & Mid(x, 8, 1))

a = S1(a)
b = S2(b)
c = S3(c)
d = S4(d)
e = S5(e)
L = S6(L)
g = S7(g)
h = S8(h)
x = a & b & c & d & e & L & g & h
x = BigShiftLeft(CStr(x), 11)
f = x
End Function
Public Sub Init()
S1 = Array(6, 5, 1, 7, 14, 0, 4, 10, 11, 9, 3, 13, 8, 12, 2, 15)
S2 = Array(14, 13, 9, 0, 8, 10, 12, 4, 7, 15, 6, 11, 3, 1, 5, 2)
S3 = Array(6, 5, 1, 7, 2, 4, 10, 0, 11, 13, 14, 3, 8, 12, 15, 9)
S4 = Array(8, 7, 3, 9, 6, 4, 14, 5, 2, 13, 0, 12, 1, 11, 10, 15)
S5 = Array(10, 9, 6, 11, 5, 1, 8, 4, 0, 13, 7, 2, 14, 3, 15, 12)
S6 = Array(5, 3, 0, 6, 11, 13, 4, 14, 10, 7, 1, 12, 2, 8, 15, 9)
S7 = Array(2, 1, 12, 3, 11, 13, 15, 7, 10, 6, 9, 14, 0, 8, 4, 5)
S8 = Array(6, 5, 1, 7, 8, 9, 4, 2, 15, 3, 13, 12, 10, 14, 11, 0)
End Sub
Public Function Encrypt(ByVal inp As String, ByVal key As String) As String
Dim k(1 To 8) As String
Dim L As String
Dim R As String

k(1) = Mid(key, 1, 8)
k(2) = Mid(key, 8, 8)
k(3) = Mid(key, 16, 8)
k(4) = Mid(key, 24, 8)
k(5) = Mid(key, 32, 8)
k(6) = Mid(key, 40, 8)
k(7) = Mid(key, 48, 8)
k(8) = Mid(key, 56, 8)
For j = 1 To Len(inp) Step 16
    DoEvents
    L = Mid(inp, j, 8)
    R = Mid(inp, j + 8, 8)
    
    For i = 1 To 3
        R = BigXOR(R, f(L, k(1)))
        L = BigXOR(L, f(R, k(2)))
        R = BigXOR(R, f(L, k(3)))
        L = BigXOR(L, f(R, k(4)))
        R = BigXOR(R, f(L, k(5)))
        L = BigXOR(L, f(R, k(6)))
        R = BigXOR(R, f(L, k(7)))
        L = BigXOR(L, f(R, k(8)))
    Next i
    R = BigXOR(R, f(L, k(8)))
    L = BigXOR(L, f(R, k(7)))
    R = BigXOR(R, f(L, k(6)))
    L = BigXOR(L, f(R, k(5)))
    R = BigXOR(R, f(L, k(4)))
    L = BigXOR(L, f(R, k(3)))
    R = BigXOR(R, f(L, k(2)))
    L = BigXOR(L, f(R, k(1)))
    
    Mid(inp, j, 8) = R
    Mid(inp, j + 8, 8) = L
Next j
Encrypt = inp
End Function
Public Function Decrypt(ByVal inp As String, ByVal key As String) As String
Dim k(1 To 8) As String
Dim L As String
Dim R As String

k(1) = Mid(key, 1, 8)
k(2) = Mid(key, 8, 8)
k(3) = Mid(key, 16, 8)
k(4) = Mid(key, 24, 8)
k(5) = Mid(key, 32, 8)
k(6) = Mid(key, 40, 8)
k(7) = Mid(key, 48, 8)
k(8) = Mid(key, 56, 8)
For j = 1 To Len(inp) Step 16
    DoEvents
    L = Mid(inp, j, 8)
    R = Mid(inp, j + 8, 8)

    R = BigXOR(R, f(L, k(1)))
    L = BigXOR(L, f(R, k(2)))
    R = BigXOR(R, f(L, k(3)))
    L = BigXOR(L, f(R, k(4)))
    R = BigXOR(R, f(L, k(5)))
    L = BigXOR(L, f(R, k(6)))
    R = BigXOR(R, f(L, k(7)))
    L = BigXOR(L, f(R, k(8)))
    For i = 1 To 3
        R = BigXOR(R, f(L, k(8)))
        L = BigXOR(L, f(R, k(7)))
        R = BigXOR(R, f(L, k(6)))
        L = BigXOR(L, f(R, k(5)))
        R = BigXOR(R, f(L, k(4)))
        L = BigXOR(L, f(R, k(3)))
        R = BigXOR(R, f(L, k(2)))
        L = BigXOR(L, f(R, k(1)))
    Next i
    
    Mid(inp, j, 8) = R
    Mid(inp, j + 8, 8) = L
Next j
Decrypt = inp
End Function
Public Function GenKey() As String
For i = 1 To 32
    Randomize
    dat = Hex(Rnd * 255)
    If Len(dat) = 1 Then dat = "0" & dat
    key = key & dat
Next i
GenKey = key
End Function
Public Function EnHex(x As String) As String
For i = 1 To Len(x)
    v = Hex(Asc(Mid(x, i, 1)))
    If Len(v) = 1 Then v = "0" & v
    inp = inp & v
Next i
EnHex = inp
End Function
Public Function DeHex(inp As String) As String
For i = 1 To Len(inp) Step 2
    x = x & Chr(Val("&H" & Mid(inp, i, 2)))
Next i
DeHex = x
End Function
Public Function PadInp(inp As String) As String
check1:
If Not (Len(inp) / 16) = (Len(inp) \ 16) Then
    inp = inp & "0"
    GoTo check1
End If
PadInp = inp
End Function
Public Sub main()
Init
key = GenKey
x = PadInp(EnHex("Ásgeir Bjarni Ingvarsson"))
L = Encrypt(CStr(x), CStr(key))
MsgBox DeHex(CStr(L))
inp = Decrypt(CStr(L), CStr(key))
x = DeHex(CStr(inp))
MsgBox x
End Sub
Function BigXOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Xor Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigXOR = Right(valueans, 8)
End Function
Function BigMod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigMod32Add = Right$(BigAdd(value1, value2), 8)
End Function
Public Function BigAdd(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = Space$(Abs(tempnum)) + value1
    ElseIf tempnum > 0 Then
        value2 = Space$(Abs(tempnum)) + value2
    End If

    tempnum = 0
    For loopit = Len(value1) To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(value1, loopit, 1)) + Val("&H" + Mid$(value2, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    If tempnum <> 0 Then
        valueans = Hex$(tempnum) + valueans
    End If

    BigAdd = Right(valueans, 8)
End Function
Public Function BigShiftLeft(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    
    For i = 1 To shifts
        For j = 1 To 32
            Mid(value1, j, 1) = Mid(value1, j + 1, 1)
            If Not Mid(value1, 1, 1) = "0" Then Mid(value1, 1, 1) = "0"
        Next j
    Next i
    tempstr = value1

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    BigShiftLeft = Right(value1, 8)
End Function
