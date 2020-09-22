Attribute VB_Name = "SkipJack"
'---------------------------------
'This is the NSA algorithm Skipjack
'according to:
'
'SKIPJACK and KEA Algorithm Specifications
'Version 2.0
'29 May 1998
'
'The algorithm can be operated in
'the following modes:
'OFB, CFB, ECB and CBC
'
'More information concerning the algorithm
'can be found at:
'http://csrc.nist.gov/encryption/skipjack-kea.htm
'---------------------------------
'If you make changes to this code then
'please send me a copy.
'---------------------------------
'Author:
'       Asgeir Bjarni Ingvarsson
'       ICQ: 9243261
'       E-Mail: abi@islandia.is
'---------------------------------
'Skipjack is property of the NSA.

Private F
Private K As Long
Private u As Long
Private key(0 To 131) As String
Public Function CBC_Decrypt(inp As String, IV As String) As String
Dim Dat As String, NewIV As String, Outp As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    If i > 1 Then
        NewIV = Mid(inp, i - 16, 16)
    End If
    Dat = Decrypt(Dat)
    If i = 1 Then
        Dat = CryptoXOR(Dat, IV)
    Else
        Dat = CryptoXOR(Dat, NewIV)
    End If
    Outp = Outp & Dat
Next i
CBC_Decrypt = Outp
End Function
Public Function CBC_Encrypt(inp As String, IV As String) As String
Dim Dat As String, NewIV As String, Outp As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    If i = 1 Then
        Dat = CryptoXOR(Dat, IV)
    Else
        Dat = CryptoXOR(Dat, NewIV)
    End If
    NewIV = Encrypt(Dat)
    Outp = Outp & NewIV
Next i
CBC_Encrypt = Outp
End Function
Public Function CFB_Decrypt(inp As String, IV As String) As String
Dim Dat As String, Outp As String, old As String
Dim OldDat As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    If i = 1 Then
        old = Encrypt(IV)
    Else
        old = Encrypt(OldDat)
    End If
    OldDat = Dat
    Outp = Outp & CryptoXOR(Dat, old)
Next i
CFB_Decrypt = Outp
End Function
Public Function CFB_Encrypt(inp As String, IV As String) As String
Dim Dat As String, Outp As String, old As String
Dim OldDat As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    If i = 1 Then
        old = Encrypt(IV)
    Else
        old = Encrypt(OldDat)
    End If
    OldDat = CryptoXOR(Dat, old)
    Outp = Outp & OldDat
Next i
CFB_Encrypt = Outp
End Function
Private Function CryptoXOR(ByVal value1 As String, ByVal value2 As String) As String
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

    CryptoXOR = Right(valueans, 16)
End Function

Private Function Decrypt(inp As String) As String
Dim w1(0 To 32) As String, w3(0 To 32) As String, w4(0 To 32) As String, w2(0 To 32) As String
Dim Counter(0 To 32) As Byte

w1(32) = Mid(inp, 1, 4)
w2(32) = Mid(inp, 5, 4)
w3(32) = Mid(inp, 9, 4)
w4(32) = Mid(inp, 13, 4)
K = 32
u = 31
For i = 0 To 32
    Counter(i) = i + 1
Next

For i = 1 To 8
    w1(K - 1) = InvG(w2(K), key())
    w2(K - 1) = BigXOR(InvG(w2(K), key()), BigXOR(w3(K), Hex(Counter(K - 1))))
    w3(K - 1) = w4(K)
    w4(K - 1) = w1(K)
    u = u - 1
    K = K - 1
Next
For i = 1 To 8
    w1(K - 1) = InvG(w2(K), key())
    w2(K - 1) = w3(K)
    w3(K - 1) = w4(K)
    w4(K - 1) = BigXOR(BigXOR(w1(K), w2(K)), Hex(Counter(K - 1)))
    u = u - 1
    K = K - 1
Next
For i = 1 To 8
    w1(K - 1) = InvG(w2(K), key())
    w2(K - 1) = BigXOR(InvG(w2(K), key()), BigXOR(w3(K), Hex(Counter(K - 1))))
    w3(K - 1) = w4(K)
    w4(K - 1) = w1(K)
    u = u - 1
    K = K - 1
Next
For i = 1 To 8
    w1(K - 1) = InvG(w2(K), key())
    w2(K - 1) = w3(K)
    w3(K - 1) = w4(K)
    w4(K - 1) = BigXOR(BigXOR(w1(K), w2(K)), Hex(Counter(K - 1)))
    u = u - 1
    K = K - 1
Next

Decrypt = w1(0) & w2(0) & w3(0) & w4(0)
End Function
Public Function ECB_Decrypt(inp As String) As String
Dim Dat As String, Outp As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    Outp = Outp & Decrypt(Dat)
Next i
ECB_Decrypt = Outp
End Function
Public Function ECB_Encrypt(inp As String) As String
Dim Dat As String, Outp As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    Outp = Outp & Encrypt(Dat)
Next i
ECB_Encrypt = Outp
End Function
Private Function Encrypt(inp As String) As String
Dim w1(0 To 32) As String, w3(0 To 32) As String, w4(0 To 32) As String, w2(0 To 32) As String
Dim Counter As Long

w1(0) = Mid(inp, 1, 4)
w2(0) = Mid(inp, 5, 4)
w3(0) = Mid(inp, 9, 4)
w4(0) = Mid(inp, 13, 4)
K = 0
Counter = 1

For i = 1 To 8
    w1(K + 1) = BigXOR(BigXOR(G(w1(K), key()), w4(K)), Hex(Counter))
    w2(K + 1) = G(w1(K), key())
    w3(K + 1) = w2(K)
    w4(K + 1) = w3(K)
    Counter = Counter + 1
    K = K + 1
Next
For i = 1 To 8
    w1(K + 1) = w4(K)
    w2(K + 1) = G(w1(K), key())
    w3(K + 1) = BigXOR(BigXOR(w1(K), w2(K)), Hex(Counter))
    w4(K + 1) = w3(K)
    Counter = Counter + 1
    K = K + 1
Next
For i = 1 To 8
    w1(K + 1) = BigXOR(BigXOR(G(w1(K), key()), w4(K)), Hex(Counter))
    w2(K + 1) = G(w1(K), key())
    w3(K + 1) = w2(K)
    w4(K + 1) = w3(K)
    Counter = Counter + 1
    K = K + 1
Next
For i = 1 To 8
    w1(K + 1) = w4(K)
    w2(K + 1) = G(w1(K), key())
    w3(K + 1) = BigXOR(BigXOR(w1(K), w2(K)), Hex(Counter))
    w4(K + 1) = w3(K)
    Counter = Counter + 1
    K = K + 1
Next

Encrypt = w1(32) & w2(32) & w3(32) & w4(32)
End Function
Private Function G(inp As String, key() As String) As String
Dim g1 As String
Dim g2 As String
Dim g3 As String
Dim g4 As String
Dim g5 As String
Dim g6 As String
Dim l As String

g1 = Mid(inp, 1, 2)
g2 = Mid(inp, 3, 2)

l = F(CByte(BigTrans(BigXOR(g2, key(4 * K)))))
g3 = BigXOR(l, g1)
l = F(CByte(BigTrans(BigXOR(g3, key((4 * K) + 1)))))
g4 = BigXOR(l, g2)
l = F(CByte(BigTrans(BigXOR(g4, key((4 * K) + 2)))))
g5 = BigXOR(l, g3)
l = F(CByte(BigTrans(BigXOR(g5, key((4 * K) + 3)))))
g6 = BigXOR(l, g4)

l = g5 & g6
G = l
End Function
Private Function BigTrans(ByVal inp As String) As Double
    inp = Right$(inp, 8)
    tempstr = String$(8 - Len(inp), "0") + inp
    inp = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                inp = inp + "1"
            Else
                inp = inp + "0"
            End If
        Next loopinner
    Next loopit

    Dim o As Double, i As Integer
    o = 0
    For i = Len(inp) To 1 Step -1
        If Mid(inp, i, 1) = "1" Then
            Y = 1
            p = (Len(inp) - i)
            x = 2
            Do While p > 0
                Do While (p / 2) = (p \ 2)
                    x = (x * x) Mod 255
                    p = p / 2
                Loop
                Y = (x * Y) Mod 255
                p = p - 1
            Loop
            o = o + Y
        End If
    Next i
    BigTrans = o
End Function
Private Function BigXOR(ByVal value1 As String, ByVal value2 As String) As String
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
Public Sub Init(Pass As String)
F = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", "F9", "E7", "2D", "4D", _
          "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", "02", "A0", "17", "F1", _
          "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", "63", "9A", "19", "7C", _
          "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", "EE", "B4", "1A", "EA", _
          "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", "5F", "66", "0B", "D8", _
          "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", "9B", "76", "97", "FC", _
          "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", "ED", "9E", "6E", "49", _
          "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", "8D", "C6", "8F", "AA", _
          "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", "70", "88", "61", "2C", "9F", "0D", "2B", "87", "50", "82", "54", _
          "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", "7F", "AB", "E6", "3E", _
          "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", "8C", "0E", "E2", "0C", _
          "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", "08", "77", "11", "BE", _
          "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", "13", "57", "25", "B5", _
          "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")
          
SetKey Pass
End Sub

Private Function InvG(inp As String, key() As String) As String
Dim g1 As String
Dim g2 As String
Dim g3 As String
Dim g4 As String
Dim g5 As String
Dim g6 As String
Dim l As String

g5 = Mid(inp, 1, 2)
g6 = Mid(inp, 3, 2)

l = F(CByte(BigTrans(BigXOR(g5, key((4 * u) + 3)))))
g4 = BigXOR(l, g6)
l = F(CByte(BigTrans(BigXOR(g4, key((4 * u) + 2)))))
g3 = BigXOR(l, g5)
l = F(CByte(BigTrans(BigXOR(g3, key((4 * u) + 1)))))
g2 = BigXOR(l, g4)
l = F(CByte(BigTrans(BigXOR(g2, key(4 * u)))))
g1 = BigXOR(l, g3)

l = g1 & g2
InvG = l
End Function
Public Function OFB_Crypto(inp As String, IV As String) As String
Dim Dat As String, Outp As String, old As String
Dim OldDat As String

For i = 1 To Len(inp) Step 16
    Dat = Mid(inp, i, 16)
    If i = 1 Then
        old = Encrypt(IV)
    Else
        old = Encrypt(OldDat)
    End If
    OldDat = old
    Outp = Outp & CryptoXOR(Dat, old)
Next i
OFB_Crypto = Outp
End Function

Private Sub SetKey(Pass As String)
For i = 0 To 131 Step 10
    If i = 130 Then
        key(i + 0) = Mid(Pass, 1, 2)
        key(i + 1) = Mid(Pass, 3, 2)
    Else
        key(i + 0) = Mid(Pass, 1, 2)
        key(i + 1) = Mid(Pass, 3, 2)
        key(i + 2) = Mid(Pass, 5, 2)
        key(i + 3) = Mid(Pass, 7, 2)
        key(i + 4) = Mid(Pass, 9, 2)
        key(i + 5) = Mid(Pass, 11, 2)
        key(i + 6) = Mid(Pass, 13, 2)
        key(i + 7) = Mid(Pass, 15, 2)
        key(i + 8) = Mid(Pass, 17, 2)
        key(i + 9) = Mid(Pass, 19, 2)
    End If
Next
End Sub
Public Sub main()
Dim Pass As String, inp As String, S As String, l As String

If Test = True Then 'Check if the algorithm has been tampered with(Atn. This is not required except for the first time)
For i = 1 To 10 'Generate random 80-bit key
    Randomize
    m = Hex(Rnd * 255)
    If Len(m) = 1 Then m = "0" & m
    Pass = Pass & m
Next i

Init Pass 'Initialize the key (This is required before encryption and decryption)

'This is the input data. (only 8 bytes at a time)
inp = EnHex("Asgeir!!")
S = Encrypt(inp) 'Encrypt the data in ECB Mode
l = Decrypt(S)
If inp = Decrypt(S) Then 'Check if decrypted correctly(This will have to be checked in another manner)
    MsgBox DeHex(l)
    End
Else
    MsgBox "Incorrect Key!" 'Did not decrypt correctly so show error message
    End
End If
Else
    MsgBox "Failed to verify the algorithm!" 'The algorithm has been tampered with so stop
    End
End If
End Sub
Public Function Test() As Boolean
'Test the algorithm using test vectors
Init "00998877665544332211"
Test = False
If Encrypt("33221100DDCCBBAA") = "2587CAE27A12D300" And Decrypt("2587CAE27A12D300") = "33221100DDCCBBAA" Then
    Test = True
    Exit Function
End If
End Function
Public Function DeHex(inp As String) As String
For i = 1 To Len(inp) Step 2
    x = x & Chr(Val("&H" & Mid(inp, i, 2)))
Next i
DeHex = x
End Function
Public Function EnHex(x As String) As String
For i = 1 To Len(x)
    v = Hex(Asc(Mid(x, i, 1)))
    If Len(v) = 1 Then v = "0" & v
    inp = inp & v
Next i
EnHex = inp
End Function
Public Function Pad(ByVal inp As String) As String
Top:
If Not Len(inp) Mod 8 = 0 Then
    inp = inp & " "
    GoTo Top
End If
Pad = inp
End Function
