Attribute VB_Name = "Sha1_Calc"
Public Function Sha1(hashthis As String) As String
Dim buf(0 To 4) As String
Dim Xin(0 To 79) As String
Dim tempnum As Integer, tempnum2 As Integer
Dim loopit As Integer, loopouter As Integer
Dim loopinner As Integer, a As String
Dim b As String, c As String
Dim d As String, e As String, tempstr As String
Dim Outp As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If

    hashthis = hashthis + String$(tempnum2, Chr$(0))
    
    For loopit = 1 To 8
        tempstr = Chr$(tempnum Mod 256) + tempstr
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit

    hashthis = hashthis + tempstr
    
    ' Set magic numbers
    buf(0) = "67452301"
    buf(1) = "efcdab89"
    buf(2) = "98badcfe"
    buf(3) = "10325476"
    buf(4) = "c3d2e1f0"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = buf(0)
        b = buf(1)
        c = buf(2)
        d = buf(3)
        e = buf(4)
        
        ' Get the 512 bits
        For loopit = 0 To 15
            Xin(loopit) = ""
            For loopinner = 4 To 1 Step -1
                Xin(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + Xin(loopit)
                If Len(Xin(loopit)) Mod 2 Then Xin(loopit) = "0" + Xin(loopit)
            Next loopinner
        Next loopit

        For loopit = 16 To 79
            Xin(loopit) = RotLeft(BXor(BXor(BXor(Xin(loopit - 3), Xin(loopit - 8)), Xin(loopit - 14)), Xin(loopit - 16)), 1)
        Next loopit

        For loopit = 0 To 19
            tempstr = Bor(BAnd(b, c), BAnd(BNot(b), d))
            tempstr = BMod32Add(RotLeft(a, 5), BMod32Add(tempstr, BMod32Add(e, BMod32Add(Xin(loopit), "5A827999"))))
            e = d
            d = c
            c = RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 20 To 39
            tempstr = BXor(BXor(b, c), d)
            tempstr = BMod32Add(RotLeft(a, 5), BMod32Add(tempstr, BMod32Add(e, BMod32Add(Xin(loopit), "6ED9EBA1"))))
            e = d
            d = c
            c = RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 40 To 59
            tempstr = Bor(Bor(BAnd(b, c), BAnd(b, d)), BAnd(c, d))
            tempstr = BMod32Add(RotLeft(a, 5), BMod32Add(tempstr, BMod32Add(e, BMod32Add(Xin(loopit), "8F1BBCDC"))))
            e = d
            d = c
            c = RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        For loopit = 60 To 79
            tempstr = BXor(BXor(b, c), d)
            tempstr = BMod32Add(RotLeft(a, 5), BMod32Add(tempstr, BMod32Add(e, BMod32Add(Xin(loopit), "CA62C1D6"))))
            e = d
            d = c
            c = RotLeft(b, 30)
            b = a
            a = tempstr
        Next loopit

        buf(0) = BMod32Add(buf(0), a)
        buf(1) = BMod32Add(buf(1), b)
        buf(2) = BMod32Add(buf(2), c)
        buf(3) = BMod32Add(buf(3), d)
        buf(4) = BMod32Add(buf(4), e)
    Next loopouter

    ' Extract Hash
    hashthis = ""
    For loopit = 0 To 4
        For loopinner = 0 To 3
            hashthis = hashthis + Chr(Val("&H" + Mid$(buf(loopit), 1 + 2 * loopinner, 2)))
        Next loopinner
    Next loopit

    ' And return it
    Sha1 = hashthis
End Function
Private Function RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotLeft = value1
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
    tempstr = Mid$(value1, rots + 1) + Left$(value1, rots)

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

    RotLeft = value1
End Function
Private Function BXor(ByVal value1 As String, ByVal value2 As String) As String
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

    BXor = valueans
End Function
Private Function Bor(ByVal value1 As String, ByVal value2 As String) As String
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
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Or Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit
    
    Bor = valueans
End Function
Private Function BAnd(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) And Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BAnd = valueans
End Function
Private Function BNot(ByVal value1 As String) As String
Dim valueans As String
Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    For loopit = 1 To 8
        valueans = valueans + Hex$(15 Xor Val("&H" + Mid$(value1, loopit, 1)))
    Next loopit
    
    BNot = valueans
End Function
Private Function BMod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BMod32Add = Right$(BAdd(value1, value2), 8)
End Function
Private Function BAdd(ByVal value1 As String, ByVal value2 As String) As String
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

    BAdd = valueans
End Function
