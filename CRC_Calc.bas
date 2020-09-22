Attribute VB_Name = "CRC_Calc"
Public CRC_Table_Computed
Private CRC_Table(0 To 255) As String
Private Buf()

Private Sub Make_CRC_Table()
For n = 0 To 255
    c = EnHex(CStr(n))
    For k = 0 To 8
        If BigAND(c, "00000001") <> "00000000" Then
            c = BigXOR(BigShiftRight(CStr(c), 1), "EDB88320")
        Else
            c = BigShiftRight(CStr(c), 1)
        End If
    Next
    CRC_Table(n) = c
Next
CRC_Table_Computed = 1
End Sub

Private Function Update_CRC(crc As String, Buf(), length As Long) As String
c = BigXOR(crc, "FFFFFFFF")
If Not CRC_Table_Computed = 1 Then Make_CRC_Table
For n = 0 To length
    c = BigXOR(CRC_Table(DeHex(BigAND(BigXOR(CStr(c), Buf(n)), "000000ff"))), BigShiftRight(CStr(c), 8))
Next
Update_CRC = BigXOR(c, "FFFFFFFF")
End Function

Public Function crc(inp As String) As String
ReDim Buf(0 To (Len(inp) - 1))
For i = 1 To Len(inp)
    Buf((i - 1)) = Hex(Asc(Mid(inp, i, 1)))
Next
crc = Update_CRC("00000000", Buf(), (Len(inp) - 1))
End Function

Private Function DeHex(inp As String) As String
DeHex = Val("&H" & inp)
End Function

Private Function EnHex(X As String) As String
For i = 1 To Len(X)
    v = Hex(Mid(X, i, 1))
    If Len(v) = 1 Then v = "0" & v
    inp = inp & v
Next i
EnHex = inp
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

Private Function BigAND(ByVal value1 As String, ByVal value2 As String) As String
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

    BigAND = valueans
End Function

Private Function BigShiftRight(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftRight = value1
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
        For j = 32 To 2 Step -1
            Mid(value1, j, 1) = Mid(value1, j - 1, 1)
        Next j
        If Not Mid(value1, 1, 1) = "0" Then Mid(value1, 1, 1) = "0"
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

    BigShiftRight = Right(value1, 8)
End Function

Public Sub main()
CRC_Table_Computed = 0
MsgBox crc("Hello World")
End Sub
