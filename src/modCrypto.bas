Attribute VB_Name = "modCrypto"
' =============================================================================
'  modCrypto  -  SHA-256 en VBA puro (sin dependencias externas ni referencias)
' -----------------------------------------------------------------------------
'  Uso:   Debug.Print SHA256("abc")
'         -> ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad
'
'  Trabaja con enteros de 32 bits sin signo emulados sobre Long/Double, de modo
'  que funciona igual en Office de 32 y 64 bits.
' =============================================================================
Option Explicit

Private K(0 To 63) As Long
Private KReady As Boolean

Private Const TWO32 As Double = 4294967296#
Private Const TWO31 As Double = 2147483648#

' ---- Conversión de/para "unsigned 32-bit" -----------------------------------

Private Function U(ByVal x As Long) As Double
    If x < 0 Then U = x + TWO32 Else U = x
End Function

Private Function DblToLong(ByVal d As Double) As Long
    d = d - Int(d / TWO32) * TWO32          ' wrap a [0, 2^32)
    If d >= TWO31 Then d = d - TWO32
    DblToLong = CLng(d)
End Function

Private Function HexToLong(ByVal h As String) As Long
    HexToLong = DblToLong(Val("&H" & h))
End Function

' ---- Operaciones de bits ----------------------------------------------------

Private Function AddMod32(ParamArray vals() As Variant) As Long
    Dim s As Double, i As Long
    For i = LBound(vals) To UBound(vals)
        s = s + U(CLng(vals(i)))
    Next i
    AddMod32 = DblToLong(s)
End Function

Private Function ShR(ByVal x As Long, ByVal n As Long) As Long
    ShR = DblToLong(Int(U(x) / (2 ^ n)))
End Function

Private Function ShL(ByVal x As Long, ByVal n As Long) As Long
    ShL = DblToLong(U(x) * (2 ^ n))
End Function

Private Function RotR(ByVal x As Long, ByVal n As Long) As Long
    RotR = ShR(x, n) Or ShL(x, 32 - n)
End Function

' ---- Constantes de ronda --------------------------------------------------

Private Sub EnsureK()
    If KReady Then Exit Sub
    Dim h As Variant, i As Long
    h = Array( _
        "428a2f98", "71374491", "b5c0fbcf", "e9b5dba5", "3956c25b", "59f111f1", "923f82a4", "ab1c5ed5", _
        "d807aa98", "12835b01", "243185be", "550c7dc3", "72be5d74", "80deb1fe", "9bdc06a7", "c19bf174", _
        "e49b69c1", "efbe4786", "0fc19dc6", "240ca1cc", "2de92c6f", "4a7484aa", "5cb0a9dc", "76f988da", _
        "983e5152", "a831c66d", "b00327c8", "bf597fc7", "c6e00bf3", "d5a79147", "06ca6351", "14292967", _
        "27b70a85", "2e1b2138", "4d2c6dfc", "53380d13", "650a7354", "766a0abb", "81c2c92e", "92722c85", _
        "a2bfe8a1", "a81a664b", "c24b8b70", "c76c51a3", "d192e819", "d6990624", "f40e3585", "106aa070", _
        "19a4c116", "1e376c08", "2748774c", "34b0bcb5", "391c0cb3", "4ed8aa4a", "5b9cca4f", "682e6ff3", _
        "748f82ee", "78a5636f", "84c87814", "8cc70208", "90befffa", "a4506ceb", "bef9a3f7", "c67178f2")
    For i = 0 To 63
        K(i) = HexToLong(CStr(h(i)))
    Next i
    KReady = True
End Sub

' ---- API pública ----------------------------------------------------------

Public Function SHA256(ByVal message As String) As String
    EnsureK

    ' UTF-8 de la cadena
    Dim bytes() As Byte
    bytes = ToUtf8(message)

    Dim origLen As Long
    origLen = (UBound(bytes) - LBound(bytes) + 1)

    ' Padding: 0x80, ceros, y longitud en bits (64 bits big-endian)
    Dim totalLen As Long
    totalLen = origLen + 1
    Do While (totalLen Mod 64) <> 56
        totalLen = totalLen + 1
    Loop
    totalLen = totalLen + 8

    Dim msg() As Byte
    ReDim msg(0 To totalLen - 1)
    Dim i As Long
    For i = 0 To origLen - 1
        msg(i) = bytes(LBound(bytes) + i)
    Next i
    msg(origLen) = &H80

    Dim bitLen As Double
    bitLen = origLen * 8#
    For i = 0 To 7
        msg(totalLen - 1 - i) = CByte(bitLen - Int(bitLen / 256#) * 256#)
        bitLen = Int(bitLen / 256#)
    Next i

    ' Estado inicial
    Dim H0 As Long, H1 As Long, H2 As Long, H3 As Long
    Dim H4 As Long, H5 As Long, H6 As Long, H7 As Long
    H0 = HexToLong("6a09e667"): H1 = HexToLong("bb67ae85")
    H2 = HexToLong("3c6ef372"): H3 = HexToLong("a54ff53a")
    H4 = HexToLong("510e527f"): H5 = HexToLong("9b05688c")
    H6 = HexToLong("1f83d9ab"): H7 = HexToLong("5be0cd19")

    Dim w(0 To 63) As Long
    Dim chunk As Long, t As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, ch As Long, maj As Long, temp1 As Long, temp2 As Long

    For chunk = 0 To totalLen - 1 Step 64
        For t = 0 To 15
            w(t) = ShL(msg(chunk + t * 4), 24) Or ShL(msg(chunk + t * 4 + 1), 16) _
                 Or ShL(msg(chunk + t * 4 + 2), 8) Or msg(chunk + t * 4 + 3)
        Next t
        For t = 16 To 63
            s0 = RotR(w(t - 15), 7) Xor RotR(w(t - 15), 18) Xor ShR(w(t - 15), 3)
            s1 = RotR(w(t - 2), 17) Xor RotR(w(t - 2), 19) Xor ShR(w(t - 2), 10)
            w(t) = AddMod32(w(t - 16), s0, w(t - 7), s1)
        Next t

        a = H0: b = H1: c = H2: d = H3: e = H4: f = H5: g = H6: hh = H7

        For t = 0 To 63
            s1 = RotR(e, 6) Xor RotR(e, 11) Xor RotR(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            temp1 = AddMod32(hh, s1, ch, K(t), w(t))
            s0 = RotR(a, 2) Xor RotR(a, 13) Xor RotR(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            temp2 = AddMod32(s0, maj)

            hh = g: g = f: f = e
            e = AddMod32(d, temp1)
            d = c: c = b: b = a
            a = AddMod32(temp1, temp2)
        Next t

        H0 = AddMod32(H0, a): H1 = AddMod32(H1, b)
        H2 = AddMod32(H2, c): H3 = AddMod32(H3, d)
        H4 = AddMod32(H4, e): H5 = AddMod32(H5, f)
        H6 = AddMod32(H6, g): H7 = AddMod32(H7, hh)
    Next chunk

    SHA256 = LongToHex(H0) & LongToHex(H1) & LongToHex(H2) & LongToHex(H3) & _
             LongToHex(H4) & LongToHex(H5) & LongToHex(H6) & LongToHex(H7)
End Function

Private Function LongToHex(ByVal x As Long) As String
    Dim d As Double, i As Long, out As String, nib As Long
    d = U(x)
    For i = 1 To 8
        nib = d - Int(d / 16#) * 16#
        out = Mid$("0123456789abcdef", nib + 1, 1) & out
        d = Int(d / 16#)
    Next i
    LongToHex = out
End Function

Private Function ToUtf8(ByVal s As String) As Byte()
    ' Convierte una cadena VBA (UTF-16) a un arreglo de bytes UTF-8.
    Dim i As Long, cp As Long, out() As Byte, n As Long

    If Len(s) = 0 Then
        ReDim out(0 To -1)          ' arreglo de longitud cero
        ToUtf8 = out
        Exit Function
    End If

    ReDim out(0 To Len(s) * 4 - 1)
    n = 0
    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp < &H80 Then
            out(n) = cp: n = n + 1
        ElseIf cp < &H800 Then
            out(n) = &HC0 Or (cp \ &H40): n = n + 1
            out(n) = &H80 Or (cp And &H3F): n = n + 1
        Else
            out(n) = &HE0 Or (cp \ &H1000): n = n + 1
            out(n) = &H80 Or ((cp \ &H40) And &H3F): n = n + 1
            out(n) = &H80 Or (cp And &H3F): n = n + 1
        End If
    Next i

    ReDim Preserve out(0 To n - 1)
    ToUtf8 = out
End Function
