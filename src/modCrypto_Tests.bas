Attribute VB_Name = "modCrypto_Tests"
' Vectores de prueba estándar para SHA-256 (FIPS 180-4 / RFC 6234).
' Ejecutar RunSHA256Tests desde el editor de VBA (F5) antes de confiar en modCrypto.
Option Explicit

Public Sub RunSHA256Tests()
    Check "", "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    Check "abc", "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    Check "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", _
          "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    Check "The quick brown fox jumps over the lazy dog", _
          "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592"
    MsgBox "SHA-256: todas las pruebas pasaron.", vbInformation
End Sub

Private Sub Check(ByVal input As String, ByVal expected As String)
    Dim got As String
    got = SHA256(input)
    If got <> expected Then
        Err.Raise vbObjectError + 1, "RunSHA256Tests", _
            "SHA256(""" & input & """) = " & got & " (se esperaba " & expected & ")"
    End If
End Sub
