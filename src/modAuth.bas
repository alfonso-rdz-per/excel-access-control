Attribute VB_Name = "modAuth"
' =============================================================================
'  modAuth  -  Control de acceso para un libro de Excel
' -----------------------------------------------------------------------------
'  Idea: el libro llega "cerrado" (solo se ve la hoja de bienvenida y el resto
'  de hojas estan MUY ocultas). El usuario debe autenticarse en frmLogin; si la
'  credencial es valida y no ha caducado, se revelan las hojas de trabajo.
'
'  La hoja "Permisos" es la tabla de usuarios:
'      A: usuario        B: fecha de caducidad        C: hash SHA-256 de la clave
'  Esta hoja permanece MUY oculta y protegida; nunca se muestra al usuario.
'
'  Para generar el hash de una clave nueva:  Debug.Print SHA256("miClave")
'
'  >>> Cambia las constantes de abajo por contrasenas propias antes de usar. <<<
' =============================================================================
Option Explicit

' --- Configuracion -----------------------------------------------------------
Public Const SHEET_PERMISOS As String = "Permisos"
Public Const SHEET_INICIO   As String = "Inicio"

Public Const MAX_INTENTOS As Long = 3

' Contrasenas de proteccion (NO son las credenciales de usuario).
' Reemplazar por valores propios. Idealmente inyectarlas en tiempo de despliegue
' en lugar de dejarlas en el codigo fuente.
Public Const PWD_HOJA_PERMISOS As String = "CAMBIAR_PWD_PERMISOS"
Public Const PWD_LIBRO         As String = "CAMBIAR_PWD_LIBRO"
Public Const PWD_HOJAS         As String = "CAMBIAR_PWD_HOJAS"

' --- Estado de la sesion ----------------------------------------------------
Private mAutenticado As Boolean

Public Property Get Autenticado() As Boolean
    Autenticado = mAutenticado
End Property

' ---------------------------------------------------------------------------
'  Bloquea el libro: oculta todo excepto la hoja de bienvenida.
'  Se llama desde Workbook_Open y (opcionalmente) antes de guardar.
' ---------------------------------------------------------------------------
Public Sub BloquearLibro()
    Dim ws As Worksheet

    ThisWorkbook.Unprotect PWD_LIBRO

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_INICIO Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' La estructura protegida impide mostrar hojas manualmente.
    ThisWorkbook.Protect Password:=PWD_LIBRO, Structure:=True

    mAutenticado = False
End Sub

' ---------------------------------------------------------------------------
'  Valida usuario + clave contra la hoja "Permisos".
'  Devuelve un ResultadoLogin describiendo el desenlace.
' ---------------------------------------------------------------------------
Public Enum ResultadoLogin
    LoginOk = 0
    LoginCredencialInvalida = 1
    LoginCaducado = 2
End Enum

Public Function ValidarCredencial(ByVal usuario As String, _
                                  ByVal clave As String) As ResultadoLogin
    Dim ws As Worksheet
    Dim ultimaFila As Long, i As Long
    Dim hashIngresado As String
    Dim caducidad As Variant

    hashIngresado = SHA256(clave)

    Set ws = ThisWorkbook.Sheets(SHEET_PERMISOS)
    ws.Unprotect PWD_HOJA_PERMISOS

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaFila
        If StrComp(Trim$(ws.Cells(i, "A").Value), Trim$(usuario), vbTextCompare) = 0 Then
            If StrComp(CStr(ws.Cells(i, "C").Value), hashIngresado, vbBinaryCompare) = 0 Then
                caducidad = ws.Cells(i, "B").Value
                ws.Protect PWD_HOJA_PERMISOS
                If IsDate(caducidad) And Date > CDate(caducidad) Then
                    ValidarCredencial = LoginCaducado
                Else
                    ValidarCredencial = LoginOk
                End If
                Exit Function
            End If
        End If
    Next i

    ws.Protect PWD_HOJA_PERMISOS
    ValidarCredencial = LoginCredencialInvalida
End Function

' ---------------------------------------------------------------------------
'  Revela las hojas de trabajo tras un login correcto.
' ---------------------------------------------------------------------------
Public Sub DesbloquearLibro()
    Dim ws As Worksheet

    ThisWorkbook.Unprotect PWD_LIBRO

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SHEET_PERMISOS
                ws.Visible = xlSheetVeryHidden
            Case SHEET_INICIO
                ws.Visible = xlSheetVeryHidden
            Case Else
                ws.Visible = xlSheetVisible
                ws.Unprotect PWD_HOJAS
        End Select
    Next ws

    ThisWorkbook.Protect Password:=PWD_LIBRO, Structure:=True
    mAutenticado = True
End Sub

' ---------------------------------------------------------------------------
'  Alta / cambio de clave de un usuario (uso administrativo desde el editor).
'      modAuth.EstablecerUsuario "ana", "claveSegura", DateSerial(2027, 1, 31)
' ---------------------------------------------------------------------------
Public Sub EstablecerUsuario(ByVal usuario As String, _
                             ByVal clave As String, _
                             ByVal caducidad As Date)
    Dim ws As Worksheet, ultimaFila As Long, i As Long, fila As Long

    Set ws = ThisWorkbook.Sheets(SHEET_PERMISOS)
    ws.Unprotect PWD_HOJA_PERMISOS

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    fila = 0
    For i = 2 To ultimaFila
        If StrComp(Trim$(ws.Cells(i, "A").Value), Trim$(usuario), vbTextCompare) = 0 Then
            fila = i
            Exit For
        End If
    Next i
    If fila = 0 Then fila = ultimaFila + 1

    ws.Cells(fila, "A").Value = usuario
    ws.Cells(fila, "B").Value = caducidad
    ws.Cells(fila, "C").Value = SHA256(clave)

    ws.Protect PWD_HOJA_PERMISOS
End Sub
