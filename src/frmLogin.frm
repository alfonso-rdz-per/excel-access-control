VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "Iniciar sesion"
   ClientHeight    =   2400
   ClientWidth     =   3600
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
'  frmLogin  -  Formulario de autenticacion
' -----------------------------------------------------------------------------
'  Controles requeridos (crearlos en el disenador del UserForm):
'      txtUsuario     TextBox
'      txtContrasena  TextBox   (propiedad PasswordChar = "*")
'      cmdAceptar     CommandButton  (Default = True)
'      lblIntentos    Label          (opcional, muestra intentos restantes)
'
'  Toda la logica de validacion vive en modAuth; este formulario solo maneja
'  la interaccion (intentos, mensajes, cierre del libro).
' =============================================================================
Option Explicit

Private mIntentos As Long

Private Sub UserForm_Initialize()
    mIntentos = 0
    ActualizarLblIntentos
End Sub

' Impide cerrar el formulario con la X sin autenticarse: cierra el libro.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And Not modAuth.Autenticado Then
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim resultado As ResultadoLogin

    mIntentos = mIntentos + 1
    resultado = modAuth.ValidarCredencial(txtUsuario.Value, txtContrasena.Value)

    Select Case resultado
        Case LoginOk
            Unload Me
            modAuth.DesbloquearLibro
            MsgBox "Acceso concedido.", vbInformation

        Case LoginCaducado
            MsgBox "Tu acceso ha caducado. El archivo se cerrara.", vbCritical
            ThisWorkbook.Close SaveChanges:=False

        Case LoginCredencialInvalida
            If mIntentos >= modAuth.MAX_INTENTOS Then
                MsgBox "Demasiados intentos fallidos. El archivo se cerrara.", vbCritical
                ThisWorkbook.Close SaveChanges:=False
            Else
                MsgBox "Usuario o contrasena incorrectos.", vbExclamation
                txtUsuario.Value = ""
                txtContrasena.Value = ""
                txtUsuario.SetFocus
                ActualizarLblIntentos
            End If
    End Select
End Sub

Private Sub ActualizarLblIntentos()
    On Error Resume Next
    lblIntentos.Caption = "Intentos restantes: " & (modAuth.MAX_INTENTOS - mIntentos)
    On Error GoTo 0
End Sub
