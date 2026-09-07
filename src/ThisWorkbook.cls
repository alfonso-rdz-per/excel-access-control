VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
'  ThisWorkbook  -  Puntos de entrada del control de acceso
' -----------------------------------------------------------------------------
'  Pega este codigo en el modulo "ThisWorkbook" del proyecto VBA.
' =============================================================================
Option Explicit

Private Sub Workbook_Open()
    modAuth.BloquearLibro
    frmLogin.Show vbModal
End Sub

' Al cerrar el libro se guarda ya "bloqueado": la proxima persona que lo abra
' solo vera la hoja de bienvenida hasta autenticarse.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modAuth.BloquearLibro
    ThisWorkbook.Save
End Sub
