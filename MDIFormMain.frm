VERSION 5.00
Begin VB.MDIForm MDIFormMain 
   BackColor       =   &H80000002&
   Caption         =   "LG Balanza"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   17280
   Icon            =   "MDIFormMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menuBalanza 
      Caption         =   "Balanza"
   End
   Begin VB.Menu menImpesiones 
      Caption         =   "&Impresiones"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu menConfiguracion 
      Caption         =   "&Configuración"
   End
   Begin VB.Menu menuAcercaDe 
      Caption         =   "&Acerca De"
   End
End
Attribute VB_Name = "MDIFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    On Error GoTo Catch
    Dim objMercaderia As New clsMercaderia
    

    If Not (CargarIni) Then GoTo Finally
    


'seto la base de datos
    Set objdDB = New clsAccesoDatos
    Set objImpresion = New clsImpresion
    If Not (objImpresion.CargarDatosBasicos) Then GoTo Finally
    If Not (objMercaderia.ObtenerMercaderia(colMercaderia)) Then GoTo Finally

        menuBalanza_Click

Finally:
        Exit Sub
Catch:
        MsgBox " Error: " & err.Description, vbCritical, Me.Caption
        On Error GoTo 0
        GoTo Finally
End Sub

Private Sub menConfiguracion_Click()
    frmConfiguracion.Show
End Sub


Private Sub mnuAcercaDe_Click()
    frmAcercade.Show
End Sub

Private Sub menImpesiones_Click()
    frmImpresiones.Show
    
End Sub

Private Sub menuAcercaDe_Click()
    frmAcercade.Show

End Sub

Private Sub menuBalanza_Click()
    frmBalanza.Show
End Sub
