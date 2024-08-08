VERSION 5.00
Begin VB.Form frmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7860
   Icon            =   "frmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7860
   Begin VB.Frame Frame5 
      Caption         =   "Datos Ticket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3960
      TabIndex        =   20
      Top             =   120
      Width           =   3735
      Begin VB.Frame Frame4 
         Caption         =   "Cerificado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3495
         Begin VB.TextBox txtValidadCert 
            Height          =   435
            Left            =   1695
            TabIndex        =   23
            Top             =   840
            Width           =   1680
         End
         Begin VB.TextBox txtCertificado 
            Height          =   435
            Left            =   1695
            TabIndex        =   22
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Validad Cert:"
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   990
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Certificado:"
            Height          =   210
            Left            =   240
            TabIndex        =   24
            Top             =   375
            Width           =   825
         End
      End
   End
   Begin VB.CommandButton cndGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Picture         =   "frmConfiguracion.frx":0B3A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3615
      Begin VB.TextBox txtLogImpresiones 
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtDataReceiving 
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Log de Impresiones:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Log de datos Recibidos :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1770
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   3615
      Begin VB.TextBox txtPesoIni 
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   11
         ToolTipText     =   "Posicion donde comineza el peso en el mensaje"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtTaraIni 
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   10
         ToolTipText     =   "Posicion donde comineza la tara en el mensaje"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Peso Posicion Inicial:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tara Posicion Inicial:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puerto Serie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtRThreshold 
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   7
         ToolTipText     =   "Establecer el umbral de recepción de bytes:18 o 1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtInputLen 
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   5
         ToolTipText     =   "Configurar la cantidad que lee a la vez el:18 o 0"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtSettings 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Configurar la velocidad de transmisión, paridad, bits de datos y bits de parada.ej:9600,E,7,2"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCommPort 
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   1
         ToolTipText     =   "Nro del Puerto Serie ej:1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "RThreshold:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "InputLen:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Settings:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CommPort:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cndGuardar_Click()
On Error GoTo Catch
    If Not EscribirINI2("ConfigPuerto", "CommPort", txtCommPort.Text) Then MsgBox " Error guardando CommPort", vbCritical, Me.Caption
    If Not EscribirINI2("ConfigPuerto", "Settings", txtSettings.Text) Then MsgBox " Error guardando Settings", vbCritical, Me.Caption
    If Not EscribirINI2("ConfigPuerto", "InputLen", txtInputLen.Text) Then MsgBox " Error guardando InputLen", vbCritical, Me.Caption
    If Not EscribirINI2("ConfigPuerto", "RThreshold", txtRThreshold.Text) Then MsgBox " Error guardando RThreshold", vbCritical, Me.Caption
        '##########Variables Parceo Mensaje
    If Not EscribirINI2("ConfigPuerto", "PesoIni", txtPesoIni.Text) Then MsgBox " Error guardando PesoIni", vbCritical, Me.Caption
    If Not EscribirINI2("ConfigPuerto", "TaraIni", txtTaraIni.Text) Then MsgBox " Error guardando TaraIni", vbCritical, Me.Caption
    ' Variable de Logueo
    If Not EscribirINI2("ConfigLog", "DataReceiving", UCase(txtDataReceiving.Text)) Then MsgBox " Error guardando DataReceiving", vbCritical, Me.Caption
    If Not EscribirINI2("ConfigLog", "LogImpresiones", UCase(txtLogImpresiones.Text)) Then MsgBox " Error guardando LogImpresiones", vbCritical, Me.Caption
    '##########Variable DataImpresion
    
    
    
    objImpresion.Certificado = UCase(txtCertificado.Text)
    If Not objImpresion.ActualizaParametro("Certificado", objImpresion.Certificado) Then MsgBox " Error guardando Certificado", vbCritical, Me.Caption
    objImpresion.ValidadCert = UCase(txtValidadCert.Text)
    If Not objImpresion.ActualizaParametro("ValidadCert", objImpresion.ValidadCert) Then MsgBox " Error guardando ValidadCert", vbCritical, Me.Caption

    If Not (CargarIni) Then GoTo Finally
   
   
   MsgBox "Datos guardados ", vbInformation, Me.Caption
Finally:
        Exit Sub
Catch:
        MsgBox " Error: " & err.Description, vbCritical, Me.Caption
        On Error GoTo 0
        GoTo Finally

End Sub

Private Sub Form_Load()
    txtCommPort.Text = intCommPort
    txtSettings.Text = strCommSettings
    txtInputLen.Text = intCommInputLen
    txtRThreshold.Text = intCommRThreshold
    txtPesoIni.Text = intPesoIni
    txtTaraIni.Text = intTaraIni
    txtDataReceiving.Text = strLogDataReceiving
    txtLogImpresiones.Text = strLogImpresiones
    txtCertificado.Text = objImpresion.Certificado
    txtValidadCert.Text = objImpresion.ValidadCert

End Sub

Private Sub txtCertificado_Change()

    txtCertificado.Text = UCase(txtCertificado.Text)
    txtCertificado.SelStart = Len(txtCertificado.Text) ' Mantiene el cursor al final del texto
End Sub

Private Sub txtCommPort_KeyPress(KeyAscii As Integer)
    ' Permitir solo números (0-9) y la tecla de suprimir (Backspace)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) Then
        ' Si no es un número o la tecla de suprimir, cancelar el evento
        KeyAscii = 0
    End If
End Sub

Private Sub txtDataReceiving_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 78, 110, 83, 115, 8
            ' ASCII codes for N, n, S, s and Backspace
            ' Do nothing, allow these keys
        Case Else
            ' Cancel the input if it's not one of the allowed keys
            KeyAscii = 0
    End Select
End Sub

Private Sub txtInputLen_KeyPress(KeyAscii As Integer)
    ' Permitir solo números (0-9) y la tecla de suprimir (Backspace)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) Then
        ' Si no es un número o la tecla de suprimir, cancelar el evento
        KeyAscii = 0
    End If
End Sub

Private Sub txtLogImpresiones_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 78, 110, 83, 115, 8
            ' ASCII codes for N, n, S, s and Backspace
            ' Do nothing, allow these keys
        Case Else
            ' Cancel the input if it's not one of the allowed keys
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPesoIni_KeyPress(KeyAscii As Integer)
    ' Permitir solo números (0-9) y la tecla de suprimir (Backspace)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) Then
        ' Si no es un número o la tecla de suprimir, cancelar el evento
        KeyAscii = 0
    End If
End Sub

Private Sub txtRThreshold_KeyPress(KeyAscii As Integer)
    ' Permitir solo números (0-9) y la tecla de suprimir (Backspace)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) Then
        ' Si no es un número o la tecla de suprimir, cancelar el evento
        KeyAscii = 0
    End If
End Sub

Private Sub txtTaraIni_KeyPress(KeyAscii As Integer)
    ' Permitir solo números (0-9) y la tecla de suprimir (Backspace)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) Then
        ' Si no es un número o la tecla de suprimir, cancelar el evento
        KeyAscii = 0
    End If
End Sub

Private Sub txtValidadCert_Change()
    txtValidadCert.Text = UCase(txtValidadCert.Text)
    txtValidadCert.SelStart = Len(txtValidadCert.Text) ' Mantiene el cursor al final del texto
End Sub
