VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmBalanza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balamza"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   8175
   Icon            =   "frmBalanza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPesoTotalB 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   26
      Text            =   "000000"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtTaraB 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   25
      Text            =   "000000"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtPesoB 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   24
      Text            =   "000000"
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Variables del Ticket"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin VB.TextBox txtIdentificadorBultoNro 
         Height          =   315
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   27
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox cmbDescripcionMercaderia 
         Height          =   330
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton cmdImpresion 
         Caption         =   "Impresión "
         Enabled         =   0   'False
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
         Picture         =   "frmBalanza.frx":16B92
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         TabIndex        =   8
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox txtTicket 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtIdentificadorBultoTxt 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Text            =   "NUMERO DE PALLET: "
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtIDContenedor 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtNroPermisoEmbarque 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtValidadCert 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtCertificado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Peso (KG)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   20
         Top             =   3960
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion de la Mercaderia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   19
         Top             =   3345
         Width           =   2850
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Identificador de Bulto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   18
         Top             =   2850
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ID de Contenedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   17
         Top             =   2355
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Permiso de Embarque:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   16
         Top             =   1845
         Width           =   2865
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Validad Cert:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   15
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Certificado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   14
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ticket:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   13
         Top             =   360
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6630
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   70556
            MinWidth        =   70556
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   8400
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Finaliza Lectura Balanza"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Picture         =   "frmBalanza.frx":1753C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Inicia Lectura Balanza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmBalanza.frx":17AC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Peso Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4440
      TabIndex        =   23
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Tara"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   22
      Top             =   360
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   21
      Top             =   360
      Width           =   465
   End
End
Attribute VB_Name = "frmBalanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declarar el control MSComm
Private WithEvents MSComm1 As MSComm
Attribute MSComm1.VB_VarHelpID = -1
Private isReceiving As Boolean
Private lngNroLinea As Long
Private intDataFile As Integer

    Dim stxPosINI As Integer
    Dim stxfin As Integer

Private Sub cmdImpresion_Click()

    On Error GoTo Catch
    Dim strPaso As String
    FinalizarLectura

    
    If txtNroPermisoEmbarque.Text = "" Then
        MsgBox "El campo ""Nro de Permiso de Embarque"" no puede estar sin datos", vbInformation, "Generarndo Impresion"
        txtNroPermisoEmbarque.SetFocus
        GoTo Finally
    End If
    If txtIDContenedor.Text = "" Then
        MsgBox "El campo ""ID de Contenedor"" no puede estar sin datos", vbInformation, "Generarndo Impresion"
        txtIDContenedor.SetFocus
        GoTo Finally
    End If
    If cmbDescripcionMercaderia.Text = "" Then
        MsgBox "El campo Descripcion de la Mercaderia no puede estar sin datos", vbInformation, "Generarndo Impresion"
        cmbDescripcionMercaderia.SetFocus
        GoTo Finally
    End If
    
    If txtPeso.Text = "" Then
        MsgBox "El campo Peso (KG) no puede estar sin datos", vbInformation, "Generarndo Impresion"
        GoTo Finally
    End If
    'objImpresion.NroTk
    objImpresion.FechaHora = Now
    'objImpresion.IdEmpresa
    'objImpresion.RazonSocial
    'objImpresion.CUIT
    'objImpresion.CodigoAduana
    'objImpresion.LotPlanta
    'objImpresion.LotBalanza
    
    'objImpresion.Certificado
    'objImpresion.ValidadCert
    objImpresion.NroPermEmbarque = txtNroPermisoEmbarque.Text
    objImpresion.IdContenedor = txtIDContenedor.Text
    objImpresion.IdentificadorBulto = txtIdentificadorBultoTxt.Text & txtIdentificadorBultoNro.Text
    objImpresion.IdMercaderia = cmbDescripcionMercaderia.ItemData(cmbDescripcionMercaderia.ListIndex)

    objImpresion.Mercaderia = cmbDescripcionMercaderia.Text
    objImpresion.Peso = txtPeso.Text
  
    If Not objImpresion.InsertarImpresion Then
        MsgBox " No pudo guardar los datos de la impresion", vbCritical, Me.Caption
        GoTo Finally
    End If
'''''    ' Crear un Recordset en memoria
'''''    strPaso = "Crear un Recordset"
'''''    Dim rs As ADODB.Recordset
'''''    strPaso = "New Recordset"
'''''    Set rs = New ADODB.Recordset
'''''    strPaso = "SET Recordset"
'''''    With rs
'''''        strPaso = "Recordset Crea los campos"
'''''        .Fields.Append "RazonSocial", adVarChar, 25
'''''        .Fields.Append "CUIT", adVarChar, 25
'''''
'''''        .Fields.Append "CodigoAduana", adVarChar, 10
'''''        .Fields.Append "LotPlanta", adVarChar, 10
'''''        .Fields.Append "LotBalanza", adVarChar, 20
'''''
'''''        .Fields.Append "Certificado", adVarChar, 20
'''''        .Fields.Append "ValidadCert", adVarChar, 20
'''''
'''''        .Fields.Append "NroPermEmbarque", adVarChar, 255
'''''        .Fields.Append "IdContenedor", adVarChar, 30
'''''        .Fields.Append "IdentificadorBulto", adVarChar, 30
'''''        .Fields.Append "DescMercaderia", adVarChar, 255
'''''        .Fields.Append "Peso", adVarChar, 255
'''''
'''''        .Open
'''''        ' Agregar datos manualmente
'''''        strPaso = "Carga Recordset"
'''''        .AddNew
'''''        .Fields("RazonSocial").Value = objImpresion.RazonSocial
'''''        .Fields("CUIT").Value = objImpresion.CUIT
'''''
'''''
'''''        .Fields("CodigoAduana").Value = objImpresion.CodigoAduana
'''''        .Fields("LotPlanta").Value = objImpresion.LotPlanta
'''''        .Fields("LotBalanza").Value = objImpresion.LotBalanza
'''''
'''''        .Fields("Certificado").Value = objImpresion.Certificado 'cambia
'''''        .Fields("ValidadCert").Value = objImpresion.ValidadCert 'cambia
'''''        .Fields("NroPermEmbarque").Value = objImpresion.NroPermEmbarque 'cambia
'''''        .Fields("IdContenedor").Value = objImpresion.IdContenedor 'cambia
'''''        .Fields("IdentificadorBulto").Value = objImpresion.IdentificadorBulto 'cambia
'''''        .Fields("DescMercaderia").Value = objImpresion.Mercaderia  'cambia
'''''        .Fields("Peso").Value = objImpresion.Peso 'cambia
'''''
'''''        .Update
'''''    End With
'''''
'''''    strPaso = "pasa parametro Sección4"
'''''    dtrImpresion.Sections("Sección4").Controls("lblFecha").Caption = objImpresion.FechaHora
'''''    strPaso = "pasa parametro Sección2"
'''''    dtrImpresion.Sections("Sección2").Controls("lblTicket").Caption = "TICKET " & Format$(objImpresion.NroTk, "0000000000") ' modifica
'''''
'''''
'''''    ' Asignar el Recordset al DataReport
'''''    Set dtrImpresion.DataSource = rs
'''''
'''''    ' Mostrar el DataReport
'''''    'dtrImpresion.Show
'''''
'''''    ' Configura y muestra el DataReport
'''''    strPaso = "Abre el reporte "
'''''    dtrImpresion.Show vbModal
    

    
    If Not objImpresion.GenerarImpresion Then
        MsgBox " No generar la impresion", vbCritical, Me.Caption
        Unload Me
    End If
        
    If Not objImpresion.ObtenerProximoNroTk Then
        MsgBox " No pudo recuperar el proximo Nro de TK", vbCritical, Me.Caption
        GoTo Finally
    End If
    txtTicket.Text = objImpresion.NroTk
Finally:
        ComenzarLectura
        Exit Sub
Catch:
        MsgBox "Error generarndo la impresion " & vbNewLine & "Paso:" & strPaso & vbNewLine & "Error: " & err.Number & "-" & err.Description, vbCritical, "Generarndo Impresion"
        On Error GoTo 0
        GoTo Finally
End Sub


Private Sub Form_Load()
    On Error GoTo Catch
    Dim objMercaderia As clsMercaderia
    
    ' Crear una instancia del control MSComm
    Set MSComm1 = Me.Controls.Add("MSCommLib.MSComm", "MSComm1")
    
    
   
    ' Configurar el control MSComm
    MSComm1.CommPort = intCommPort ' Configurar el puerto COM1
    MSComm1.Settings = strCommSettings ' Configurar la velocidad de transmisión, paridad, bits de datos y bits de parada
    MSComm1.InputLen = intCommInputLen ' Leer 18 caracteres a la vez o de 0 si se coloca cero se corrije desde MSComm1_OnComm()
    txtCertificado.Text = objImpresion.Certificado
    txtValidadCert.Text = objImpresion.ValidadCert
    

    ' Inicialmente, la comunicación está detenida
    isReceiving = False
    cmbDescripcionMercaderia.AddItem ""
    For Each objMercaderia In colMercaderia
        cmbDescripcionMercaderia.AddItem objMercaderia.Mercaderia
        cmbDescripcionMercaderia.ItemData(cmbDescripcionMercaderia.NewIndex) = objMercaderia.IdMercaderia
    Next
    
Finally:
        Exit Sub
Catch:
        MsgBox " Error: " & err.Description, vbCritical, Me.Caption
        On Error GoTo 0
        GoTo Finally
End Sub

Private Sub cmdStart_Click()
On Error GoTo Catch
    If Not (CargarIni) Then GoTo Finally
    
    stxPosINI = 0
    stxfin = 0
    
    
    
    intDataFile = 0
    lngNroLinea = 0
    ComenzarLectura


    cmdStart.Enabled = False
    cmdStop.Enabled = True
    Frame1.Enabled = True
    cmdImpresion.Enabled = True
    
    If Not objImpresion.ObtenerProximoNroTk Then
        MsgBox " No pudo recuperar el proximo Nro de TK", vbCritical, Me.Caption
        Unload Me
    End If
    txtTicket.Text = objImpresion.NroTk
    
Finally:
        Exit Sub
Catch:
        MsgBox " Error: " & err.Description, vbCritical, Me.Caption
        On Error GoTo 0
        GoTo Finally
'ErrorHandler:
'    MsgBox "Error al iniciar la recepción de datos: " & Err.Description, vbCritical, "Error"
End Sub



Private Sub cmdStop_Click()
    On Error GoTo ErrorHandler
    FinalizarLectura
    ' Cerrar el archivo de texto
    If intDataFile > 0 Then
        Close #intDataFile
    End If
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    Frame1.Enabled = False
    cmdImpresion.Enabled = False
Exit Sub
ErrorHandler:
    MsgBox "Error al detener la recepción de datos: " & err.Description, vbCritical, "Error"
End Sub
Sub FinalizarLectura()
    On Error GoTo ErrorHandler
  
    ' Cerrar el puerto y el archivo de texto cuando se hace clic en "Detener"
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False ' Cerrar el puerto
    End If
    
    ' Marcar que la recepción está detenida
    isReceiving = False
    Exit Sub
ErrorHandler:
    MsgBox "Error al detener la recepción de datos: " & err.Description, vbCritical, "Error"
End Sub
Sub ComenzarLectura()

    ' Inicialmente, la comunicación está detenida
    isReceiving = False
    
    ' Abrir el puerto y el archivo de texto cuando se hace clic en "Iniciar"
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True ' Abrir el puerto
    End If

    ' Configurar el evento OnComm para que se active cuando se reciban datos
    MSComm1.RThreshold = intCommRThreshold ' Establecer el umbral de recepción a 18 bytes o 1 si el valor de MSComm1.InputLen esta en cero
    
    ' Marcar que la recepción está activa
    isReceiving = True
    
End Sub


''bueno funcionando con *
'Private Sub MSComm1_OnComm_()
'    On Error GoTo ErrorHandler
'    Dim strDatoslog As String
'
'    Dim incomingData As String
'    Static buffer As String
'
'
'    Dim message As String
'
'   ' txtPesoB.Text = ""
'    'txtTaraB.Text = ""
'    'txtPesoTotalB.Text = ""
'
'    ' Comprobar el evento de comunicación y si la recepción está activa
'    If isReceiving Then
'        Select Case MSComm1.CommEvent
'            Case comEvReceive ' Evento de recepción de datos
'                ' Leer los datos recibidos del puerto serial
'                incomingData = MSComm1.Input
'                StatusBar.Panels(1).Text = "Data Receive:" + incomingData
'                buffer = buffer & incomingData
'                StatusBar.Panels(2).Text = "Data buffer:" + buffer
'                strDatoslog = lngNroLinea & " - incomingData:" & incomingData & "| buffer:" & buffer
'                LogDatosRecividos (strDatoslog)
'
'                ' Buscar el inicio de un mensaje válido
'                stxPosINI = InStr(buffer, "*")
'                If Len(buffer) > 20 Then
'                    'ObtenerValoresASCII buffer
'                    ' Buscar el próximo inicio de mensaje
'                    stxfin = InStr(stxPosINI + 1, buffer, "*")
'                    If (stxfin > 0) Then
'
'
'                    ' Extraer el mensaje completo
'                    message = Mid(buffer, stxPosINI, 18)
'                    strDatoslog = lngNroLinea & " - message:" & message
'                    LogDatosRecividos (strDatoslog)
'
'                    ' Incrementar el número de línea
'                    lngNroLinea = lngNroLinea + 1
'
'                    ' Escribir los datos recibidos en el archivo de texto'''
'
'                    txtPesoB.Text = Mid(message, intPesoIni, 6)
'                    txtTaraB.Text = Mid(message, intTaraIni, 6)
'                    txtPesoTotalB.Text = CLng(txtPesoB.Text) + CLng(txtTaraB.Text)
'                    txtPeso.Text = txtPesoTotalB.Text ''
'
'                    stxPosINI = 0
'                    stxfin = 0
'                    buffer = ""
'                   End If
'
'                End If
'            Case Else
'                ' Manejar otros eventos del puerto serial si es necesario
'        End Select
'    End If
'
'    Exit Sub
'ErrorHandler:
'    MsgBox "Error Procesando datos recividos, error:" & Err.Description & vbNewLine & vbNewLine & "Finalizara la lectura de la Balanza", vbCritical, "Error"
'    LogError ("Error Procesando datos recividos, error:" & Err.Description)
'    LogError (strDatoslog)
'    cmdStop_Click
'    'Resume Next
'End Sub

'nuevo con formano de manje correcto cortando por <STX>
Private Sub MSComm1_OnComm()
    On Error GoTo ErrorHandler
    Dim strDatoslog As String
    Dim incomingData As String
    Static buffer As String

    Dim message As String

    
    ' Comprobar el evento de comunicación y si la recepción está activa
    If isReceiving Then
        Select Case MSComm1.CommEvent
            Case comEvReceive ' Evento de recepción de datos
                ' Leer los datos recibidos del puerto serial
                incomingData = MSComm1.Input
                ' logueo en el status StatusBar lo que recibo en MSComm1.Input
                StatusBar.Panels(1).Text = "Data Receive:" + incomingData
                
                ' guardo los datos recividos en el buffer
                buffer = buffer & incomingData
                
                ' logueo en el status StatusBar lo que tiene el bufer
                StatusBar.Panels(2).Text = "Data buffer:" + buffer
                strDatoslog = lngNroLinea & " - incomingData:" & incomingData & "| buffer:" & buffer
                
                LogDatosRecividos (strDatoslog)
                
                ' Buscar el inicio de un mensaje válido,dato comienza con <STX>
                ' El carácter \u0002 es un carácter de control en la tabla ASCII y se denomina "Start of Text" (STX).
                ' ASCII del carácter \u0002 es 2
                ' Verificar si el dato comienza con <STX> == Chr(2)
                'If Asc(Mid(incomingData, 1, 1)) = &H2 Then
                
                stxPosINI = InStr(buffer, Chr(2))
                stxfin = InStr(stxPosINI + 1, buffer, Chr(13))
                
                
                If stxPosINI > 0 And stxfin > 0 Then
                'If Len(buffer) > 20 And stxPosINI > 0 Then
                    ObtenerValoresASCII buffer
                       
                    ' Extraer el mensaje completo
                    message = Mid(buffer, stxPosINI, stxfin)
                    strDatoslog = lngNroLinea & " - message:" & message
                    LogDatosRecividos (strDatoslog)

                    ' Incrementar el número de línea
                    lngNroLinea = lngNroLinea + 1

                    ' Escribir los datos recibidos en el archivo de texto'''

                    txtPesoB.Text = Mid(message, intPesoIni, 6)
                    txtTaraB.Text = Mid(message, intTaraIni, 6)
                    On Error GoTo Resindronizar
                    txtPesoTotalB.Text = CLng(txtPesoB.Text) + CLng(txtTaraB.Text)
                    txtPeso.Text = txtPesoTotalB.Text ''
Resindronizar:
                    stxPosINI = 0
                    stxfin = 0
                    buffer = ""
                    
                    txtPeso.Text = txtPesoTotalB.Text ''
                    
                End If
            Case Else
                ' Manejar otros eventos del puerto serial si es necesario
        End Select
    End If
    
    Exit Sub
ErrorHandler:
    buffer = ""
    MsgBox "Error Procesando datos recividos, error:" & err.Description & vbNewLine & vbNewLine & "Finalizara la lectura de la Balanza", vbCritical, "Error"
    LogError ("Error Procesando datos recividos, error:" & err.Description)
    LogError (strDatoslog)
    cmdStop_Click
    'Resume Next
End Sub
'otra bariante de lectura que nunca probe en ela balanza

Private Sub MSComm1_OnComm_1()
    On Error GoTo ErrorHandler
    Dim strDatoslog As String
    Dim incomingData As String
    Static buffer As String

    Dim message As String

    
    ' Comprobar el evento de comunicación y si la recepción está activa
    If isReceiving Then
        Select Case MSComm1.CommEvent
            Case comEvReceive ' Evento de recepción de datos
                ' Leer los datos recibidos del puerto serial
                incomingData = MSComm1.Input
                ' logueo en el status StatusBar lo que recibo en MSComm1.Input
                StatusBar.Panels(1).Text = "Data Receive:" + incomingData
                
                ' guardo los datos recividos en el buffer
                buffer = buffer & incomingData
                
                ' logueo en el status StatusBar lo que tiene el bufer
                StatusBar.Panels(2).Text = "Data buffer:" + buffer
                strDatoslog = lngNroLinea & " - incomingData:" & incomingData & "| buffer:" & buffer
                
                LogDatosRecividos (strDatoslog)
                
                ' Buscar el inicio de un mensaje válido,dato comienza con <STX>
                ' El carácter \u0002 es un carácter de control en la tabla ASCII y se denomina "Start of Text" (STX).
                ' ASCII del carácter \u0002 es 2
                ' Verificar si el dato comienza con <STX> == Chr(2)
                'If Asc(Mid(incomingData, 1, 1)) = &H2 Then
                
                stxPosINI = InStr(buffer, Chr(2))
                stxfin = InStr(stxPosINI + 1, buffer, Chr(13))
                
                
                If stxPosINI > 0 And stxfin > 0 Then
                'If Len(buffer) > 20 And stxPosINI > 0 Then
                    ObtenerValoresASCII buffer
                    ' Buscar el próximo inicio de mensaje
                    stxfin = InStr(stxPosINI + 1, buffer, Chr(2))
                    If (stxfin > 0) Then

                    
                    ' Extraer el mensaje completo
                    message = Mid(buffer, stxPosINI, stxfin)
                    strDatoslog = lngNroLinea & " - message:" & message
                    LogDatosRecividos (strDatoslog)

                    ' Incrementar el número de línea
                    lngNroLinea = lngNroLinea + 1

                    ' Escribir los datos recibidos en el archivo de texto'''

                    txtPesoB.Text = Mid(message, intPesoIni, 6)
                    txtTaraB.Text = Mid(message, intTaraIni, 6)
                    'txtPesoTotalB.Text = CLng(txtPesoB.Text) + CLng(txtTaraB.Text)
                    txtPeso.Text = txtPesoTotalB.Text ''
                   
                    stxPosINI = 0
                    stxfin = 0
                    buffer = ""
                   End If
                    
                End If
            Case Else
                ' Manejar otros eventos del puerto serial si es necesario
        End Select
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Procesando datos recividos, error:" & err.Description & vbNewLine & vbNewLine & "Finalizara la lectura de la Balanza", vbCritical, "Error"
    LogError ("Error Procesando datos recividos, error:" & err.Description)
    LogError (strDatoslog)
    cmdStop_Click
    'Resume Next
End Sub

'Private Sub MSComm1_OnComm()
'    On Error GoTo ErrorHandler
'    txtPesoB.Text = ""
'    txtTaraB.Text = ""
'    txtPesoTotalB.Text = ""
'
'    Dim incomingData As String
'    lngNroLinea = 1 + lngNroLinea
'    ' Comprobar el evento de comunicación y si la recepción está activa
'    If isReceiving Then
'        Select Case MSComm1.CommEvent
'            Case comEvReceive ' Evento de recepción de datos
'                ' Leer los datos recibidos del puerto serial
'                incomingData = MSComm1.Input
'                'StatusBar.Panels(1).Text = "Data Receive:" + incomingData
'                'LogDatosRecividos (lngNroLinea & " - " & incomingData)
'                ' Verificar si el dato comienza con <STX> (02h)
'                If Asc(Mid(incomingData, 1, 1)) = &H2 Then
'                    ' Escribir los datos recibidos en el archivo de texto
'
'                    ''''''''Text2.Text = lngNroLinea & " - " & incomingData & vbCrLf & Text2.Text
'
'                    txtPesoB.Text = Mid(incomingData, intPesoIni, 6)
'                    txtTaraB.Text = Mid(incomingData, intTaraIni, 6)
'                    txtPesoTotalB.Text = CLng(txtPesoB.Text) + CLng(txtTaraB.Text)
'
'                    'StatusBar.Panels(2).Text = "Data: " + incomingData
'
'
'                End If
'            Case Else
'                ' Manejar otros eventos del puerto serial si es necesario
'        End Select
'    End If
'
'    Exit Sub
'ErrorHandler:
'    MsgBox "Error al recibir datos: " & Err.Description, vbCritical, "Error"
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' Cerrar el puerto y el archivo de texto cuando se cierre el formulario
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If
    If isReceiving Then
         ' Cerrar el archivo de texto
        If intDataFile > 0 Then
            Close #intDataFile
        End If

    End If
    ' Liberar el control MSComm
    Set MSComm1 = Nothing
End Sub

Private Sub LogDatosRecividos(strDatos As String)
    If strLogDataReceiving = "S" Then
        If intDataFile = 0 Then
            Dim strArchivo As String
            strArchivo = VB.App.Path & "\Datos_Recibidos" & Format(Now, "yyyymmdd_hhmm") & ".txt"
            ' Abrir el archivo de texto para escribir los datos recibidos
            intDataFile = FreeFile
            Open strArchivo For Append As #intDataFile
        End If
        Print #intDataFile, strDatos
    End If
    'ObtenerValoresASCII strDatos
End Sub

Private Sub LogError(strDatos As String)
    Dim intLogError As Integer
    Dim strArchivo As String
    strArchivo = VB.App.Path & "\LogError_" & Format(Now, "yyyymm") & ".txt"
    intLogError = FreeFile
    Open strArchivo For Append As #intLogError
    Print #intLogError, Format(Now, "dd_hhmmss") & " - " & strDatos
    Close #intLogError
End Sub


Private Sub txtIDContenedor_Change()
    txtIDContenedor.Text = UCase(txtIDContenedor.Text)
    txtIDContenedor.SelStart = Len(txtIDContenedor.Text) ' Mantiene el cursor al final del texto
End Sub

Private Sub txtIdentificadorBultoTxt_Change()
    txtIdentificadorBultoTxt.Text = UCase(txtIdentificadorBultoTxt.Text)
    txtIdentificadorBultoTxt.SelStart = Len(txtIdentificadorBultoTxt.Text) ' Mantiene el cursor al final del texto
End Sub

Private Sub txtNroPermisoEmbarque_Change()
    txtNroPermisoEmbarque.Text = UCase(txtNroPermisoEmbarque.Text)
    txtNroPermisoEmbarque.SelStart = Len(txtNroPermisoEmbarque.Text) ' Mantiene el cursor al final del texto
End Sub
Sub ObtenerValoresASCII(cadena As String)
    Dim resultado As String
    Dim i As Long
    

    Debug.Print cadena
    For i = 1 To Len(cadena)
        resultado = ""
        'resultado = resultado & Asc(Mid(cadena, i, 1)) & " "
        resultado = "Posicion:" & Format(i, "000") & "-Valor=:" & Mid(cadena, i, 1) & "-Valor ASCII:" & Asc(Mid(cadena, i, 1)) & "-"
        Debug.Print resultado
    Next i

End Sub
'ROTOCOLO DE COMUNICACIÓN W180-T
'salida SERIE
'La salida de datos por RS232C es en forma contínua, la velocidad se puede
'programar
'entre 1200 baud y 9600 con 7 bits de datos, paridad par y dos bits de stop.
'Con cada conversión A/D se envían 18 caracteres en el siguiente formato:
'<STX><A><B><C><PESO><TARA><CR><LF>
'6 de 9
'STX
'Carácter de sincronización = 02h
'A - B C
'Palabras de estatus
'Peso
'Peso Neto en 6 dígitos, sin punto decimal
'Tara
'Tara en 6 dígitos, sin punto decimal
'CR / LF
'Fin de línea = 0Dh / 0Ah
'Estatus A Bits: 7 6 5 4 3 2 1 0
'PPPPPPPP 00000000 11111111 101 011 011 101 xxx125xxx215000 00100011
'11000001 10011110 00000000, , ,,0000 000 00
'Estatus B 0 1 = peso Neto
'1 1 = peso negativo
'2 1 = fuera de rango
'3 1 = fuera de equilibrio
'4 1
'5 1
'6 0
'7 paridad
'Estatus C 0 0
'1 0
'2 0
'3 1 = tecla <AD> pulsada
'4 0
'5 1
'6 0
'7 paridad
