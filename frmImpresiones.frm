VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.OCX"
Object = "{FB1451DA-4C9C-11D2-AD92-00C0F012D38C}#11.0#0"; "DPINPUTBOX.OCX"
Begin VB.Form frmImpresiones 
   Caption         =   "Impresiones"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   17175
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkNroTicket 
      Caption         =   "Nro Ticket"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   24
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de búsqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Width           =   16695
      Begin VB.ComboBox cmbDescripcionMercaderia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox chkMercaderia 
         Caption         =   "Mercaderia"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkFechas 
         Caption         =   "Fechas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkUltimoImpreso 
         Caption         =   "Ultimo Impreso"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
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
         Left            =   15240
         Picture         =   "frmImpresiones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtNrotkBusqueda 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin DPInputBox.DPInputBoxCtrol txtFechaDesde 
         Height          =   315
         Left            =   2760
         TabIndex        =   27
         Tag             =   "8"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         DecimalPoint    =   " "
         Enabled         =   0   'False
         RequiredLabel   =   "<Obligatorio>"
         IllegalValueStr =   "Error: Fecha Inválida"
         MaxLength       =   10
         BackColor       =   -2147483643
         FieldType       =   4
         DateFormat      =   2
      End
      Begin DPInputBox.DPInputBoxCtrol txtFechaHasta 
         Height          =   315
         Left            =   4800
         TabIndex        =   28
         Tag             =   "8"
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         DecimalPoint    =   " "
         Enabled         =   0   'False
         RequiredLabel   =   "<Obligatorio>"
         IllegalValueStr =   "Error: Fecha Inválida"
         MaxLength       =   10
         BackColor       =   -2147483643
         FieldType       =   4
         DateFormat      =   2
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
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
         Height          =   240
         Left            =   4200
         TabIndex        =   26
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         Height          =   240
         Left            =   2040
         TabIndex        =   25
         Top             =   1200
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   16695
      Begin VB.TextBox txtMercaderia 
         Height          =   315
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCertificado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtValidadCert 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNroPermisoEmbarque 
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtIDContenedor 
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtIdentificadorBultoTxt 
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtTicket 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1695
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdImpresion 
         Caption         =   "Impresión "
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
         Left            =   15120
         Picture         =   "frmImpresiones.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   17
         Top             =   360
         Width           =   630
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
         Left            =   3465
         TabIndex        =   16
         Top             =   375
         Width           =   1095
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
         Left            =   6705
         TabIndex        =   15
         Top             =   390
         Width           =   1245
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
         TabIndex        =   14
         Top             =   885
         Width           =   2865
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
         Left            =   5505
         TabIndex        =   13
         Top             =   915
         Width           =   1725
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
         TabIndex        =   12
         Top             =   1410
         Width           =   2115
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
         Left            =   5880
         TabIndex        =   11
         Top             =   1410
         Width           =   2850
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
         Left            =   345
         TabIndex        =   10
         Top             =   1920
         Width           =   930
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdImpresines 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   8
      Cols            =   11
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12648447
      BackColorSel    =   14737632
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmImpresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum Grilla
    grd_IdImpresion = 0
    grd_NroTk = 1
    grd_FechaHora = 2
    grd_RazonSocial = 3
    grd_CUIT = 4
    grd_CodigoAduana = 5
    grd_LotPlanta = 6
    grd_LotBalanza = 7
    grd_Certificado = 8
    grd_ValidadCert = 9
    grd_NroPermEmbarque = 10
    grd_IdContenedor = 11
    grd_IdentificadorBulto = 12
    grd_Mercaderia = 13
    grd_Peso = 14
End Enum

Dim objImpresionTmp As clsImpresion


Private Sub chkFechas_Click()
    If chkFechas.Value = 1 Then
        txtFechaDesde.Enabled = True
        txtFechaHasta.Enabled = True
        Label9.Enabled = True
        Label10.Enabled = True
        
    Else
        txtFechaDesde.Enabled = False
        txtFechaHasta.Enabled = False
        Label9.Enabled = False
        Label10.Enabled = False
    End If
End Sub

Private Sub chkMercaderia_Click()

    If chkMercaderia.Value = 1 Then
        cmbDescripcionMercaderia.Enabled = True
    Else
        cmbDescripcionMercaderia.Enabled = False
    End If

End Sub

Private Sub chkUltimoImpreso_Click()

    If chkUltimoImpreso.Value = 1 Then
        chkFechas.Value = 0
        chkMercaderia.Value = 0
        chkNroTicket.Value = 0
        
        chkFechas.Enabled = False
        chkMercaderia.Enabled = False
        chkNroTicket.Enabled = False
    Else
        chkFechas.Enabled = True
        chkMercaderia.Enabled = True
        chkNroTicket.Enabled = True
        txtNrotkBusqueda.Enabled = True
        
        
    End If
End Sub
Private Sub chkNroTicket_Click()
    If chkNroTicket.Value = 1 Then
        chkFechas.Value = 0
        chkMercaderia.Value = 0
        chkUltimoImpreso.Value = 0
        
        chkFechas.Enabled = False
        chkMercaderia.Enabled = False
        chkUltimoImpreso.Enabled = False
        txtNrotkBusqueda.Enabled = True
    Else
        chkFechas.Enabled = True
        chkMercaderia.Enabled = True
        chkUltimoImpreso.Enabled = True
        txtNrotkBusqueda.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim strSql As String
    Dim strFechaDesde As String
    LimpiarTextosImpresion
    
    If chkUltimoImpreso.Value = 0 And chkNroTicket.Value = 0 And chkFechas.Value = 0 And chkMercaderia.Value = 0 Then
        MsgBox "Debe seleccionar una opcion de Busqueda", vbInformation, "Buscando Impresion"
        Exit Sub
    End If
    'campos
    strSql = "SELECT Impresiones.IdImpresion, Impresiones.NroTk, Impresiones.FechaHora, Empresa.RazonSocial, Empresa.CUIT, Impresiones.CodigoAduana, Impresiones.LotPlanta, Impresiones.LotBalanza, Impresiones.Certificado, Impresiones.ValidadCert, Impresiones.NroPermEmbarque, Impresiones.IdContenedor, Impresiones.IdentificadorBulto, Mercaderia.Mercaderia, Impresiones.Peso"
    'from
    strSql = strSql & " FROM (Empresa INNER JOIN Impresiones ON Empresa.IdEmpresa = Impresiones.IdEmpresa) INNER JOIN Mercaderia ON Impresiones.IdMercaderia = Mercaderia.IdMercaderia"

    If chkUltimoImpreso.Value = 1 Then
        'strSql = "select * from Cons_UltimaImpresion"
        'where
        strSql = strSql & " WHERE Impresiones.NroTk =(select max( Impresiones.NroTk) from  Impresiones);"
    End If
    
    If chkNroTicket.Value = 1 Then
        If txtNrotkBusqueda.Text = "" Then
            MsgBox "El campo ""Nro Ticket"" no puede estar sin datos", vbInformation, "Buscando Impresion"
            txtNrotkBusqueda.SetFocus
            Exit Sub
        End If
        strSql = strSql & " WHERE Impresiones.NroTk =" & txtNrotkBusqueda.Text & ";"
        
    End If
    
    If chkFechas.Value = 1 Then
        If txtFechaDesde.Value = "" Then
            MsgBox "El campo ""Fecha Desde"" no puede estar sin datos", vbInformation, "Buscando Impresion"
            txtFechaDesde.SetFocus
            Exit Sub
        End If
        If Not IsDate(txtFechaDesde.Value) Then
            MsgBox "La ""Fecha Desde"" ingresada no es válida. Por favor, ingrese una fecha en el formato correcto."
            txtFechaDesde.SetFocus
            Exit Sub
        End If
        
        If txtFechaHasta.Value = "" Then
            strFechaDesde = Format(Now, "yyyy-mm-dd hh:nn:ss")
        Else
            If Not IsDate(txtFechaDesde.Value) Then
                MsgBox "La ""Fecha Hasta"" ingresada no es válida. Por favor, ingrese una fecha en el formato correcto."
                txtFechaDesde.SetFocus
                Exit Sub
            End If
            strFechaDesde = txtFechaDesde.Value
        End If
    End If
    If chkMercaderia.Value = 1 Then
        If cmbDescripcionMercaderia.Text = "" Then
            MsgBox "El campo ""Mercaderia"" no puede estar sin datos", vbInformation, "Buscando Impresion"
            cmbDescripcionMercaderia.SetFocus
            Exit Sub
        End If
    End If
    
    
    If chkFechas.Value = 1 And chkMercaderia.Value = 0 Then
        strSql = strSql & " WHERE Impresiones.FechaHora >=#" & Format(txtFechaDesde.Value, "yyyy-mm-dd hh:nn:ss") & "# and Impresiones.FechaHora <=#" & strFechaDesde & "#  ;"
    End If
    

    If chkFechas.Value = 0 And chkMercaderia.Value = 1 Then
        strSql = strSql & " WHERE Impresiones.IdMercaderia=" & cmbDescripcionMercaderia.ItemData(cmbDescripcionMercaderia.ListIndex) & ";"
    End If
    
    If chkFechas.Value = 1 And chkMercaderia.Value = 1 Then
        strSql = strSql & " WHERE Impresiones.IdMercaderia=" & cmbDescripcionMercaderia.ItemData(cmbDescripcionMercaderia.ListIndex) & " and "
        strSql = strSql & " Impresiones.FechaHora >=#" & Format(txtFechaDesde.Value, "yyyy-mm-dd hh:nn:ss") & "# and Impresiones.FechaHora <=#" & strFechaDesde & "#  ;"
    End If
    
    
    ' Cargo los datos del Cliente
    Dim objRS As ADODB.Recordset
    
    If Not objImpresion.ObtenerImpresiones(objRS, strSql) Then
        Exit Sub
    End If
    CargaGrillaxRecordSet objRS
End Sub



Private Sub cmdImpresion_Click()
    If txtPeso.Text = "" Then
        MsgBox "El campo Peso (KG) no puede estar sin datos, selecciones una impresion", vbInformation, "Generarndo Impresion"
        Exit Sub
    End If
    
    If Not objImpresionTmp.GenerarImpresion Then
        MsgBox " No generar la impresion", vbInformation, Me.Caption
    End If
    
End Sub

Private Sub Form_Load()
    cmbDescripcionMercaderia.AddItem ""
    For Each objMercaderia In colMercaderia
        cmbDescripcionMercaderia.AddItem objMercaderia.Mercaderia
        cmbDescripcionMercaderia.ItemData(cmbDescripcionMercaderia.NewIndex) = objMercaderia.IdMercaderia
    Next
    ConfiguroGrilla
End Sub


Public Sub CargaGrillaxRecordSet(rsResult As ADODB.Recordset)
    Dim i As Integer

    Dim iCantMaxRegistros As Integer

    On Error GoTo Catch

    '-----------------------------------
    '-- Valido la cantidad de respuestas
    '-----------------------------------
'    If rsResult.RecordCount >= 5000 Then
'        MsgBox ("La cantidad de registros devueltos es muy grande. Solo se mostraran los primeros 5.000 registros. Trate de aplicar mas filtros a la busqueda"), vbInformation, Me.Caption
'
'        iCantMaxRegistros = 5000
'    Else
'        iCantMaxRegistros = rsResult.RecordCount
'    End If

    ConfiguroGrilla

    grdImpresines.Visible = False
    DoEvents
    grdImpresines.Rows = iCantMaxRegistros + 1


    If rsResult.RecordCount > 0 Then
        Me.grdImpresines.Rows = rsResult.RecordCount + 1
        i = 1
        While rsResult.EOF = False

            With Me.grdImpresines
                .TextMatrix(i, grd_IdImpresion) = CStr(Format$(rsResult.Fields("IdImpresion").Value, "0000000000"))
                .TextMatrix(i, grd_NroTk) = nullToString(rsResult("NroTk"))
                .TextMatrix(i, grd_FechaHora) = nullToString(rsResult("FechaHora"))
                .TextMatrix(i, grd_RazonSocial) = nullToString(rsResult("RazonSocial"))
                .TextMatrix(i, grd_CUIT) = nullToString(rsResult("CUIT"))
                .TextMatrix(i, grd_CodigoAduana) = nullToString(rsResult("CodigoAduana"))
                .TextMatrix(i, grd_LotPlanta) = nullToString(rsResult.Fields("LotPlanta").Value)
                .TextMatrix(i, grd_LotBalanza) = nullToString(rsResult("LotBalanza"))
                .TextMatrix(i, grd_Certificado) = nullToString(rsResult("Certificado"))
                .TextMatrix(i, grd_ValidadCert) = nullToString(rsResult("ValidadCert"))
                .TextMatrix(i, grd_IdContenedor) = nullToString(rsResult("IdContenedor"))
                .TextMatrix(i, grd_NroPermEmbarque) = nullToString(rsResult("NroPermEmbarque"))
                .TextMatrix(i, grd_IdentificadorBulto) = nullToString(rsResult("IdentificadorBulto"))
                .TextMatrix(i, grd_Mercaderia) = nullToString(rsResult("Mercaderia"))
                .TextMatrix(i, grd_Peso) = nullToString(rsResult("Peso"))
            End With
            i = i + 1

            rsResult.MoveNext
        Wend

    Else
        MsgBox ("No hay datos para esta consulta"), vbInformation, Me.Caption
    End If


    If grdImpresines.Rows > 1 Then
        grdImpresines.Row = 1
    End If
    grdImpresines.Refresh
    grdImpresines.Visible = True
    Exit Sub
Catch:
        MsgBox " Error cargando la grilla de impresiones: " & err.Description, vbCritical, Me.Caption
        On Error GoTo 0
End Sub

'FUNCIONES PARA LIMPIAR Y CONFIGURAR CONTROLES
Public Sub ConfiguroGrilla()

    grdImpresines.Visible = False
    grdImpresines.Clear

    grdImpresines.SelectionMode = flexSelectionFree

    grdImpresines.ScrollBars = flexScrollBarBoth
    grdImpresines.AllowUserResizing = flexResizeColumns
    grdImpresines.FixedCols = 0
    grdImpresines.Cols = 15
    grdImpresines.Rows = 13

    grdImpresines.Row = 0
     

    grdImpresines.Col = grd_IdImpresion
    grdImpresines.Text = "Id Impresion"
    grdImpresines.ColWidth(grd_IdImpresion) = 0
    grdImpresines.ColAlignment(grd_IdImpresion) = 9

    grdImpresines.Col = grd_NroTk         '1
    grdImpresines.Text = "Nro Tk"
    grdImpresines.ColWidth(grd_NroTk) = 1400
    grdImpresines.ColAlignment(grd_NroTk) = 9

    grdImpresines.Col = grd_FechaHora         '2
    grdImpresines.Text = "Fecha Hora"
    grdImpresines.ColWidth(grd_FechaHora) = 1800
    grdImpresines.ColAlignment(grd_FechaHora) = 9

    grdImpresines.Col = grd_RazonSocial     '3
    grdImpresines.Text = "Razon Social"
    grdImpresines.ColWidth(grd_RazonSocial) = 0
    grdImpresines.ColAlignment(grd_RazonSocial) = 0

    grdImpresines.Col = grd_CUIT      '4
    grdImpresines.Text = "CUIT"
    grdImpresines.ColWidth(grd_CUIT) = 0
    grdImpresines.ColAlignment(grd_CUIT) = 9

    grdImpresines.Col = grd_CodigoAduana    '5
    grdImpresines.Text = "Codigo Aduana"
    grdImpresines.ColWidth(grd_CodigoAduana) = 0
    grdImpresines.ColAlignment(grd_CodigoAduana) = 9

    grdImpresines.Col = grd_LotPlanta       '6
    grdImpresines.Text = "Lot Planta"
    grdImpresines.ColWidth(grd_LotPlanta) = 0
    grdImpresines.ColAlignment(grd_LotPlanta) = 9

    grdImpresines.Col = grd_LotBalanza
    grdImpresines.Text = "Lot Balanza"
    grdImpresines.ColWidth(grd_LotBalanza) = 0
    grdImpresines.ColAlignment(grd_LotBalanza) = 9

    grdImpresines.Col = grd_Certificado    '8
    grdImpresines.Text = "Certificado"
    grdImpresines.ColWidth(grd_Certificado) = 1100
    grdImpresines.ColAlignment(grd_Certificado) = 9

    grdImpresines.Col = grd_ValidadCert         '9
    grdImpresines.Text = "ValidadCert"
    grdImpresines.ColWidth(grd_ValidadCert) = 1100
    grdImpresines.ColAlignment(grd_Certificado) = 9

    grdImpresines.Col = grd_NroPermEmbarque     '10
    grdImpresines.Text = "Nro Permso Embarque"
    grdImpresines.ColWidth(grd_NroPermEmbarque) = 1800
    grdImpresines.ColAlignment(grd_NroPermEmbarque) = 9
    
    grdImpresines.Col = grd_IdContenedor     '10
    grdImpresines.Text = "Id Contenedor"
    grdImpresines.ColWidth(grd_IdContenedor) = 1500
    grdImpresines.ColAlignment(grd_IdContenedor) = 9
    
    grdImpresines.Col = grd_IdentificadorBulto     '10
    grdImpresines.Text = "Identificador Bulto"
    grdImpresines.ColWidth(grd_IdentificadorBulto) = 2100
    grdImpresines.ColAlignment(grd_IdentificadorBulto) = 9
    
    grdImpresines.Col = grd_Mercaderia     '10
    grdImpresines.Text = "Mercaderia"
    grdImpresines.ColWidth(grd_Mercaderia) = 2500
    grdImpresines.ColAlignment(grd_Mercaderia) = 9
    
    grdImpresines.Col = grd_Peso     '10
    grdImpresines.Text = "Peso"
    grdImpresines.ColWidth(grd_Peso) = 2000
    grdImpresines.ColAlignment(grd_Peso) = 9

    grdImpresines.Col = 0

    grdImpresines.Visible = True
    
End Sub

Private Sub grdImpresines_Click()
    getTextos
End Sub
Private Sub LimpiarTextosImpresion()
        txtTicket.Text = ""
        txtCertificado.Text = ""
        txtValidadCert.Text = ""
        txtNroPermisoEmbarque.Text = ""
        txtIDContenedor.Text = ""
        txtIdentificadorBultoTxt.Text = ""
        txtMercaderia.Text = ""
        txtPeso.Text = ""
End Sub


Private Sub getTextos()
    On Error GoTo Catch
    LimpiarTextosImpresion
    
    Set objImpresionTmp = New clsImpresion

    If grdImpresines.Rows > 1 Then
'        If grdImpresines.TextMatrix(1, 1) = "" Then
'            Exit Sub
'        End If
        
        grdImpresines.Col = grd_IdImpresion
        objImpresionTmp.IdImpresion = grdImpresines.Text
        
        grdImpresines.Col = grd_NroTk
        objImpresionTmp.NroTk = grdImpresines.Text
        
        grdImpresines.Col = grd_FechaHora
        objImpresionTmp.FechaHora = grdImpresines.Text
        
        grdImpresines.Col = grd_RazonSocial
        objImpresionTmp.RazonSocial = grdImpresines.Text
        
        grdImpresines.Col = grd_CUIT
        objImpresionTmp.CUIT = grdImpresines.Text
        
        grdImpresines.Col = grd_CodigoAduana
        objImpresionTmp.CodigoAduana = grdImpresines.Text
        
        grdImpresines.Col = grd_LotPlanta
        objImpresionTmp.LotPlanta = grdImpresines.Text
        
        grdImpresines.Col = grd_LotBalanza
        objImpresionTmp.LotBalanza = grdImpresines.Text
        
        grdImpresines.Col = grd_Certificado
        objImpresionTmp.Certificado = grdImpresines.Text
        
        grdImpresines.Col = grd_ValidadCert
        objImpresionTmp.ValidadCert = grdImpresines.Text
        
        grdImpresines.Col = grd_NroPermEmbarque
        objImpresionTmp.NroPermEmbarque = grdImpresines.Text
        
        grdImpresines.Col = grd_IdContenedor
        objImpresionTmp.IdContenedor = grdImpresines.Text
        
        grdImpresines.Col = grd_IdentificadorBulto
        objImpresionTmp.IdentificadorBulto = grdImpresines.Text
        
        grdImpresines.Col = grd_Mercaderia
        objImpresionTmp.Mercaderia = grdImpresines.Text
        
        grdImpresines.Col = grd_Peso
        objImpresionTmp.Peso = grdImpresines.Text
        
        txtTicket.Text = objImpresionTmp.NroTk
        txtCertificado.Text = objImpresionTmp.Certificado
        txtValidadCert.Text = objImpresionTmp.ValidadCert
        txtNroPermisoEmbarque.Text = objImpresionTmp.NroPermEmbarque
        txtIDContenedor.Text = objImpresionTmp.IdContenedor
        txtIdentificadorBultoTxt.Text = objImpresionTmp.IdentificadorBulto
        txtMercaderia.Text = objImpresionTmp.Mercaderia
        txtPeso.Text = objImpresionTmp.Peso
    End If
    Exit Sub
Catch:
        'sgBox " Error cargando la grilla de impresiones: " & err.Description, vbCritical, Me.Caption
        'On Error GoTo 0
    ''On Error
    Resume Next
End Sub





Private Sub grdImpresines_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            ' Lógica para cuando se presiona la tecla de flecha hacia arriba
            getTextos
        Case vbKeyDown
            ' Lógica para cuando se presiona la tecla de flecha hacia abajo
            getTextos
'        Case vbKeyLeft
'            ' Lógica para cuando se presiona la tecla de flecha hacia la izquierda
'            MsgBox "Tecla de flecha izquierda presionada"
'        Case vbKeyRight
'            ' Lógica para cuando se presiona la tecla de flecha hacia la derecha
'            MsgBox "Tecla de flecha derecha presionada"
    End Select
End Sub

Private Sub grdImpresines_KeyPress(KeyAscii As Integer)
    getTextos
End Sub


Private Sub txtNrotkBusqueda_KeyPress(KeyAscii As Integer)
    ' Verificar si la tecla presionada es un número (0-9), Backspace (8) o Delete (127)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        ' Si no es un número o una tecla permitida, cancelar la entrada
        If Not (KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
End Sub
