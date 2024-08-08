VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mNroTk As Long
Private mIdImpresion As Long
Private mFechaHora As Date
Private mIdEmpresa As Long
Private mRazonSocial As String

Private mCUIT As String
Private mCodigoAduana As String
Private mLotPlanta As String
Private mLotBalanza As String

Private mCertificado As String
Private mValidadCert As String
Private mNroPermEmbarque As String
Private mIdContenedor As String
Private mIdentificadorBulto As String
Private mIdMercaderia As Long
Private mMercaderia As String
Private mPeso As Long




'PROPIEDADES


Public Property Let IdImpresion(Value As Long)
    mIdImpresion = Value
End Property

Public Property Get IdImpresion() As Long
    IdImpresion = mIdImpresion
End Property

Public Property Let NroTk(Value As Long)
    mNroTk = Value
End Property

Public Property Get NroTk() As Long
    NroTk = mNroTk
End Property


Public Property Let FechaHora(Value As Date)
    mFechaHora = Value
End Property

Public Property Get FechaHora() As Date
    FechaHora = mFechaHora
End Property

Public Property Let IdEmpresa(Value As Long)
    mIdEmpresa = Value
End Property

Public Property Get IdEmpresa() As Long
    IdEmpresa = mIdEmpresa
End Property


Public Property Let RazonSocial(Value As String)
    mRazonSocial = Value
End Property

Public Property Get RazonSocial() As String
    RazonSocial = mRazonSocial
End Property


Public Property Let CUIT(Value As String)
    mCUIT = Value
End Property

Public Property Get CUIT() As String
    CUIT = mCUIT
End Property

Public Property Let CodigoAduana(Value As String)
    mCodigoAduana = Value
End Property

Public Property Get CodigoAduana() As String
    CodigoAduana = mCodigoAduana
End Property


Public Property Let LotPlanta(Value As String)
    mLotPlanta = Value
End Property

Public Property Get LotPlanta() As String
    LotPlanta = mLotPlanta
End Property

Public Property Let LotBalanza(Value As String)
    mLotBalanza = Value
End Property

Public Property Get LotBalanza() As String
    LotBalanza = mLotBalanza
End Property


Public Property Let Certificado(Value As String)
    mCertificado = Value
End Property

Public Property Get Certificado() As String
    Certificado = mCertificado
End Property

Public Property Let ValidadCert(Value As String)
    mValidadCert = Value
End Property

Public Property Get ValidadCert() As String
    ValidadCert = mValidadCert
End Property


Public Property Let NroPermEmbarque(Value As String)
    mNroPermEmbarque = Value
End Property

Public Property Get NroPermEmbarque() As String
    NroPermEmbarque = mNroPermEmbarque
End Property

Public Property Let IdContenedor(Value As String)
    mIdContenedor = Value
End Property

Public Property Get IdContenedor() As String
    IdContenedor = mIdContenedor
End Property

Public Property Let IdentificadorBulto(Value As String)
    mIdentificadorBulto = Value
End Property

Public Property Get IdentificadorBulto() As String
    IdentificadorBulto = mIdentificadorBulto
End Property

Public Property Let IdMercaderia(Value As Long)
    mIdMercaderia = Value
End Property

Public Property Get IdMercaderia() As Long
    IdMercaderia = mIdMercaderia
End Property


Public Property Let Mercaderia(Value As String)
    mMercaderia = Value
End Property

Public Property Get Mercaderia() As String
    Mercaderia = mMercaderia
End Property

Public Property Let Peso(Value As Long)
    mPeso = Value
End Property

Public Property Get Peso() As Long
    Peso = mPeso
End Property

'METODOS
Private Sub Class_Initialize()
    mNroTk = 0
    mIdEmpresa = 0
    mRazonSocial = ""
    mCUIT = ""
    mCertificado = ""
    mValidadCert = ""
End Sub
'    Limpiar
'End Sub

Private Sub Limpiar()
    mNroTk = 0
    mFechaHora = Now
    mIdEmpresa = 0
    mRazonSocial = ""
    mCUIT = ""
    mCodigoAduana = ""
    mLotPlanta = ""
    mCertificado = ""
    mValidadCert = ""
    mNroPermEmbarque = ""
    mIdContenedor = ""
    mIdentificadorBulto = ""
    mIdMercaderia = 0
    mMercaderia = ""
    mPeso = 0
End Sub

Public Function CargarDatosBasicos() As Boolean
    
    CargarDatosBasicos = True
    On Error GoTo Catch
    
    
    ' Abrir la conexión
    objdDB.AbrirConexion
    

    ' Cargo los datos del Cliente
    Dim rs As ADODB.Recordset
    Set rs = objdDB.EjecutarConsulta("SELECT * FROM Empresa")
    ' Procesar los resultados
    While Not rs.EOF
        mIdEmpresa = rs.Fields("IdEmpresa").Value
        mRazonSocial = rs.Fields("RazonSocial").Value
        mCUIT = rs.Fields("CUIT").Value
        rs.MoveNext
    Wend
    ' Cerrar el recordset
    rs.Close
        

    
    'Cargo los semipermanentes permanentes del TK
    Set rs = objdDB.EjecutarConsulta("SELECT * FROM Parametros where CodParametro in('Certificado','ValidadCert','CodigoAduana','LotPlanta','LotBalanza') ")
    While Not rs.EOF
        
        Select Case rs.Fields("CodParametro").Value
            Case "Certificado"
                 mCertificado = rs.Fields("ValorText").Value
            Case "ValidadCert"
                 mValidadCert = rs.Fields("ValorText").Value
            Case "CodigoAduana"
                 mCodigoAduana = rs.Fields("ValorText").Value
            Case "LotPlanta"
                 mLotPlanta = rs.Fields("ValorText").Value
            Case "LotBalanza"
                 mLotBalanza = rs.Fields("ValorText").Value
            Case Else

        End Select
        rs.MoveNext
        
    Wend
     
        
    'obtengo ep priximo Nro de tk
            
            
    Set rs = objdDB.EjecutarConsulta("SELECT * FROM Cons_UltimoNroTk")
    While Not rs.EOF
        If IsNull(rs.Fields("Ultimo").Value) Then
            mNroTk = 1
        Else
            mNroTk = rs.Fields("Ultimo").Value + 1
           
        End If
        rs.MoveNext
    Wend
            
    Set rs = Nothing

    ' Cerrar la conexión
    objdDB.CerrarConexion
    'Set db = Nothing
    
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando en CargarDatosBasicos , Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        CargarDatosBasicos = False
        GoTo Finally
End Function


Public Function ActualizaParametro(strCodParametro As String, strVatorText As String) As Boolean
    
    ActualizaParametro = True
    Dim strSql As String
    On Error GoTo Catch
        
    ' Abrir la conexión
    objdDB.AbrirConexion
    strSql = "Update Parametros set ValorText='" & strVatorText & "' where CodParametro='" & strCodParametro & "'"
    objdDB.EjecutarConsulta (strSql)
    ' Cerrar la conexión
    objdDB.CerrarConexion
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando en ActualizaParametro " & strCodParametro & ", Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        ActualizaParametro = False
        GoTo Finally
End Function


Public Function InsertarImpresion() As Boolean
    
    InsertarImpresion = True
    Dim strSql As String
    On Error GoTo Catch
        
    ' Abrir la conexión
    objdDB.AbrirConexion
    strSql = "INSERT INTO Impresiones ("
    strSql = strSql & "NroTk,"
    strSql = strSql & "FechaHora,"
    strSql = strSql & "IdEmpresa,"
    strSql = strSql & "CodigoAduana,"
    strSql = strSql & "LotPlanta,"
    strSql = strSql & "LotBalanza,"
    strSql = strSql & "Certificado,"
    strSql = strSql & "ValidadCert,"
    strSql = strSql & "NroPermEmbarque,"
    strSql = strSql & "IdContenedor,"
    strSql = strSql & "IdentificadorBulto,"
    strSql = strSql & "IdMercaderia,"
    strSql = strSql & "Peso) VALUES ("
    strSql = strSql & mNroTk & ","
    strSql = strSql & "'" & Format(mFechaHora, "yyyy-mm-dd hh:nn:ss") & "',"
    strSql = strSql & mIdEmpresa & ",'"
    strSql = strSql & mCodigoAduana & "','"
    strSql = strSql & mLotPlanta & "','"
    strSql = strSql & mLotBalanza & "','"
    strSql = strSql & mCertificado & "','"
    strSql = strSql & mValidadCert & "','"
    strSql = strSql & mNroPermEmbarque & "','"
    strSql = strSql & mIdContenedor & "','"
    strSql = strSql & mIdentificadorBulto & "',"
    strSql = strSql & mIdMercaderia & ","
    strSql = strSql & mPeso & ")"

    objdDB.EjecutarConsulta (strSql)
    ' Cerrar la conexión
    objdDB.CerrarConexion
Finally:
        Exit Function
Catch:
        MsgBox "Error guardando la  Impresion, Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        InsertarImpresion = False
        GoTo Finally
End Function

Public Function ObtenerImpresiones(ByRef rs As ADODB.Recordset, strSql As String) As Boolean
    Dim objMerca As clsMercaderia
    ObtenerImpresiones = True
    On Error GoTo Catch
     
    ' Abrir la conexión
    objdDB.AbrirConexion
    
    Set rs = objdDB.EjecutarConsulta(strSql)
    ' Cerrar la conexión
    'objdDB.CerrarConexion
    'Set db = Nothing
    
Finally:
        ' Cerrar la conexión
        
        Exit Function
Catch:
        MsgBox "Error en ObtenerImpresiones , Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        ObtenerImpresiones = False
        objdDB.CerrarConexion
        GoTo Finally
End Function

Public Function GenerarImpresion() As Boolean
    GenerarImpresion = True
    
    ' Crear un Recordset en memoria
    Dim strPaso As String
    
    strPaso = "Crear un Recordset"
    Dim rs As ADODB.Recordset
    strPaso = "New Recordset"
    Set rs = New ADODB.Recordset
    strPaso = "SET Recordset"
    With rs
         strPaso = "Recordset Crea los campos"
        .Fields.Append "RazonSocial", adVarChar, 25
        .Fields.Append "CUIT", adVarChar, 25

        .Fields.Append "CodigoAduana", adVarChar, 10
        .Fields.Append "LotPlanta", adVarChar, 10
        .Fields.Append "LotBalanza", adVarChar, 20

        .Fields.Append "Certificado", adVarChar, 20
        .Fields.Append "ValidadCert", adVarChar, 20

        .Fields.Append "NroPermEmbarque", adVarChar, 255
        .Fields.Append "IdContenedor", adVarChar, 30
        .Fields.Append "IdentificadorBulto", adVarChar, 30
        .Fields.Append "DescMercaderia", adVarChar, 255
        .Fields.Append "Peso", adVarChar, 255

        .Open
        ' Agregar datos manualmente
        strPaso = "Carga Recordset"
        .AddNew
        .Fields("RazonSocial").Value = RazonSocial
        .Fields("CUIT").Value = CUIT


        .Fields("CodigoAduana").Value = CodigoAduana
        .Fields("LotPlanta").Value = LotPlanta
        .Fields("LotBalanza").Value = LotBalanza

        .Fields("Certificado").Value = Certificado 'cambia
        .Fields("ValidadCert").Value = ValidadCert 'cambia
        .Fields("NroPermEmbarque").Value = NroPermEmbarque 'cambia
        .Fields("IdContenedor").Value = IdContenedor 'cambia
        .Fields("IdentificadorBulto").Value = IdentificadorBulto 'cambia
        .Fields("DescMercaderia").Value = Mercaderia  'cambia
        .Fields("Peso").Value = Peso 'cambia

        .Update
    End With

    strPaso = "Pasa parametro Sección4"
    dtrImpresion.Sections("Sección4").Controls("lblFecha").Caption = FechaHora
    strPaso = "Pasa parametro Sección2"
    dtrImpresion.Sections("Sección2").Controls("lblTicket").Caption = "TICKET " & Format$(NroTk, "0000000000") ' modifica


    strPaso = "Asignar el Recordset al DataReport"
    Set dtrImpresion.DataSource = rs
    
    strPaso = "Mostrar el DataReport"
    dtrImpresion.Show vbModal
    
Finally:
    Set rs = Nothing
    Exit Function
Catch:
    MsgBox "Error en GenerarImpresion paso:" & strPaso & " , Err.Description:" & err.Description, vbCritical, "Error"
    On Error GoTo 0
    GenerarImpresion = False
    GoTo Finally
End Function


Public Function ObtenerProximoNroTk() As Boolean
    ObtenerProximoNroTk = True
    On Error GoTo Catch
    
    
    ' Abrir la conexión
    objdDB.AbrirConexion
    ' Cargo los datos del Cliente
    Dim rs As ADODB.Recordset
    'obtengo ep priximo Nro de tk
            
    Set rs = objdDB.EjecutarConsulta("SELECT * FROM Cons_UltimoNroTk")
    While Not rs.EOF
        If IsNull(rs.Fields("Ultimo").Value) Then
            mNroTk = 1
        Else
            mNroTk = rs.Fields("Ultimo").Value + 1
           
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing

    ' Cerrar la conexión
    objdDB.CerrarConexion
    'Set db = Nothing
    
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando en ObtenerProximoNroTk , Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        ObtenerProximoNroTk = False
        GoTo Finally
End Function



    
    




