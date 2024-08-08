VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private mIdMercaderia As Long
Private mMercaderia As String


'PROPIEDADES
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


Public Function ObtenerMercaderia(ByRef colMercaderia As Collection) As Boolean
    Dim objMerca As clsMercaderia
    ObtenerMercaderia = True
    On Error GoTo Catch
    Set colMercaderia = New Collection
    
    ' Abrir la conexión
    objdDB.AbrirConexion
    
    ' Cargo los datos del Cliente
    Dim rs As ADODB.Recordset
    Set rs = objdDB.EjecutarConsulta("SELECT * FROM Mercaderia")
    ' Procesar los resultados
    While Not rs.EOF
        Set objMerca = New clsMercaderia
        objMerca.IdMercaderia = rs.Fields("IdMercaderia").Value
        objMerca.Mercaderia = rs.Fields("Mercaderia").Value
        colMercaderia.Add objMerca
        rs.MoveNext
    Wend
    ' Cerrar la conexión
    objdDB.CerrarConexion
    'Set db = Nothing
    
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando en ObtenerMercaderia , Err.Description:" & Err.Description, vbCritical, "Error"
        On Error GoTo 0
        ObtenerMercaderia = False
        GoTo Finally
End Function

