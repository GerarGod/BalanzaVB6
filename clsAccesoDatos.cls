VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccesoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declaración de objetos ADO
Private m_cnx As ADODB.Connection
Private m_rs As ADODB.Recordset

' Cadena de conexión a la base de datos
Private CADENA_CONEXION As String

Private Sub Class_Initialize()
    'CADENA_CONEXION = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & VB.App.Path & "\DB_BASSO.mdb;"
    
    CADENA_CONEXION = "Provider=Microsoft.ACE.OLEDB.12.0 ;Data Source=" & VB.App.Path & "\DB_BASSO.accdb; Jet OLEDB:Database Password=gerargod;"
    
End Sub


' Método para abrir la conexión
Public Sub AbrirConexion()
    Set m_cnx = New ADODB.Connection
    m_cnx.ConnectionString = CADENA_CONEXION
    m_cnx.Open
End Sub

' Método para cerrar la conexión
Public Sub CerrarConexion()
    If Not m_cnx Is Nothing Then
        If m_cnx.State = adStateOpen Then
            m_cnx.Close
        End If
        Set m_cnx = Nothing
    End If
End Sub

' Método para ejecutar una consulta SQL
Public Function EjecutarConsulta(ByVal sql As String) As ADODB.Recordset
    Set m_rs = New ADODB.Recordset
    m_rs.Open sql, m_cnx, adOpenStatic, adLockReadOnly
    Set EjecutarConsulta = m_rs
End Function

' Método para ejecutar un comando SQL (INSERT, UPDATE, DELETE)
Public Sub EjecutarComando(ByVal sql As String)
    If m_cnx Is Nothing Then
        err.Raise vbObjectError + 9999, "AccesoDatos", "Conexión no abierta."
    End If
    m_cnx.Execute sql
End Sub
