Attribute VB_Name = "Ini"
Option Explicit
'####### base de datos
Public objdDB As clsAccesoDatos
'####### Clas
Public objImpresion As clsImpresion

Public colMercaderia As Collection
    


Dim strRrutaArchivo As String
'##########Variables de puertos
Public intCommPort As Integer
Public strCommSettings As String
Public intCommInputLen As Integer
Public intCommRThreshold As Integer
'##########Variables Parceo Mensaje
Public intPesoIni As Integer
Public intTaraIni As Integer

'##########Variable de Logueo
Public strLogDataReceiving As String
Public strLogImpresiones As String

''''##########Variable de TK
'''Public lngNroTk As Long
'''
''''##########Variable DataImpresion
'''Public strCertificado As String
'''Public strValidadCert As String
'''Public strMercaderia() As String


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Long

Public Function LeerINI(ByVal seccion As String, ByVal clave As String, ByVal rutaArchivo As String) As String
    Dim buffer As String * 256
    Dim longitud As Long
    
On Error GoTo Catch
    longitud = GetPrivateProfileString(seccion, clave, "", buffer, Len(buffer), rutaArchivo)
    LeerINI = Left(buffer, longitud)
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando CargarIni clave:" & clave & ", Err.Description:" & err.Description, vbCritical, "Error"
        'On Error GoTo 0
End Function

Public Function EscribirINI(ByVal seccion As String, ByVal clave As String, ByVal valor As String, ByVal rutaArchivo As String) As Boolean
    Dim resultado As Long
    
    resultado = WritePrivateProfileString(seccion, clave, valor, rutaArchivo)
    EscribirINI = (resultado <> 0)
End Function
Public Function EscribirINI2(ByVal seccion As String, ByVal clave As String, ByVal valor As String) As Boolean
    EscribirINI2 = EscribirINI(seccion, clave, valor, strRrutaArchivo)
End Function

Public Sub VerificarYCrearINI()
    Dim objFso As Object
    strRrutaArchivo = App.Path & "\config.ini"
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si el archivo INI existe
    If Not objFso.FileExists(strRrutaArchivo) Then
        ' Crear el archivo INI con la sección y clave predeterminadas
        Dim archivo As Object
        Set archivo = objFso.CreateTextFile(strRrutaArchivo, True)
        '##########Variables de puertos
        archivo.WriteLine ";Configurar para el puerto COM"
        archivo.WriteLine "[ConfigPuerto]"
        archivo.WriteLine ";Nro del Puerto Serie ej:3"
        archivo.WriteLine "CommPort=3"
        archivo.WriteLine ";Configurar la velocidad de transmisión, paridad, bits de datos y bits de parada.ej:9600,E,7,2"
        archivo.WriteLine "Settings=9600,E,7,2"
        archivo.WriteLine ";Configurar la cantidad que lee a la vez el:18 o 0"
        archivo.WriteLine "InputLen=0"
        archivo.WriteLine ";Establecer el umbral de recepción de bytes:18 0 1"
        archivo.WriteLine "RThreshold=1"
        '##########Variables Parceo Mensaje
        archivo.WriteLine ";Posicion donde comineza el peso en el mensaje"
        archivo.WriteLine "PesoIni=5"
        archivo.WriteLine ";Posicion donde comineza la tara en el mensaje"
        archivo.WriteLine "TaraIni=11"
        '##########Variable de Logueo
        archivo.WriteLine "[ConfigLog]"
        archivo.WriteLine "DataReceiving=N"
        archivo.WriteLine "LogImpresiones=N"
        
'''        '##########Variable de TK
'''        archivo.WriteLine "[ConfigTK]"
'''        archivo.WriteLine "NroTk=0"
'''
'''        '##########Variable DataImpresion
'''        archivo.WriteLine "[ConfigDataImpresion]"
'''        archivo.WriteLine "Certificado=307-43682"
'''        archivo.WriteLine "ValidadCert=29.01.2025"
'''        archivo.WriteLine "Mecaderia=SEMILLA DE CEBOLLA;SEMILLA DE MAIZ DULCE;SEMILLA DE MAIZ PISINGALLO;SEMILLA DE ZAPALLITO DE TRONCO;SEMILLA DE ZAPALLO;SEMILLA DE ZUCHINI"

        archivo.Close
    End If
    Set objFso = Nothing
End Sub

Public Function CargarIni() As Boolean
CargarIni = True

    

    On Error GoTo Catch
    VerificarYCrearINI
    ' Leer del archivo INI
    
    intCommPort = CInt(LeerINI("ConfigPuerto", "CommPort", strRrutaArchivo))
    strCommSettings = LeerINI("ConfigPuerto", "Settings", strRrutaArchivo)
    intCommInputLen = CInt(LeerINI("ConfigPuerto", "InputLen", strRrutaArchivo))
    intCommRThreshold = CInt(LeerINI("ConfigPuerto", "RThreshold", strRrutaArchivo))
    '##########Variables Parceo Mensaje
    intPesoIni = CInt(LeerINI("ConfigPuerto", "PesoIni", strRrutaArchivo))
    intTaraIni = CInt(LeerINI("ConfigPuerto", "TaraIni", strRrutaArchivo))
    ' Variable de Logueo
    strLogDataReceiving = LeerINI("ConfigLog", "DataReceiving", strRrutaArchivo)
    strLogImpresiones = LeerINI("ConfigLog", "LogImpresiones", strRrutaArchivo)
    
'''    '##########Variable de TK
'''    lngNroTk = CLng(LeerINI("ConfigTK", "NroTk", strRrutaArchivo))
'''
'''    '##########Variable DataImpresion
'''    strCertificado = LeerINI("ConfigDataImpresion", "Certificado", strRrutaArchivo)
'''    strValidadCert = LeerINI("ConfigDataImpresion", "ValidadCert", strRrutaArchivo)
'''    strMercaderia = Split(LeerINI("ConfigDataImpresion", "Mecaderia", strRrutaArchivo), ";")
    
        
Finally:
        Exit Function
Catch:
        MsgBox "Error cargando CargarIni, Err.Description:" & err.Description, vbCritical, "Error"
        On Error GoTo 0
        CargarIni = False
        GoTo Finally
End Function

'si un valor Double de la consulta es null lo convierto a 0
Public Function nullToDouble(Value As Variant) As Double
    On Error GoTo err
    nullToDouble = CDbl(Value)
    Exit Function
  
err:
    nullToDouble = 0
End Function
'si un valor String de la consulta es null lo convierto a ""
Public Function nullToString(Value As Variant) As String
    On Error GoTo err
    nullToString = CStr(Value)
    Exit Function
  
err:
    nullToString = ""
End Function
'si un valor Integer de la consulta es null lo convierto a 0
Public Function nullToInt(Value As Variant) As Integer
    On Error GoTo err
    nullToInt = CInt(Value)
    Exit Function
  
err:
    nullToInt = 0
End Function
'si un valor Long de la consulta es null lo convierto a 0
Public Function nullToLng(Value As Variant) As Long
    On Error GoTo err
    nullToLng = CLng(Value)
    Exit Function
  
err:
    nullToLng = 0
End Function
'si un valor Currency de la consulta es null lo convierto a 0
Public Function nullToCurr(Value As Variant) As Currency
    On Error GoTo err
    nullToCurr = CCur(Value)
    Exit Function
  
err:
    nullToCurr = 0
End Function
'si un valor Date de la consulta es null lo convierto a 0
Public Function nullToDate(Value As Variant) As Date
    On Error GoTo err
    nullToDate = Format(Value, "dd/mm/yyyy")
    Exit Function
  
err:
    nullToDate = 0
End Function
