Attribute VB_Name = "librerias"
Public MiConexion As ADODB.Connection
Public MiCommand As ADODB.Command
Public oracle As Boolean
Public conectado As Boolean

Public Sub conecta()
    
    'Variables de conexión
    Dim cadenaConexion As String
    Dim path As String
    
    On Error GoTo error_de_conexion
   
    conectado = False
    If oracle Then
        cadenaConexion = "Provider=MSDAORA.1;Password=jMWW245a;User ID=lsi03;Data Source=labora; Persist Security Info=False"
    Else
        path = "BBDD.mdb" 'evidentemente aqui tiene que escribir cada uno su PATH
        cadenaConexion = "Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path
    End If
    
    Set MiConexion = CreateObject("ADODB.Connection")
    MiConexion.CursorLocation = adUseClient
    MiConexion.Open cadenaConexion

    Set MiCommand = CreateObject("ADODB.Command")
    MiCommand.ActiveConnection = MiConexion
    conectado = True
    
error_de_conexion:
    'Se devuelve el control al modulo identificacion
    
End Sub

Public Sub crea_record(rs As Recordset, cadena As Variant, depurar As Boolean)
    
    If depurar = True Then MsgBox ("Ejectutando SQL: //" & cadena & "//")
    MiCommand.CommandType = adCmdText
    MiCommand.CommandText = CStr(cadena)
    Set rs = MiCommand.Execute
End Sub

Public Sub desconecta()
    If conectado = True Then
        MiConexion.Close
    End If
End Sub

