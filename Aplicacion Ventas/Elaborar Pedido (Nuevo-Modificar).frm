VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Elaborar_pedido_nuevo_modif 
   ClientHeight    =   8490
   ClientLeft      =   3570
   ClientTop       =   1530
   ClientWidth     =   8790
   Icon            =   "Elaborar Pedido (Nuevo-Modificar).frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bIncidencia 
      Caption         =   "Incidencia Pedido"
      Height          =   375
      Left            =   4680
      TabIndex        =   44
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Pedido"
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   8535
      Begin VB.TextBox cFechaElabPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cCodPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cCodRepresOperadora 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox cDireccion 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox cNumero 
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox cPuerta 
         Height          =   285
         Left            =   5880
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox cLocalidad 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox cProvincia 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox cPais 
         Height          =   285
         Left            =   6600
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox cCP 
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo Representante/Operadora"
         Height          =   255
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección Envío"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Nº"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Pta"
         Height          =   255
         Left            =   5880
         TabIndex        =   25
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "País"
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "CP"
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton bGuardar 
      Caption         =   "&Guardar "
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton bEnviarAlmacen 
      Caption         =   "Enviar &a almacén"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   8535
      Begin VB.CommandButton bEliminarLinea 
         Caption         =   "&Eliminar Línea"
         Height          =   375
         Left            =   3360
         TabIndex        =   39
         Top             =   1920
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "referencia"
            Caption         =   "Referencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "precio"
            Caption         =   "Precio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2415,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005,165
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton bConsultaCat 
         Caption         =   "&Consultar Catálogo"
         Height          =   375
         Left            =   6120
         TabIndex        =   33
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton bAñadir 
         Caption         =   "&Añadir Línea"
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nueva línea"
         Height          =   975
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   7815
         Begin VB.TextBox cDescripcion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   37
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox cPrecioUnidad 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6240
            TabIndex        =   36
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox cCantidad 
            Height          =   285
            Left            =   4800
            TabIndex        =   35
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox cCod_Articulo 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Precio"
            Height          =   255
            Left            =   6240
            TabIndex        =   43
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   4800
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   1560
            TabIndex        =   41
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Referencia"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         ItemData        =   "Elaborar Pedido (Nuevo-Modificar).frx":27A2
         Left            =   1320
         List            =   "Elaborar Pedido (Nuevo-Modificar).frx":27A4
         TabIndex        =   18
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox cTotal 
         Height          =   285
         Left            =   6960
         TabIndex        =   16
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox cIVA 
         Height          =   285
         Left            =   6960
         TabIndex        =   14
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   4080
         Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":27A6
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1440
         Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":2BE8
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6960
         Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":302A
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label17 
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Total"
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "IVA"
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Top             =   4320
         Width           =   375
      End
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   5280
      Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":346C
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   7560
      Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":38AE
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3120
      Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":3CF0
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   840
      Picture         =   "Elaborar Pedido (Nuevo-Modificar).frx":4132
      Top             =   7440
      Width           =   480
   End
End
Attribute VB_Name = "Elaborar_Pedido_Nuevo_Modif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim espedidonuevo As Boolean
Dim codcli As String

'si se trata de un pedido nuevo, los datos de la direccion deben ser
'los propios del cliente. Tambien se pone la fecha del sistema
Sub PonDireccionEnvio(cliente As String)
    espedidonuevo = True
    Me.cDireccion.Text = Elaborar_pedido_CU02.cCalle.Text
    Me.cNumero.Text = Elaborar_pedido_CU02.cNum.Text
    Me.cPuerta.Text = Elaborar_pedido_CU02.cPta.Text
    Me.cLocalidad.Text = Elaborar_pedido_CU02.cLocalidad.Text
    Me.cProvincia.Text = Elaborar_pedido_CU02.cProvincia.Text
    Me.cPais.Text = Elaborar_pedido_CU02.cPais.Text
    Me.cCP.Text = Elaborar_pedido_CU02.cCP.Text
    
    Me.cFechaElabPedido.Text = Date
    
    'tambien se le ha de dar ya un codigo al pedido
    Dim rs As ADODB.Recordset
    Dim rsnew As ADODB.Recordset
    Dim consulta As String
        
    consulta = "SELECT max(codigo_pedido) as ultimo FROM Orden_pedido"
    crea_record rs, consulta, False
    Me.cCodPedido.Text = rs("ultimo") + 1
    'la forma de pago dependera del ratio de confianza del cliente
    FormaPago cliente
    If oracle Then
        consulta = "INSERT INTO orden_pedido (codigo_pedido,cliente,usuario_ventas,calle_envio,numero_envio,pta_envio,localidad_envio,provincia_envio,pais_envio,cp_envio,forma_pago,fecha_elaboracion) VALUES (" & Me.cCodPedido.Text & ",'" & cliente & "','" & Me.cCodRepresOperadora.Text & "','" & Me.cDireccion.Text & "'," & Me.cNumero.Text & "," & Me.cPuerta.Text & ",'" & Me.cLocalidad.Text & "','" & Me.cProvincia.Text & "','" & Me.cPais.Text & "'," & Me.cCP.Text & ",'" & Me.cmbFormaPago.ListIndex & "', TO_Date( '" & Me.cFechaElabPedido.Text & "', 'DD/MM/YYYY HH:MI:SS AM'))"
    Else
        consulta = "INSERT INTO orden_pedido (codigo_pedido,cliente,usuario_ventas,calle_envio,numero_envio,pta_envio,localidad_envio,provincia_envio,pais_envio,cp_envio,forma_pago,fecha_elaboracion) VALUES (" & Me.cCodPedido.Text & ",'" & cliente & "','" & Me.cCodRepresOperadora.Text & "','" & Me.cDireccion.Text & "'," & Me.cNumero.Text & "," & Me.cPuerta.Text & ",'" & Me.cLocalidad.Text & "','" & Me.cProvincia.Text & "','" & Me.cPais.Text & "'," & Me.cCP.Text & ",'" & Me.cmbFormaPago.ListIndex & "',#" & Me.cFechaElabPedido.Text & "#)"
    End If
    
    'MsgBox consulta
    MiCommand.CommandText = consulta
    MiCommand.Execute
End Sub


'******************** CLASE CLIENTE: calcular_forma_pago ***************************
Sub FormaPago(cliente As String)
    Dim rscli As ADODB.Recordset
    Dim consulta As String
    
    codcli = cliente
    consulta = "SELECT * FROM cliente WHERE codigo='" & cliente & "'"
    crea_record rscli, consulta, False
    Select Case rscli("ratio_confianza")
        Case "bueno", "excelente", "Bueno", "Excelente" 'si es bueno o excelente puede elegir
            Me.cmbFormaPago.Clear 'limpiar el contenido por si habia algo
            Me.cmbFormaPago.AddItem "Al contado", 0
            Me.cmbFormaPago.AddItem "A crédito", 1
            Me.cmbFormaPago.ListIndex = 0 'al contado
        Case Else 'en cualquier otro caso no podra elegir
            Me.cmbFormaPago.Clear
            Me.cmbFormaPago.AddItem "Al contado", 0
            Me.cmbFormaPago.ListIndex = 0
    End Select
'****************************** fin calcular_forma_pago ****************************
End Sub


'**************** CLASE ORDEN DE PEDIDO: consultar_orden_de_pedido *****************
'se va a modificar un pedido y con este metodo se cargan los datos del pedido
Sub Cargar_Pedido(cod_ped As String)
    espedidonuevo = False
    Dim rsped As ADODB.Recordset
    Dim consulta As String
    Dim consulta2 As String

    MiConexion.BeginTrans
     
    consulta = "SELECT * FROM orden_pedido WHERE codigo_pedido=" & cod_ped & ""
    consulta2 = "SELECT linea_pedido.codigo_pedido, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad,linea_pedido.precio,linea_pedido.cantidad * linea_pedido.precio as total From linea_pedido, producto where codigo_pedido=" & cod_ped & " and producto.referencia=linea_pedido.referencia"
    Set rslinped = New ADODB.Recordset
    rslinped.Open consulta2, MiConexion, adOpenDynamic, adLockOptimistic
    crea_record rsped, consulta, False
    Set Me.DataGrid1.DataSource = rslinped
    
    Me.cCodPedido.Text = cod_ped
    Me.cCodRepresOperadora.Text = rsped("usuario_ventas")
    Me.cFechaElabPedido.Text = rsped("fecha_elaboracion")
    Me.cDireccion.Text = rsped("calle_envio")
    Me.cNumero.Text = rsped("numero_envio")
    Me.cPuerta.Text = rsped("pta_envio")
    Me.cLocalidad.Text = rsped("localidad_envio")
    Me.cProvincia.Text = rsped("provincia_envio")
    Me.cPais.Text = rsped("pais_envio")
    Me.cCP.Text = rsped("cp_envio")
    Me.FormaPago rsped("cliente")
    Me.cmbFormaPago.ListIndex = rsped("forma_pago")
'**************************** fin consultar_orden_de_pedido ************************
    calcula_ivatotal cod_ped
    
End Sub

Sub calcula_ivatotal(cod_ped As String)
    Dim rs As ADODB.Recordset
    Dim cons As String
    cons = "SELECT sum(linea_pedido.precio * cantidad) as total From linea_pedido where codigo_pedido=" & cod_ped & ""
    crea_record rs, cons, False
    Me.cIVA.Text = rs("total") * 0.16
    Me.cTotal.Text = rs("total") + (rs("total") * 0.16)
    'limpiamos los campos de nueva línea
    Me.cCod_Articulo.Text = ""
    Me.cDescripcion.Text = ""
    Me.cCantidad.Text = ""
    Me.cPrecioUnidad.Text = ""
End Sub

Private Sub bAñadir_Click()
    Dim rs As ADODB.Recordset
    Dim cant As Long
    
    If Me.cCod_Articulo.Text = "" Or Me.cCantidad.Text = "" Or Me.cPrecioUnidad.Text = "" Then
        MsgBox "Debe rellenar obligatoriamente el campo de referencia, la cantidad y el precio unitario."
    Else
        'añadir la linea al grid
        'si ya existe la referencia, se actualiza el registro con la nueva cantidad
        'si no existe, pues se inserta
        'si es un pedido nuevo, el datagrid aun no tiene un recordset asociado
        Dim rsprod As ADODB.Recordset
        MiCommand.CommandText = "SELECT * FROM producto WHERE referencia='" & Me.cCod_Articulo.Text & "'"
        Set rsprod = MiCommand.Execute
        
        cant = CLng(Me.cCantidad.Text)
        If cant <= 0 Then
            MsgBox "La cantidad pedida debe ser mayor que cero."
            Me.cCantidad.SetFocus
        Else
            If cant > rsprod("max_razonable") Then
                MsgBox "La cantidad pedida excede el máximo razonable para este producto, que es de " & CStr(rsprod("max_razonable")) & " unidades."
                Me.cCantidad.SetFocus
            Else
                On Error GoTo gridsinrecordset
                    Set rs = Me.DataGrid1.DataSource
                    rs.Find "referencia='" & Me.cCod_Articulo.Text & "'", , , 1
                    If Not rs.EOF Then 'existe ya una linea con ese producto
                    If MsgBox("Ya existe una línea con esa referencia, ¿desea sobreescribirla?", vbOKCancel, "Atención") = vbOK Then
                        On Error GoTo valor_cantidad
                        Me.DataGrid1.Row = CInt(rs.Bookmark) - 1
                        rs("cantidad").Value = cant
                        rs.Update
                    End If
                    Else 'es una linea nueva
                        'ya no hay q comprobar q exista la referencia
                        'pq ya se hace cuando se rellenan los datos de la nueva línea
                        'If rsprod.RecordCount = 1 Then
                        'el pedido ya debe existir en la bbdd para poder insertar una línea
                        MiCommand.CommandText = "INSERT INTO linea_pedido (codigo_pedido,referencia,cantidad,precio) VALUES (" & Me.cCodPedido.Text & ",'" & Me.cCod_Articulo.Text & "'," & Me.cCantidad.Text & "," & Me.cPrecioUnidad.Text & ")"
                        MiCommand.Execute
                    End If
                    rs.Requery
                    Me.calcula_ivatotal Me.cCodPedido.Text
            End If
        End If
    End If
    Exit Sub
    
gridsinrecordset:  'el grid no tenia recordset asignado todavia
            'el pedido ya debe existir bbdd para poder insertar una línea
            crea_record rs, "SELECT linea_pedido.codigo_pedido, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad,linea_pedido.precio,linea_pedido.cantidad * linea_pedido.precio as total From linea_pedido, producto where codigo_pedido=" & Me.cCodPedido.Text & " and producto.referencia=linea_pedido.referencia", False
            Set Me.DataGrid1.DataSource = rs
            If rs.RecordCount < 1 Then
                MiCommand.CommandText = "INSERT INTO linea_pedido (codigo_pedido,referencia,cantidad,precio) VALUES (" & Me.cCodPedido.Text & ",'" & Me.cCod_Articulo.Text & "'," & Me.cCantidad.Text & "," & Me.cPrecioUnidad.Text & ")"
                MiCommand.Execute
                'vuelvo a crear el recordset para reasignar la consulta sql
                crea_record rs, "SELECT linea_pedido.codigo_pedido, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad,linea_pedido.precio,linea_pedido.cantidad * linea_pedido.precio as total From linea_pedido, producto where codigo_pedido=" & Me.cCodPedido.Text & " and producto.referencia=linea_pedido.referencia", False
                Set Me.DataGrid1.DataSource = rs
            End If
            Me.calcula_ivatotal Me.cCodPedido.Text
            Exit Sub
            
valor_cantidad:
            Dim sql As String
            Dim rslinped As ADODB.Recordset
            
            '************** CLASE ORDEN DE PEDIDO: calcular_precio_total ****************
            sql = "SELECT linea_pedido.codigo_pedido, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad,linea_pedido.precio,linea_pedido.cantidad * linea_pedido.precio as total From linea_pedido, producto where codigo_pedido=" & cCodPedido & " and producto.referencia=linea_pedido.referencia"
            Set rslinped = New ADODB.Recordset
            rslinped.Open sql, MiConexion, adOpenDynamic, adLockOptimistic
            rslinped.Bookmark = rs.Bookmark
            Set Me.DataGrid1.DataSource = rslinped
            
            rslinped("cantidad").Value = CLng(Me.cCantidad.Text)
            rslinped.Update
            rslinped.Requery
            '************************** fin calcular_precio_total ***********************
            Me.calcula_ivatotal Me.cCodPedido.Text


End Sub

Private Sub bConsultaCat_Click()
    Me.Visible = False
    Consultar_catalogo.Show
End Sub


'******************* CLASE LINEA DE PEDIDO: eliminar_línea_de_pedido ***************
Private Sub bEliminarLinea_Click()
    Dim rs As ADODB.Recordset
    Dim cons As String
    
    'si queda solo una linea no se puede eliminar
    Set rs = Me.DataGrid1.DataSource
    On Error GoTo gridsinrs
        If rs.RecordCount <= 1 Then
            MsgBox "No se puede eliminar la línea del pedido porque es la única que hay."
        Else
            If MsgBox("¿Está seguro que desea eliminar esa línea del pedido?", vbOKCancel) = 1 Then
                Me.DataGrid1.AllowDelete = True
                Me.DataGrid1.Col = 0
                cons = "DELETE FROM linea_pedido WHERE codigo_pedido=" & Me.cCodPedido.Text & " AND referencia='" & Me.DataGrid1.Text & "'"
                MiCommand.CommandText = cons
                rs.CancelUpdate
                MiCommand.Execute
                rs.Requery
                calcula_ivatotal Me.cCodPedido.Text
            End If
        End If
        Exit Sub
gridsinrs:
    MsgBox "No se puede eliminar la línea del pedido porque es la única que hay."
'************************ fin eliminar_línea_de_pedido *****************************
End Sub


'****************** CLASE ORDEN DE PEDIDO: modificar_orden_de_pedido ***************
'este metodo envia un pedido al almacen correspondiente a la region a la q pertenece el pais de envio
'avisa de q se ha enviado el pedido al almacen, actualiza los datos y cierra el form
Private Sub bEnviarAlmacen_Click()
    'sea nuevo o modificado, se han de guardar los datos de envio y la forma de pago,
    'por si han sido modificados
    Dim rs As ADODB.Recordset
    Set rs = Me.DataGrid1.DataSource
    On Error GoTo gridsinrs
    If rs.RecordCount >= 1 Then
        On Error GoTo falloenvio
            If oracle Then
                MiCommand.CommandText = "UPDATE orden_pedido SET fecha_llegada_almacen=TO_Date( '" & Date & "', 'DD/MM/YYYY HH:MI:SS AM'),forma_pago='" & Me.cmbFormaPago.ListIndex & "',calle_envio='" & Me.cDireccion.Text & "',numero_envio=" & Me.cNumero.Text & ",pta_envio=" & Me.cPuerta.Text & ",localidad_envio='" & Me.cLocalidad.Text & "',provincia_envio='" & Me.cProvincia.Text & "',pais_envio='" & Me.cPais.Text & "',cp_envio=" & Me.cCP.Text & " WHERE codigo_pedido=" & Me.cCodPedido.Text
            Else
                MiCommand.CommandText = "UPDATE orden_pedido SET fecha_llegada_almacen=#" & Date & "#,forma_pago='" & Me.cmbFormaPago.ListIndex & "',calle_envio='" & Me.cDireccion.Text & "',numero_envio=" & Me.cNumero.Text & ",pta_envio=" & Me.cPuerta.Text & ",localidad_envio='" & Me.cLocalidad.Text & "',provincia_envio='" & Me.cProvincia.Text & "',pais_envio='" & Me.cPais.Text & "',cp_envio=" & Me.cCP.Text & " WHERE codigo_pedido=" & Me.cCodPedido.Text
            End If
            MiCommand.Execute
            Unload Me
    End If
    If espedidonuevo = False Then
        MiConexion.CommitTrans
    End If
    Exit Sub
    
falloenvio:
        MsgBox "Se produjo un error y el pedido no ha podido ser enviado al almacén."
        Exit Sub
    
gridsinrs:
        MsgBox "El pedido no puede estar vacío. Ha de tener al menos una línea."
End Sub

Private Sub bGuardar_Click()
    'sea nuevo o modificado, se han de guardar los datos de envio y la forma de pago,
    'por si han sido modificados
    Dim rs As ADODB.Recordset
    Set rs = Me.DataGrid1.DataSource
    On Error GoTo gridsinrs
    If rs.RecordCount >= 1 Then
        MiCommand.CommandText = "UPDATE orden_pedido SET forma_pago='" & Me.cmbFormaPago.ListIndex & "',calle_envio='" & Me.cDireccion.Text & "',numero_envio=" & Me.cNumero.Text & ",pta_envio=" & Me.cPuerta.Text & ",localidad_envio='" & Me.cLocalidad.Text & "',provincia_envio='" & Me.cProvincia.Text & "',pais_envio='" & Me.cPais.Text & "',cp_envio=" & Me.cCP.Text & " WHERE codigo_pedido=" & Me.cCodPedido.Text
        MiCommand.Execute
        Unload Me
    End If
    If espedidonuevo = False Then
        MiConexion.CommitTrans
    End If
    Exit Sub
    
gridsinrs:
        MsgBox "El pedido no puede estar vacío. Ha de tener al menos una línea."

'***************************** fin modificar_orden_de_pedido ***********************

End Sub

Private Sub bIncidencia_Click()
        Me.Visible = False
        Incidencia_pedido.cargar_datos_incidencia Me.cCodPedido
        Incidencia_pedido.Show vbModal
End Sub

Private Sub bSalir_Click()
    If espedidonuevo = True Then
        MiCommand.CommandText = "DELETE FROM orden_pedido WHERE codigo_pedido=" & Me.cCodPedido.Text
        MiCommand.Execute
    End If
    If espedidonuevo = False Then
        MiConexion.RollbackTrans
    End If
    Unload Me
End Sub
'este metodo comprueba que el numero metido no es negativo
Private Function compruebaNum(num As Long) As Boolean
    If num >= 1 Then
        compruebaNum = True 'es positivo
    Else
        compruebaNum = False 'el num era negativo
        MsgBox "El número que acaba de introducir no puede ser negativo."
    End If
End Function

Private Sub cCantidad_LostFocus()
    On Error GoTo noesnum
        If Not compruebaNum(CLng(Me.cCantidad.Text)) Then
            Me.cCantidad.SetFocus
        End If
        Exit Sub
noesnum:
    MsgBox "Debe introducir un número entero positivo."
    'Me.cCantidad.SetFocus
End Sub

Private Sub cCod_Articulo_LostFocus()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    If Me.cCod_Articulo.Text <> "" Then
        consulta = "SELECT * FROM producto WHERE referencia = '" & cCod_Articulo.Text & "'"
        crea_record rs, consulta, False
        If rs.RecordCount = 1 Then
            cDescripcion.Text = rs("nombre")
            cPrecioUnidad.Text = rs("precio")
            Me.cCantidad.Text = 1
        Else
            MsgBox "No se ha encontrado ningún producto con esa referencia."
            Me.cCod_Articulo.SetFocus
            DoEvents
        End If
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Me.DataGrid1.Row > -1 Then
        Me.DataGrid1.Col = 2
        If Me.DataGrid1.Text = Me.cCantidad.Text Then
            DataGrid1.Col = 0
            cCod_Articulo.Text = DataGrid1.Text
            DataGrid1.Col = 1
            cDescripcion.Text = DataGrid1.Text
            DataGrid1.Col = 2
            cCantidad.Text = DataGrid1.Text
            DataGrid1.Col = 3
            cPrecioUnidad.Text = DataGrid1.Text
        End If
    End If
End Sub
Private Sub cCP_LostFocus()
    On Error GoTo noesnum
        If Not compruebaNum(CLng(Me.cCP.Text)) Then
            Me.cCP.SetFocus
        End If
        Exit Sub
noesnum:
    MsgBox "Debe introducir un número entero positivo."
    'Me.cCP.SetFocus
End Sub

Private Sub cNumero_LostFocus()
    On Error GoTo noesnum
        If Not compruebaNum(CLng(Me.cNumero.Text)) Then
            Me.cNumero.SetFocus
        End If
        Exit Sub
noesnum:
    MsgBox "Debe introducir un número entero positivo."
    'Me.cNumero.SetFocus
End Sub

Private Sub cPais_LostFocus()
    Dim rs As ADODB.Recordset
    MiCommand.CommandText = "SELECT NOMBRE from pais where nombre='" & Me.cPais.Text & "'"
    Set rs = MiCommand.Execute
    If rs.RecordCount <> 1 Then MsgBox "No existe ningún país con el nombre especificado"
            
End Sub

Private Sub cPuerta_LostFocus()
    On Error GoTo noesnum
        If Not compruebaNum(CLng(Me.cPuerta.Text)) Then
            Me.cPuerta.SetFocus
        End If
        Exit Sub
noesnum:
    MsgBox "Debe introducir un número entero positivo."
    'Me.cPuerta.SetFocus
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'las siguientes lineas se añaden porque si no el ultimo pedido creado no se muestra en el grid
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs1 = Elaborar_pedido_CU02.DataGrid1.DataSource
    rs1.Requery
    MiCommand.CommandText = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & codcli & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
    Set rs2 = MiCommand.Execute
    Set Elaborar_pedido_CU02.DataGrid2.DataSource = rs2
    Elaborar_pedido_CU02.Visible = True
End Sub
