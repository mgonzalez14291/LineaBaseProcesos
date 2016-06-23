VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Atender_pedido 
   Caption         =   "Atender Pedido"
   ClientHeight    =   7545
   ClientLeft      =   4245
   ClientTop       =   1845
   ClientWidth     =   9720
   Icon            =   "Atender Pedido CU08.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cFechaAtencion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Incidencia 
      Caption         =   "Incidencia"
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton btPasarEnvio 
      Caption         =   "&Pasar a Envíos"
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos del Técnico de Almacen"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9495
      Begin VB.TextBox cNif 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox cNomTecnico 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton bGuardar 
      Caption         =   "&Guardar "
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox cCod_pedido 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   9495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "referencia"
            Caption         =   "Cód. Artículo"
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
            Caption         =   "Cant. Solicitada"
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
            DataField       =   "cant_asignada"
            Caption         =   "Cant. Asignada"
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
            DataField       =   "stock"
            Caption         =   "Stock Real"
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
         BeginProperty Column05 
            DataField       =   "stock_disponible"
            Caption         =   "Stock  Disponible"
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
         BeginProperty Column06 
            DataField       =   "stock_asignado"
            Caption         =   "Stock Asignado"
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
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1844,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1379,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1275,024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Pedido"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   9495
      Begin VB.TextBox cFechaLlegadaAlm 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox cCalle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox cNum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox cPta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox cLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox cProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox cPais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox cCP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de atención "
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha llegada almacén"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Nº"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Pta"
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   7080
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "País"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "CP"
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección Envío"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Código del Pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   7440
      Picture         =   "Atender Pedido CU08.frx":27A2
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5520
      Picture         =   "Atender Pedido CU08.frx":2BE4
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3600
      Picture         =   "Atender Pedido CU08.frx":3026
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1680
      Picture         =   "Atender Pedido CU08.frx":3468
      Top             =   6480
      Width           =   480
   End
End
Attribute VB_Name = "Atender_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsped As ADODB.Recordset
Public rslinped As ADODB.Recordset

Private Sub btPasarEnvio_Click()
    Dim rs_envio As ADODB.Recordset
    Dim cad_envio As String
    Dim consultagrid As String
    Dim i As Integer
    Dim completo As Boolean
    
    Set rs_envio = New ADODB.Recordset
    cad_envio = "SELECT * FROM Linea_pedido WHERE codigo_pedido =" & Me.cCod_pedido & ""
    
    crea_record rs_envio, cad_envio, False
    If Not rs_envio.EOF Then
        i = 1
        completo = True
        rs_envio.MoveFirst
        While i <= rs_envio.RecordCount
            'rs_envio.Bookmark = i
            If rs_envio("cantidad") <> rs_envio("cant_asignada") Then
                completo = False
            End If
            i = i + 1
            rs_envio.MoveNext
        Wend
        
        If Not completo Then
            If MsgBox("El pedido está incompleto. ¿Desea continuar?", vbOKCancel) = 1 Then
                If MsgBox("¿Desea crear un nuevo pedido con las cantidades restantes?", vbOKCancel) = 1 Then
                    
                    Dim num_ped As Integer
                    Dim rs_contador As ADODB.Recordset
                    Set rs_contador = New ADODB.Recordset
                    
                    crea_record rs_contador, "select * from orden_pedido", False
                    
                    num_ped = rs_contador.RecordCount + 1
                    
                    Dim rs_insert As ADODB.Recordset
                    Dim cad_insert As String
                    Set rs_insert = New ADODB.Recordset
                    cad_insert = "select * from orden_pedido where codigo_pedido = " & Me.cCod_pedido & ""
                    crea_record rs_insert, cad_insert, False
                    
                    'hemos de hacer el commit para que al insertar las filas en Linea_pedido no dé error de clave ajena
                    MiConexion.BeginTrans
                    If oracle Then
                        MiCommand.CommandText = "INSERT INTO Orden_pedido (codigo_pedido, cliente, usuario_ventas, CP_envio, Pais_envio, Provincia_envio, Localidad_envio, Pta_envio, numero_envio, calle_envio, forma_pago, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion) VALUES (" & num_ped & ", '" & rs_insert("cliente") & "', '" & rs_insert("usuario_ventas") & "', " & rs_insert("CP_envio") & ", '" & rs_insert("Pais_envio") & "', '" & rs_insert("Provincia_envio") & "', '" & rs_insert("Localidad_envio") & "', " & rs_insert("Pta_envio") & ", " & rs_insert("numero_envio") & ", '" & rs_insert("calle_envio") & "', '" & rs_insert("forma_pago") & "',TO_Date( '" & rs_insert("fecha_elaboracion") & "', 'DD/MM/YYYY HH:MI:SS AM'), TO_Date( '" & rs_insert("fecha_llegada_almacen") & "', 'DD/MM/YYYY HH:MI:SS AM'),TO_Date( '" & rs_insert("fecha_atencion") & "', 'DD/MM/YYYY HH:MI:SS AM'))"
                    Else
                        MiCommand.CommandText = "INSERT INTO Orden_pedido (codigo_pedido, cliente, usuario_ventas, CP_envio, Pais_envio, Provincia_envio, Localidad_envio, Pta_envio, numero_envio, calle_envio, forma_pago, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion) VALUES (" & num_ped & ", '" & rs_insert("cliente") & "', '" & rs_insert("usuario_ventas") & "', " & rs_insert("CP_envio") & ", '" & rs_insert("Pais_envio") & "', '" & rs_insert("Provincia_envio") & "', '" & rs_insert("Localidad_envio") & "', " & rs_insert("Pta_envio") & ", " & rs_insert("numero_envio") & ", '" & rs_insert("calle_envio") & "', '" & rs_insert("forma_pago") & "', #" & rs_insert("fecha_elaboracion") & "#, #" & rs_insert("fecha_llegada_almacen") & "#, #" & rs_insert("fecha_atencion") & "#)"
                    End If
                    MiCommand.Execute
                    MiConexion.CommitTrans
                    
                          
                    Dim nueva_cant As Integer
                    i = 1
                    rs_envio.MoveFirst
                    While i <= rs_envio.RecordCount
                        If rs_envio("cantidad") <> rs_envio("cant_asignada") Then
                        
                            'Creamos el nuevo pedido
                            nueva_cant = rs_envio("cantidad") - rs_envio("cant_asignada")
                            MiCommand.CommandText = "INSERT INTO Linea_pedido (codigo_pedido, referencia, cantidad, precio, cant_asignada) VALUES (" & num_ped & ", '" & rs_envio("referencia") & "', " & nueva_cant & ", " & ((rs_envio("precio") / rs_envio("cantidad")) * nueva_cant) & ", 0)"
                            MiCommand.Execute
                            
                            'Actualizamos el existente
                            If rs_envio("cant_asignada") = 0 Then
                                MiCommand.CommandText = "DELETE * FROM Linea_pedido  WHERE codigo_pedido = " & rs_envio("codigo_pedido") & " AND referencia ='" & rs_envio("referencia") & "'"
                                MiCommand.Execute
                            Else
                                MiCommand.CommandText = "UPDATE Linea_pedido SET cantidad = " & rs_envio("cant_asignada") & " WHERE codigo_pedido = " & rs_envio("codigo_pedido") & " AND referencia ='" & rs_envio("referencia") & "'"
                                MiCommand.Execute
                            End If
                                
                        End If
                        i = i + 1
                        rs_envio.MoveNext
                    Wend
                    
                    'pasamos a envío y actualizamos el grid
                    pasar_envio
                    
                Else 'pasamos a envío y actualizamos el grid
                    pasar_envio
                End If
            End If
        Else 'pasamos a envío y actualizamos el grid
            pasar_envio
        End If
    End If
        
End Sub
Sub pasar_envio() 'función que pasa a envio, actualiza el grid, y cierra la ventana actual

    If oracle Then
        MiCommand.CommandText = "UPDATE orden_pedido SET fecha_listo_envio = TO_Date( '" & Date & "', 'DD/MM/YYYY HH:MI:SS AM') WHERE codigo_pedido = " & Me.cCod_pedido & ""
    Else
        MiCommand.CommandText = "UPDATE orden_pedido SET fecha_listo_envio = #" & Date & "# WHERE codigo_pedido = " & Me.cCod_pedido & ""
    End If
    MiCommand.Execute
                
    consultagrid = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion FROM Orden_pedido WHERE fecha_llegada_almacen IS NOT NULL AND fecha_atencion is NOT NULL and fecha_listo_envio is NULL and pais_envio in (select nombre from pais where codigo_region='" & Tecnico_Almacen.region & "')"
                
    Tecnico_Almacen.rs_ordenes.Close
    Tecnico_Almacen.rs_ordenes.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
    
    Set Tecnico_Almacen.DataGrid2.DataSource = Tecnico_Almacen.rs_ordenes
    
    Unload Me
    Tecnico_Almacen.Visible = True
End Sub
Private Sub bGuardar_Click()
    'rsped.Update
    Dim noatendido As Boolean
    noatendido = True
    
    Me.rslinped.MoveFirst
    While Not Me.rslinped.EOF
        If CInt(Me.rslinped("cant_asignada")) <> 0 Then
            noatendido = False
        End If
        Me.rslinped.MoveNext
    Wend
    If Not noatendido Or Not IsNull(rsped!fecha_LISTO_ENVIO) Then
        Guardar_Pedido_Atencion (cFechaAtencion.Text)
        MiConexion.CommitTrans
    Else
        MiConexion.RollbackTrans
    End If
    Unload Me
    Tecnico_Almacen.Visible = True
End Sub

Private Sub bSalir_Click()
    MiConexion.RollbackTrans
    Unload Me
    Tecnico_Almacen.Visible = True
End Sub
'al hacer doble clic sobre el datagrid se abrira una ventana
'en la q el usuario podra asignar stock disponible al producto
'de la linea sobre la q ha hecho doble clic
Private Sub DataGrid1_DblClick()
    Asignar_Stock.stock_asignado = rslinped!stock_asignado
    Asignar_Stock.stock_real = rslinped!stock
    Asignar_Stock.stock_disponible = rslinped!stock_disponible
    
    Asignar_Stock.cCod_Articulo = rslinped!referencia
    Asignar_Stock.cNombre = rslinped!nombre
    Asignar_Stock.cant_solicitada = rslinped!cantidad
    Asignar_Stock.cCantidad = rslinped!cant_asignada
    
    Asignar_Stock.Show vbModal
    
    If Asignar_Stock.Aceptar_cambio = True Then
        bGuardar.Enabled = True
    End If
End Sub

Sub Cargar_pedido(cod_ped As String)
    
    Dim consulta As String
    Dim consulta2 As String
    
    Set rsped = New ADODB.Recordset
    Set rslinped = New ADODB.Recordset
       
    consulta = "SELECT * FROM orden_pedido WHERE codigo_pedido=" & cod_ped & ""
    consulta2 = "SELECT linea_pedido.codigo_pedido,linea_pedido.cant_asignada, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad, producto_almacen.stock_asignado, producto_almacen.stock,producto_almacen.stock - producto_almacen.stock_asignado as stock_disponible From linea_pedido, producto, producto_almacen where codigo_pedido=" & cod_ped & " and producto.referencia=linea_pedido.referencia and linea_pedido.referencia = producto_almacen.referencia and producto_almacen.almacen='" & Tecnico_Almacen.cod_almacen & "'"
    
    rsped.Open consulta, MiConexion, adOpenDynamic, adLockOptimistic
    rslinped.Open consulta2, MiConexion, adOpenDynamic, adLockOptimistic
    
    MiConexion.BeginTrans
    
    Set DataGrid1.DataSource = rslinped
    
    ' Para cuando tengo el datagrid no esta seleccionado
    DataGrid1.Tag = False
    
    cCod_pedido.Text = rsped!codigo_pedido
    cFechaLlegadaAlm.Text = rsped!fecha_llegada_almacen
        
    cNif.Text = Identificacion.adRst!nif
    cNomTecnico.Text = Identificacion.adRst!nombre
        
    ' si ya esta en la bd se recupera, sino el actual
        
    If IsNull(rsped!fecha_atencion) Then
        Atender_pedido.Caption = Atender_pedido.Caption + ". [Primera atención]"
        cFechaAtencion.Text = Date
        Me.btPasarEnvio.Enabled = False
    Else
        Atender_pedido.Caption = Atender_pedido.Caption + " [Modificacion]"
        cFechaAtencion.Text = rsped!fecha_atencion
        Me.btPasarEnvio.Enabled = True
    End If
                
    cCalle.Text = rsped!calle_envio
    cNum.Text = rsped!numero_envio
    cPta.Text = rsped!pta_envio
    cLocalidad.Text = rsped!localidad_envio
    cProvincia.Text = rsped!provincia_envio
    cPais.Text = rsped!pais_envio
    cCP.Text = rsped!cp_envio

End Sub

Private Sub Guardar_Pedido_Atencion(Fecha As Date)

    rsped!fecha_atencion = Fecha
    rsped.Update

    Tecnico_Almacen.rs_ordenes.Requery
End Sub

Private Sub Incidencia_Click()
    Me.Visible = False
    Incidencia_pedido.cargar_datos_incidencia rsped!codigo_pedido
    Incidencia_pedido.Show vbModal
End Sub

