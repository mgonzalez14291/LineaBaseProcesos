VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Elaborar_pedido_CU02 
   Caption         =   "Elaborar Pedido"
   ClientHeight    =   8880
   ClientLeft      =   720
   ClientTop       =   345
   ClientWidth     =   10665
   Icon            =   "Elaborar Pedido CU02.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8880
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8640
      TabIndex        =   44
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos del Representante de Ventas"
      Height          =   1215
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton bGestionClientes 
         Caption         =   "&Gestión de Clientes"
         Height          =   375
         Left            =   8520
         TabIndex        =   43
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox cNombreRepres 
         Height          =   285
         Left            =   2520
         TabIndex        =   42
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox cCodigoRepres 
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   9120
         Picture         =   "Elaborar Pedido CU02.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label12 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2520
         TabIndex        =   40
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos En Elaboración"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   10455
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1095
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1931
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "codigo_pedido"
            Caption         =   "Codigo"
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
            DataField       =   "fecha_elaboracion"
            Caption         =   "Fecha Elaboracion"
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
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2085,166
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "&Cancelar Pedido"
         Height          =   375
         Left            =   8640
         TabIndex        =   30
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton bModificar 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   6720
         TabIndex        =   29
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton bNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   4800
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   9240
         Picture         =   "Elaborar Pedido CU02.frx":074C
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   7320
         Picture         =   "Elaborar Pedido CU02.frx":0B8E
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   5400
         Picture         =   "Elaborar Pedido CU02.frx":0FD0
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pedidos Enviados a Almacen"
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   10455
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1095
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   1931
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "codigo_pedido"
            Caption         =   "Código"
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
            DataField       =   "fecha_elaboracion"
            Caption         =   "Fecha elaboración"
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
            DataField       =   "fecha_llegada_almacen"
            Caption         =   "Fecha llegada almacén"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1890,142
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton bConsultar 
         Caption         =   "Consultar &Pedido"
         Height          =   375
         Left            =   8520
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   9120
         Picture         =   "Elaborar Pedido CU02.frx":1412
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos cliente"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10455
      Begin VB.TextBox cEsoperadora 
         Height          =   375
         Left            =   5400
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox cPta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   37
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox cNum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   36
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox cCalle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox cPais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   33
         Top             =   2385
         Width           =   1455
      End
      Begin VB.TextBox cPerContacto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   26
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox cTlfContacto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   24
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox cEmail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   22
         Top             =   2385
         Width           =   2175
      End
      Begin VB.TextBox cFax 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Top             =   2385
         Width           =   2175
      End
      Begin VB.TextBox cCp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9360
         TabIndex        =   18
         Top             =   1785
         Width           =   855
      End
      Begin VB.TextBox cProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   16
         Top             =   1785
         Width           =   2175
      End
      Begin VB.TextBox cLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox cNombre 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CommandButton bBuscar 
         Caption         =   "&Buscar"
         Height          =   405
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox cTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox cDni 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox cCodigoCli 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5520
         Picture         =   "Elaborar Pedido CU02.frx":1854
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label22 
         Caption         =   "País"
         Height          =   255
         Left            =   7080
         TabIndex        =   34
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Persona de Contacto"
         Height          =   255
         Left            =   7080
         TabIndex        =   27
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label17 
         Caption         =   "Telefono contacto"
         Height          =   255
         Left            =   7080
         TabIndex        =   25
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Fax"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "CP"
         Height          =   255
         Left            =   9360
         TabIndex        =   19
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Pta"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Calle"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Telefono"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Dni/Cif"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   9240
      Picture         =   "Elaborar Pedido CU02.frx":1C96
      Top             =   7800
      Width           =   480
   End
End
Attribute VB_Name = "Elaborar_pedido_CU02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim esoperadora As Boolean
Dim nombre As String

Private Sub bBuscar_Click()
    Dim codigocli As String
    Dim nifrepres As String
    Dim dni As String
    Dim consulta As String
    Dim consulta2 As String
    Dim consultagrid As String
    Dim consultagrid2 As String
    Dim nombre As String
   
    Dim adRst1 As Recordset
    Dim RstGrid As Recordset
    Dim RstGrid2 As Recordset
    
    nifrepres = cCodigoRepres.Text
    
    '************************ CLASE CLIENTE: buscar_cliente ************************
    'Código equivalente a la operación de búsqueda definida en la clase cliente del
    'modelo de análisis/diseño (diagrama de clases)
    If cNombre.Text = "" Then
        MsgBox "Introduzca un nombre para la búsqueda"
    Else
    nom = cNombre.Text
    If Not esoperadora Then
        If oracle Then
            consulta = "SELECT * FROM cliente WHERE nombre LIKE '" & nom & "%' AND representante = '" & nifrepres & "'"
        Else
            consulta = "SELECT * FROM cliente WHERE nombre LIKE '" & nom & "%' AND representante = '" & nifrepres & "'"
        End If
    Else 'si es operadora puede ver todos los clientes
        If oracle Then
            consulta = "SELECT * FROM cliente WHERE nombre LIKE '" & nom & "%'"
        Else
            consulta = "SELECT * FROM cliente WHERE nombre LIKE '" & nom & "%'"
        End If
    End If
    ' ****************************** fin_buscar_cliente ******************************
    
    crea_record adRst1, consulta, False
    If adRst1.RecordCount = 0 Then
        MsgBox "Error , no existe ningun cliente con ese nombre, o no está representado por ud."
    Else
        If adRst1.RecordCount = 1 Then 'Asignar datos del cliente a los campos de la interfaz
            cNombre.Text = adRst1("nombre")
            cDni.Text = adRst1("nif_cif")
            cCodigoCli.Text = adRst1("codigo")
            cPerContacto.Text = IIf(adRst1("persona_contacto") <> "", adRst1("persona_contacto"), "")
            cTlfContacto.Text = IIf(adRst1("tlf_pers_contacto") <> "", adRst1("tlf_pers_contacto"), "")
            cCalle.Text = IIf(adRst1("calle") <> "", adRst1("calle"), "")
            cNum.Text = IIf(adRst1("numero") <> "", adRst1("numero"), "")
            cPta.Text = IIf(adRst1("puerta") <> "", adRst1("puerta"), "")
            cLocalidad.Text = IIf(adRst1("localidad") <> "", adRst1("localidad"), "")
            cProvincia.Text = IIf(adRst1("provincia") <> "", adRst1("provincia"), "")
            cCp.Text = IIf(adRst1("CP") <> "", adRst1("CP"), "")
            cTelefono.Text = IIf(adRst1("telefono") <> "", adRst1("telefono"), "")
            cFax.Text = IIf(adRst1("fax") <> "", adRst1("fax"), "")
            cEmail.Text = IIf(adRst1("email") <> "", adRst1("email"), "")
            cPais.Text = IIf(adRst1("pais") <> "", adRst1("pais"), "")
                
            'Para llenar el grid
            consultagrid = "SELECT codigo_pedido, fecha_elaboracion FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is null"
            consultagrid2 = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
                
            crea_record RstGrid2, consultagrid2, False
            Set RstGrid = New ADODB.Recordset
            RstGrid.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
            Set DataGrid1.DataSource = RstGrid
            Set DataGrid2.DataSource = RstGrid2
       Else
            Elegir_nombre.Show
       End If
    End If
End If
    
    
End Sub

'*************** CLASE ORDEN DE PEDIDO: baja_orden_de_pedido **********************
Private Sub bCancelar_Click()
  If Me.cCodigoCli.Text <> "" And Me.cDni.Text <> "" Then
    Dim rs As ADODB.Recordset
    Set rs = Me.DataGrid1.DataSource
    If rs.RecordCount >= 1 And Me.DataGrid1.Row >= 0 Then
      If MsgBox("¿Está seguro que desea eliminar el pedido?", vbOKCancel) = 1 Then
        Me.DataGrid1.AllowDelete = True
        rs.Delete (adAffectCurrent)
        Me.DataGrid1.Refresh
      End If
    End If
  End If
  '************************* fin baja_orden_de_pedido *****************************
End Sub


'*************** CLASE ORDEN DE PEDIDO: consultar_orden_de_pedido *****************
Private Sub bConsultar_Click()
  If Me.cCodigoCli.Text <> "" And Me.cDni.Text <> "" Then
    Dim rs As ADODB.Recordset
    Set rs = Me.DataGrid2.DataSource
    If rs.RecordCount >= 1 And Me.DataGrid2.Row >= 0 Then
        Me.Visible = False
        Me.DataGrid2.Col = 0
        Elaborar_pedido_consultar_enviados_almacen.Cargar_Pedido (Me.DataGrid2.Text)
        Elaborar_pedido_consultar_enviados_almacen.Show
    End If
  End If
  '*********************** fin consultar_orden_de_pedido **************************
End Sub


'*********************** CASO DE USO: Elaborar Pedido (modificar) *****************
Private Sub bModificar_Click()
  If Me.cCodigoCli.Text <> "" And Me.cDni.Text <> "" Then
    Dim rs As ADODB.Recordset
    Set rs = Me.DataGrid1.DataSource
    If rs.RecordCount >= 1 And Me.DataGrid1.Row >= 0 Then
        Me.Visible = False
        Me.DataGrid1.Col = 0
        Elaborar_Pedido_Nuevo_Modif.Cargar_Pedido (Me.DataGrid1.Text)
        Elaborar_Pedido_Nuevo_Modif.Caption = "Modificar Pedido"
        Elaborar_Pedido_Nuevo_Modif.Show
    End If
  End If
  '********************************************************************************
End Sub


'*********************** CASO DE USO: Elaborar Pedido (nuevo) *****************
Private Sub bNuevo_Click()
    If Me.cCodigoCli.Text <> "" And Me.cDni.Text <> "" Then
        Me.Visible = False
        Elaborar_Pedido_Nuevo_Modif.cCodRepresOperadora.Text = Me.cCodigoRepres.Text
        Elaborar_Pedido_Nuevo_Modif.PonDireccionEnvio Me.cCodigoCli.Text 'Me.cDni.Text
        Elaborar_Pedido_Nuevo_Modif.Caption = "Elaborar Pedido Nuevo"
        Elaborar_Pedido_Nuevo_Modif.Show
    End If
'********************************************************************************
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub


'********************** CLASE CLIENTE: consultar_datos_cliente *******************
Private Sub cCodigoCli_LostFocus()
    If Not Me.cCodigoCli.Text = "" Then
        Me.cDni.Text = ""
    End If
    
    Dim codigocli As String
    Dim nifrepres As String
    Dim dni As String
    Dim consulta As String
    Dim consulta2 As String
    Dim consultagrid As String
    Dim consultagrid2 As String
    Dim consultagrid3 As String
    Dim consultagrid4 As String
   
    Dim adRst As Recordset
    Dim RstGrid As Recordset
    Dim RstGrid2 As Recordset
    
    nifrepres = cCodigoRepres.Text
    codigocli = cCodigoCli.Text
    dni = cDni.Text
    If Not esoperadora Then
        consulta2 = "SELECT * FROM cliente WHERE codigo ='" & codigocli & "' AND representante ='" & nifrepres & "'"
    Else
        'si es operadora puede ver todos los clientes
        consulta2 = "SELECT * FROM cliente WHERE codigo ='" & codigocli & "'"
    End If
    consultagrid = "SELECT codigo_pedido, fecha_elaboracion FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is null"
    consultagrid2 = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
    
    If codigocli <> "" Then
            crea_record adRst, consulta2, False
            If adRst.RecordCount = 0 Then
                MsgBox "Error , este cliente no existe o no es representado por usted."
            Else
                cNombre.Text = adRst("nombre")
                cDni.Text = adRst("nif_cif")
                cPerContacto.Text = IIf(adRst("persona_contacto") <> "", adRst("persona_contacto"), "")
                cTlfContacto.Text = IIf(adRst("tlf_pers_contacto") <> "", adRst("tlf_pers_contacto"), "")
                cCalle.Text = IIf(adRst("calle") <> "", adRst("calle"), "")
                cNum.Text = IIf(adRst("numero") <> "", adRst("numero"), "")
                cPta.Text = IIf(adRst("puerta") <> "", adRst("puerta"), "")
                cLocalidad.Text = IIf(adRst("localidad") <> "", adRst("localidad"), "")
                cProvincia.Text = IIf(adRst("provincia") <> "", adRst("provincia"), "")
                cCp.Text = IIf(adRst("CP") <> "", adRst("CP"), "")
                cTelefono.Text = IIf(adRst("telefono") <> "", adRst("telefono"), "")
                cFax.Text = IIf(adRst("fax") <> "", adRst("fax"), "")
                cEmail.Text = IIf(adRst("email") <> "", adRst("email"), "")
                cPais.Text = IIf(adRst("pais") <> "", adRst("pais"), "")
                
                'Para llenar el grid
                crea_record RstGrid2, consultagrid2, False
                Set RstGrid = New ADODB.Recordset
                RstGrid.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
                Set DataGrid1.DataSource = RstGrid
                Set DataGrid2.DataSource = RstGrid2
            End If
    End If
End Sub


Private Sub cDni_LostFocus()
    If Not Me.cDni.Text = "" Then
        Me.cCodigoCli.Text = ""
    End If
    
    Dim codigocli As String
    Dim nifrepres As String
    Dim dni As String
    Dim consulta As String
    Dim consulta2 As String
    Dim consultagrid As String
    Dim consultagrid2 As String
    Dim consultagrid3 As String
    Dim consultagrid4 As String
   
    Dim adRst As Recordset
    Dim RstGrid As Recordset
    Dim RstGrid2 As Recordset
    
    nifrepres = cCodigoRepres.Text
    codigocli = cCodigoCli.Text
    dni = cDni.Text
    If Not esoperadora Then
        consulta = "SELECT * FROM cliente WHERE nif_cif ='" & dni & "' AND representante = '" & nifrepres & "'"
    Else 'si es operadora puede ver todos los clientes
        consulta = "SELECT * FROM cliente WHERE nif_cif ='" & dni & "'" ' AND representante = '" & nifRepres & "'"
    End If
    consultagrid = "SELECT codigo_pedido, fecha_elaboracion FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is null"
    consultagrid2 = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & cCodigoCli.Text & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
    
    If cDni <> "" Then
            crea_record adRst, consulta, False
            If adRst.RecordCount = 0 Then
                MsgBox "Error , este cliente no existe o no es representado por usted."
            Else
                cNombre.Text = adRst("nombre")
                cCodigoCli.Text = adRst("codigo")
                cPerContacto.Text = IIf(adRst("persona_contacto") <> "", adRst("persona_contacto"), "")
                cTlfContacto.Text = IIf(adRst("tlf_pers_contacto") <> "", adRst("tlf_pers_contacto"), "")
                cCalle.Text = IIf(adRst("calle") <> "", adRst("calle"), "")
                cNum.Text = IIf(adRst("numero") <> "", adRst("numero"), "")
                cPta.Text = IIf(adRst("puerta") <> "", adRst("puerta"), "")
                cLocalidad.Text = IIf(adRst("localidad") <> "", adRst("localidad"), "")
                cProvincia.Text = IIf(adRst("provincia") <> "", adRst("provincia"), "")
                cCp.Text = IIf(adRst("CP") <> "", adRst("CP"), "")
                cTelefono.Text = IIf(adRst("telefono") <> "", adRst("telefono"), "")
                cFax.Text = IIf(adRst("fax") <> "", adRst("fax"), "")
                cEmail.Text = IIf(adRst("email") <> "", adRst("email"), "")
                cPais.Text = IIf(adRst("pais") <> "", adRst("pais"), "")
                
                'Para llenar el grid
                crea_record RstGrid2, consultagrid2, False
                Set RstGrid = New ADODB.Recordset
                RstGrid.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
                Set DataGrid1.DataSource = RstGrid
                Set DataGrid2.DataSource = RstGrid2
            End If
        End If
'*************************** fin consultar_datos_cliente *************************
End Sub



Private Sub Form_Load()
    Dim id As String
    Dim adRst As Recordset
    id = Identificacion.cUsuario.Text
    'Obtenemos el nif del representante automáticamente a partir de su id
    consulta = "SELECT nif, cargo, nombre FROM empleado WHERE login='" & id & "'"
    crea_record adRst, consulta, False
    cCodigoRepres.Text = adRst("nif")
    cCodigoRepres.Enabled = False
    cNombreRepres.Text = adRst("nombre")
    cNombreRepres.Enabled = False
    If adRst("cargo") = "Operadora" Then
        esoperadora = True
        cEsoperadora = 1
    Else
        esoperadora = False
        cEsoperadora = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Identificacion.Visible = True
End Sub

