VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Elaborar_pedido_consultar_enviados_almacen 
   Caption         =   "Consultar Pedidos Enviados al Almacén"
   ClientHeight    =   7665
   ClientLeft      =   3165
   ClientTop       =   1725
   ClientWidth     =   8625
   Icon            =   "Elaborar Pedido (Consultar enviados almacen).frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      Height          =   4335
      Left            =   0
      TabIndex        =   22
      Top             =   2160
      Width           =   8535
      Begin VB.ComboBox cmbFormaPago 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Elaborar Pedido (Consultar enviados almacen).frx":0442
         Left            =   1440
         List            =   "Elaborar Pedido (Consultar enviados almacen).frx":044C
         TabIndex        =   29
         Top             =   3555
         Width           =   1455
      End
      Begin VB.TextBox cIVA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   24
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox cTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   23
         Top             =   3840
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3015
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.Label Label15 
         Caption         =   "IVA"
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Total"
         Height          =   255
         Left            =   6480
         TabIndex        =   26
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   21
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Pedido"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox cCP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox cPais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   9
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox cProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox cLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox cPuerta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox cNumero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox cDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox cCodRepresOperadora 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox cCodPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cFechaElabPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "CP"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "País"
         Height          =   255
         Left            =   6600
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Pta"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Nº"
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección Envío"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo Representante/Operadora"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6600
      Picture         =   "Elaborar Pedido (Consultar enviados almacen).frx":0467
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "Elaborar_pedido_consultar_enviados_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bSalir_Click()
    Unload Me
    Elaborar_pedido_CU02.Visible = True
End Sub
Sub Cargar_Pedido(cod_ped As String)
    Dim rsped As ADODB.Recordset
    Dim consulta As String
    Dim consulta2 As String
     
'**************** CLASE ORDEN DE PEDIDO: consultar_orden_de_pedido *****************
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
    Me.cmbFormaPago.ListIndex = rsped("forma_pago")
    
    calcula_ivatotal cod_ped
'************************** fin consultar_orden_de_pedido **************************
End Sub

Sub calcula_ivatotal(cod_ped As String)
    Dim rs As ADODB.Recordset
    Dim cons As String
    cons = "SELECT sum(linea_pedido.precio * cantidad) as total From linea_pedido where codigo_pedido=" & cod_ped & ""
    crea_record rs, cons, False
    Me.cIVA.Text = rs("total") * 0.16
    Me.cTotal.Text = rs("total") + (rs("total") * 0.16)
End Sub

