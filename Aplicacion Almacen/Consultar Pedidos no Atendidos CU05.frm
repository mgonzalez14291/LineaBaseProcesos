VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Consultar_pedidos_no_atendidos_CU05 
   Caption         =   "Consultar Pedidos no Atendidos"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "Consultar Pedidos no Atendidos CU05.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Atender Pedido"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Pedido"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      Begin VB.TextBox cFechaLlegaAlm 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox cFechaElabPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox cCodPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox cNumero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox cPuerta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox cLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox cProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox cPais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox cCP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de llegada al almacén"
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Elaboracion"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Dirección Envío"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Nº"
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Pta"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "País"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "CP"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.CommandButton bSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   8535
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3135
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5530
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
            DataField       =   "Referencia"
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
            DataField       =   "Cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4935,118
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6120
      Picture         =   "Consultar Pedidos no Atendidos CU05.frx":0442
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2400
      Picture         =   "Consultar Pedidos no Atendidos CU05.frx":0884
      Top             =   6000
      Width           =   480
   End
End
Attribute VB_Name = "Consultar_pedidos_no_atendidos_CU05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsped As ADODB.Recordset
Dim rslinped As ADODB.Recordset

Private Sub bSalir_Click()
    Unload Me
    Tecnico_Almacen.Visible = True
End Sub
'pasa el pedido a "en atencion" y abre la ventana de atender pedido
'para comenzar a atenderlo
Private Sub Command1_Click()
    If rslinped.RecordCount >= 1 Then
        'rsped("fecha_atencion") = Date
        'rsped.Update
        Me.Visible = False
        Atender_pedido.Cargar_pedido (cCodPedido.Text)
        Atender_pedido.Show vbModal
    End If
'    Tecnico_Almacen.Fecha_Atencion_Pedido (Text1.Text)
'    Unload Me
End Sub
 
Sub Cargar_pedido(cod_ped As String)

    Dim consulta As String
    Dim consulta2 As String
     
    consulta = "SELECT * FROM orden_pedido WHERE codigo_pedido=" & cod_ped & ""
    consulta2 = "SELECT linea_pedido.codigo_pedido, linea_pedido.referencia,producto.nombre,linea_pedido.cantidad,linea_pedido.precio,linea_pedido.cantidad * linea_pedido.precio as total From linea_pedido, producto where codigo_pedido=" & cod_ped & " and producto.referencia=linea_pedido.referencia"
    Set rslinped = New ADODB.Recordset
    rslinped.Open consulta2, MiConexion, adOpenDynamic, adLockOptimistic
    Set rsped = New ADODB.Recordset
    rsped.Open consulta, MiConexion, adOpenDynamic, adLockOptimistic

    Set Me.DataGrid1.DataSource = rslinped
    
    Me.cCodPedido.Text = cod_ped

    Me.cFechaElabPedido.Text = rsped("fecha_elaboracion")
    Me.cFechaLlegaAlm.Text = rsped("fecha_llegada_almacen")
    Me.cDireccion.Text = rsped("calle_envio")
    Me.cNumero.Text = rsped("numero_envio")
    Me.cPuerta.Text = rsped("pta_envio")
    Me.cLocalidad.Text = rsped("localidad_envio")
    Me.cProvincia.Text = rsped("provincia_envio")
    Me.cPais.Text = rsped("pais_envio")
    Me.cCP.Text = rsped("cp_envio")
End Sub
