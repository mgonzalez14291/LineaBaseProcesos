VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Tecnico_almacen 
   Caption         =   "Técnico de Almacén"
   ClientHeight    =   6795
   ClientLeft      =   3855
   ClientTop       =   2430
   ClientWidth     =   9195
   Icon            =   "Tecnico de Almacen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "No Atendidos"
      TabPicture(0)   =   "Tecnico de Almacen.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DataGrid1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btConsultarNoAtendidos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btCancelarPedido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "btSalir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btAtender"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "En Atención"
      TabPicture(1)   =   "Tecnico de Almacen.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Image3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Image4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DataGrid2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "bAtenderPed"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btCancelPedEnAtencion"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "bPasoListoEnvio"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "btSalir2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Listos para Envío"
      TabPicture(2)   =   "Tecnico de Almacen.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Image10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DataGrid3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btConsultarModificar"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "btCancelPedListoEnvio"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "btSalir3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Enviados"
      TabPicture(3)   =   "Tecnico de Almacen.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Image11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Image12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "DataGrid4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "btConsultar"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "btSalir4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.CommandButton btAtender 
         Caption         =   "&Atender Pedido"
         Height          =   375
         Left            =   -72600
         TabIndex        =   17
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btSalir4 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton btSalir3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   -68760
         TabIndex        =   15
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btSalir2 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   -68280
         TabIndex        =   14
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   -68280
         TabIndex        =   13
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton bPasoListoEnvio 
         Caption         =   "&Pasar a listo para envío"
         Height          =   375
         Left            =   -70440
         TabIndex        =   12
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btConsultar 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton btCancelPedListoEnvio 
         Caption         =   "Cancelar &Pedido"
         Height          =   375
         Left            =   -71400
         TabIndex        =   9
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btConsultarModificar 
         Caption         =   "&Consultar / Modificar"
         Height          =   375
         Left            =   -74040
         TabIndex        =   8
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btCancelPedEnAtencion 
         Caption         =   "&Cancelar Pedido"
         Height          =   375
         Left            =   -72600
         TabIndex        =   6
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton bAtenderPed 
         Caption         =   "&Atender Pedido"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btCancelarPedido 
         Caption         =   "Cancelar &Pedido"
         Height          =   375
         Left            =   -70440
         TabIndex        =   3
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton btConsultarNoAtendidos 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   -74760
         TabIndex        =   2
         Top             =   5880
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8070
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
         ColumnCount     =   3
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
         BeginProperty Column02 
            DataField       =   "fecha_llegada_almacen"
            Caption         =   "Fecha Llegada Almacen"
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
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2954,835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2984,882
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8070
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
         ColumnCount     =   4
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
         BeginProperty Column02 
            DataField       =   "fecha_llegada_almacen"
            Caption         =   "Fecha Llegada Almacen"
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
            DataField       =   "fecha_atencion"
            Caption         =   "Fecha Atencion"
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
               ColumnWidth     =   1649,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1980,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2399,811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2055,118
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   7
         Top             =   600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8070
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "codigo_pedido"
            Caption         =   "Codigo Pedido"
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
         BeginProperty Column03 
            DataField       =   "fecha_atencion"
            Caption         =   "Fecha atención"
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
            DataField       =   "fecha_listo_envio"
            Caption         =   "Fecha listo envio"
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
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   4575
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   -71880
         Picture         =   "Tecnico de Almacen.frx":04B2
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   5880
         Picture         =   "Tecnico de Almacen.frx":08F4
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   2760
         Picture         =   "Tecnico de Almacen.frx":0D36
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   -68040
         Picture         =   "Tecnico de Almacen.frx":1178
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   -70680
         Picture         =   "Tecnico de Almacen.frx":15BA
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   -73320
         Picture         =   "Tecnico de Almacen.frx":19FC
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   -67560
         Picture         =   "Tecnico de Almacen.frx":1E3E
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   -69720
         Picture         =   "Tecnico de Almacen.frx":2280
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   -74040
         Picture         =   "Tecnico de Almacen.frx":26C2
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -67560
         Picture         =   "Tecnico de Almacen.frx":2B04
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -69720
         Picture         =   "Tecnico de Almacen.frx":2F46
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -71880
         Picture         =   "Tecnico de Almacen.frx":3388
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74040
         Picture         =   "Tecnico de Almacen.frx":37CA
         Top             =   5280
         Width           =   480
      End
   End
End
Attribute VB_Name = "Tecnico_Almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public region As String
Public cod_almacen As String

Private rs As ADODB.Recordset
Public rs_ordenes As ADODB.Recordset

Sub cargar_almacen(nif As String)

    MiCommand.CommandText = "select * from almacen where tecnico_almacen='" & nif & "'"
    Set rs = MiCommand.Execute
    region = rs("codigo_region")
    Me.cod_almacen = rs("codigo_almacen")
    rs.Close
End Sub

Private Sub bAtenderPed_Click()
    If Not rs_ordenes.EOF Then
        Me.Visible = False
        Atender_pedido.Cargar_pedido (rs_ordenes!codigo_pedido)
        Atender_pedido.Show vbModal
    Else
        MsgBox "No tiene pedidos que atender."
    End If
End Sub

Private Sub bPasoListoEnvio_Click()
    Dim rs_envio As ADODB.Recordset
    Dim cad_envio As String
    Dim consultagrid As String
    Dim i As Integer
    Dim completo As Boolean
    
    Set rs_envio = New ADODB.Recordset
    cad_envio = "SELECT * FROM Linea_pedido WHERE codigo_pedido =" & rs_ordenes("codigo_pedido") & ""
    
    If Not rs_ordenes.EOF Then
        'rs_envio.Open cad_envio, MiConexion, adOpenStatic, adLockOptimistic
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
                        
                        crea_record rs_contador, "select max(codigo_pedido) as ultimo from orden_pedido ", False
                        
                        num_ped = rs_contador("ultimo") + 1
                        
                        Dim rs_insert As ADODB.Recordset
                        Dim cad_insert As String
                        Set rs_insert = New ADODB.Recordset
                        cad_insert = "select * from orden_pedido where codigo_pedido = " & rs_ordenes("codigo_pedido") & ""
                        crea_record rs_insert, cad_insert, False
                        
                        'hemos de hacer el commit para que al insertar las filas en Linea_pedido no dé error de clave ajena
                        'MiConexion.BeginTrans
                        If oracle Then
                            MiCommand.CommandText = "INSERT INTO Orden_pedido (codigo_pedido, cliente, usuario_ventas, CP_envio, Pais_envio, Provincia_envio, Localidad_envio, Pta_envio, numero_envio, calle_envio, forma_pago, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion) VALUES (" & num_ped & ", '" & rs_insert("cliente") & "', '" & rs_insert("usuario_ventas") & "', " & rs_insert("CP_envio") & ", '" & rs_insert("Pais_envio") & "', '" & rs_insert("Provincia_envio") & "', '" & rs_insert("Localidad_envio") & "', " & rs_insert("Pta_envio") & ", " & rs_insert("numero_envio") & ", '" & rs_insert("calle_envio") & "', '" & rs_insert("forma_pago") & "',TO_Date( '" & rs_insert("fecha_elaboracion") & "', 'DD/MM/YYYY HH:MI:SS AM'), TO_Date( '" & rs_insert("fecha_llegada_almacen") & "', 'DD/MM/YYYY HH:MI:SS AM'),TO_Date( '" & rs_insert("fecha_atencion") & "', 'DD/MM/YYYY HH:MI:SS AM'))"
                        Else
                            MiCommand.CommandText = "INSERT INTO Orden_pedido (codigo_pedido, cliente, usuario_ventas, CP_envio, Pais_envio, Provincia_envio, Localidad_envio, Pta_envio, numero_envio, calle_envio, forma_pago, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion) VALUES (" & num_ped & ", '" & rs_insert("cliente") & "', '" & rs_insert("usuario_ventas") & "', " & rs_insert("CP_envio") & ", '" & rs_insert("Pais_envio") & "', '" & rs_insert("Provincia_envio") & "', '" & rs_insert("Localidad_envio") & "', " & rs_insert("Pta_envio") & ", " & rs_insert("numero_envio") & ", '" & rs_insert("calle_envio") & "', '" & rs_insert("forma_pago") & "', #" & rs_insert("fecha_elaboracion") & "#, #" & rs_insert("fecha_llegada_almacen") & "#, #" & rs_insert("fecha_atencion") & "#)"
                        End If
                        MiCommand.Execute
                        'MiConexion.CommitTrans
                        
                        
                        
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
    End If
    
End Sub
Sub pasar_envio()

    If oracle Then
        MiCommand.CommandText = "UPDATE orden_pedido SET fecha_listo_envio = TO_Date( '" & Date & "', 'DD/MM/YYYY HH:MI:SS AM') WHERE codigo_pedido = " & rs_ordenes("codigo_pedido") & ""
    Else
        MiCommand.CommandText = "UPDATE orden_pedido SET fecha_listo_envio = #" & Date & "# WHERE codigo_pedido = " & rs_ordenes("codigo_pedido") & ""
    End If
    MiCommand.Execute
                
    consultagrid = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion FROM Orden_pedido WHERE fecha_llegada_almacen IS NOT NULL AND fecha_atencion is NOT NULL and fecha_listo_envio is NULL and pais_envio in (select nombre from pais where codigo_region='" & region & "')"
                
    rs_ordenes.Close
    rs_ordenes.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
    If Not rs_ordenes.EOF Then
        Set DataGrid2.DataSource = rs_ordenes
    Else
        Set DataGrid2.DataSource = Null
    End If
    
End Sub

Private Sub btAtender_Click()
    'llama a atender pedido
    bAtenderPed_Click
End Sub

'el pedido esta en envios pero se quiere consultar/modificar sus datos
Private Sub btConsultarModificar_Click()
    Dim ped As Integer
    If Not rs_ordenes.EOF Then
        Me.Visible = False
        ped = rs_ordenes!codigo_pedido
        Atender_pedido.Cargar_pedido (ped)
        Atender_pedido.btPasarEnvio.Enabled = False
        Atender_pedido.Show vbModal
        '******************
        Dim rsaux As ADODB.Recordset
        Dim borralo As Boolean
        MiCommand.CommandText = "SELECT * From linea_pedido, producto, producto_almacen where codigo_pedido=" & ped & " and producto.referencia=linea_pedido.referencia and linea_pedido.referencia = producto_almacen.referencia and producto_almacen.almacen='" & Tecnico_Almacen.cod_almacen & "'"
        Set rsaux = MiCommand.Execute
        borralo = False
        If Not rsaux.EOF Then
            rsaux.MoveFirst
            While (Not rsaux.EOF And Not borralo)
                If CInt(rsaux("cant_asignada")) < CInt(rsaux("cantidad")) Then
                    borralo = True
                End If
                rsaux.MoveNext
            Wend
            If borralo Then
                If MsgBox("¿Está seguro que desea pasar el pedido '" & ped & "' a la lista de atendidos?", vbOKCancel) = 1 Then
                    'MiCommand.CommandText = "SELECT * FROM Orden_pedido WHERE codigo_pedido=" & rs_ordenes("codigo_pedido") & ""
                    Dim cons As String
                    cons = "SELECT * FROM Orden_pedido WHERE codigo_pedido=" & ped & ""
                    rsaux.Close
                    rsaux.Open cons, MiConexion, adOpenDynamic, adLockOptimistic
                    'Set rsaux = MiCommand.Execute
                    
                    rsaux("fecha_listo_envio") = Null
                    rsaux.Update
                    rs_ordenes.Requery
                End If
            End If
        End If
        '***********************
        Atender_pedido.btPasarEnvio.Enabled = True
    Else
        MsgBox "No hay pedidos para consultar."
    End If
End Sub

Private Sub btConsultarNoAtendidos_Click()
    
    On Error GoTo gridsinrs
        
        If Not rs_ordenes.EOF Then
            Me.Visible = False
            Consultar_pedidos_no_atendidos_CU05.Cargar_pedido rs_ordenes.Fields("codigo_pedido")
            Consultar_pedidos_no_atendidos_CU05.Show vbModal
            rs_ordenes.Requery
        Else
            MsgBox "No tiene datos que consultar."
        End If
        
    Exit Sub
    
gridsinrs:
        MsgBox "Se ha producido un error..."
    
End Sub

Private Sub btSalir2_Click()
    Unload Me
End Sub

Private Sub btSalir3_Click()
    Unload Me
End Sub

Private Sub btSalir4_Click()
    Unload Me
End Sub

Private Sub btCancelarPedido_Click()
    If rs_ordenes.RecordCount >= 1 And DataGrid1.Row >= 0 Then
        If MsgBox("¿Está seguro que desea eliminar el pedido '" & rs_ordenes("codigo_pedido") & "'?", vbOKCancel) = 1 Then
            rs_ordenes.Delete
        End If
    End If
End Sub

Private Sub btSalir_Click()
    Unload Me
End Sub
'este metodo cancela un pedido q ya esta en atención
'ha de actualizar los stocks
Private Sub btCancelPedEnAtencion_Click()
    If rs_ordenes.RecordCount >= 1 And DataGrid2.Row >= 0 Then
        If MsgBox("¿Está seguro que desea eliminar el pedido '" & rs_ordenes("codigo_pedido") & "'?", vbOKCancel) = 1 Then
            'ahora hay q restituir el stock q tenia asignado el pedido
            Dim consulta2 As String
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
                        
            consulta2 = "SELECT linea_pedido.codigo_pedido,linea_pedido.cant_asignada," & _
              "linea_pedido.referencia,linea_pedido.cantidad, producto_almacen.stock_asignado," & _
              "producto_almacen.stock,producto_almacen.stock - producto_almacen.stock_asignado as stock_disponible " & _
              "From linea_pedido, producto, producto_almacen where codigo_pedido=" & rs_ordenes("codigo_pedido") & _
              " and producto.referencia=linea_pedido.referencia and linea_pedido.referencia = producto_almacen.referencia " & _
              "and producto_almacen.almacen='" & Tecnico_Almacen.cod_almacen & "'"
            
            crea_record rs, consulta2, False
                        
            'hay q recorrer cada linea del pedido para ir liberando el stock de cada producto
            If Not rs.EOF Then
                rs.MoveFirst
                While Not rs.EOF
                    MiCommand.CommandText = "UPDATE producto_almacen SET stock_asignado=" & _
                        rs("stock_asignado") - rs("cant_asignada") & " WHERE referencia='" & _
                        rs("referencia") & "' and almacen = '" & Tecnico_Almacen.cod_almacen & "'"
                
                    MiCommand.Execute
                    rs.MoveNext
                Wend
            End If
            'hay q borrar las incidencias relacionadas
            'MiCommand.CommandText = "DELETE * FROM INCIDENCIAS WHERE codigo_pedido = " & rs_ordenes("codigo_pedido") & ""
            'MiCommand.Execute
            rs_ordenes.Delete
        End If
    End If
End Sub
'este metodo cancela un pedido q esta listo para ser enviado
'ha de actualizar los stocks
Private Sub btCancelPedListoEnvio_Click()
    If rs_ordenes.RecordCount >= 1 And DataGrid3.Row >= 0 Then
        If MsgBox("¿Está seguro que desea eliminar el pedido '" & rs_ordenes("codigo_pedido") & "'?", vbOKCancel) = 1 Then
            'ahora hay q restituir el stock q tenia asignado el pedido
            Dim consulta2 As String
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
                        
            consulta2 = "SELECT linea_pedido.codigo_pedido,linea_pedido.cant_asignada," & _
              "linea_pedido.referencia,linea_pedido.cantidad, producto_almacen.stock_asignado," & _
              "producto_almacen.stock,producto_almacen.stock - producto_almacen.stock_asignado as stock_disponible " & _
              "From linea_pedido, producto, producto_almacen where codigo_pedido=" & rs_ordenes("codigo_pedido") & _
              " and producto.referencia=linea_pedido.referencia and linea_pedido.referencia = producto_almacen.referencia " & _
              "and producto_almacen.almacen='" & Tecnico_Almacen.cod_almacen & "'"
            
            crea_record rs, consulta2, False
                        
            'hay q recorrer cada linea del pedido para ir liberando el stock de cada producto
            If Not rs.EOF Then
                rs.MoveFirst
                While Not rs.EOF
                    MiCommand.CommandText = "UPDATE producto_almacen SET stock_asignado=" & _
                        rs("stock_asignado") - rs("cant_asignada") & " WHERE referencia='" & _
                        rs("referencia") & "' and almacen = '" & Tecnico_Almacen.cod_almacen & "'"
                    
                    MiCommand.Execute
                    rs.MoveNext
                Wend
            End If
            'antes hay q borrar todas las incidencias relacionadas
            'MiCommand.CommandText = "DELETE * FROM Incidencias WHERE codigo_pedido = " & rs_ordenes("codigo_pedido") & ""
            'MiCommand.Execute
            rs_ordenes.Delete
        End If
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    'Debug.Print DataGrid1.Text; DataGrid1.Row; DataGrid1.Col
    
End Sub

Private Sub Form_Load()
    'elige la primra pestaña
    Tecnico_Almacen.SSTab1.Tab = 0
    'carga los datos para la primera pestaña
    SSTab1_Click (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Identificacion.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    Set rs_ordenes = New ADODB.Recordset
    
    Dim consultagrid As String
    
    Select Case SSTab1.Tab
        Case 0
            consultagrid = "SELECT * FROM Orden_pedido WHERE fecha_llegada_almacen IS NOT NULL AND fecha_atencion is NULL and pais_envio in (select nombre from pais where codigo_region='" & region & "')"
            
            rs_ordenes.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
            If Not rs_ordenes.EOF Then
                Set DataGrid1.DataSource = rs_ordenes
            End If
    
        Case 1
            consultagrid = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion FROM Orden_pedido WHERE fecha_llegada_almacen IS NOT NULL AND fecha_atencion is NOT NULL and fecha_listo_envio is NULL and pais_envio in (select nombre from pais where codigo_region='" & region & "')"
            
            rs_ordenes.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
            If Not rs_ordenes.EOF Then
                Set DataGrid2.DataSource = rs_ordenes
            End If

        Case 2
            consultagrid = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen, fecha_atencion, fecha_listo_envio FROM Orden_pedido WHERE fecha_llegada_almacen IS NOT NULL AND fecha_atencion is NOT NULL and fecha_listo_envio is not NULL and fecha_salida_almacen is null and pais_envio in (select nombre from pais where codigo_region='" & region & "')"
            
            rs_ordenes.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
                        
            If Not rs_ordenes.EOF Then
                Set DataGrid3.DataSource = rs_ordenes
            End If
            
    End Select

End Sub
