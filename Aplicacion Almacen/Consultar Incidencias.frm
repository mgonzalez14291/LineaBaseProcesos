VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Consultar_Incidencias 
   Caption         =   "Consultar Incidencias"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   Icon            =   "Consultar Incidencias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bConsultar 
      Caption         =   "Consultar Detalles"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incidencias Registradas "
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin MSDataGridLib.DataGrid Grid_Incidencias 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5741
         _Version        =   393216
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
            DataField       =   "codigo_incidencia"
            Caption         =   "Codigo Incidencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "fecha_incidencia"
            Caption         =   "Fecha Incidencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "codigo_pedido"
            Caption         =   "Codigo Pedido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1679,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1709,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1680
      Picture         =   "Consultar Incidencias.frx":014A
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4680
      Picture         =   "Consultar Incidencias.frx":058C
      Top             =   3960
      Width           =   480
   End
End
Attribute VB_Name = "Consultar_Incidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bConsultar_Click()
    Dim rs As ADODB.Recordset
    
    Set rs = Me.Grid_Incidencias.DataSource
    If rs.RecordCount >= 1 And Me.Grid_Incidencias.Row >= 0 Then
       ' Me.Visible = False
        Me.Grid_Incidencias.Col = 0
        Me.Visible = False
        Consultar_Detalles_Incidencia.cargar_detalles_incidencia (Me.Grid_Incidencias.Text)
        Consultar_Detalles_Incidencia.Show vbModal
    End If

    
End Sub

Private Sub bSalir_Click()
    Unload Me
    Atender_pedido.Visible = True
End Sub

Sub Cargar_incidencias()
    
    Dim consulta As String

    Set rsinci = New ADODB.Recordset
    consulta = "SELECT * FROM Incidencias"
    rsinci.Open consulta, MiConexion, adOpenDynamic, adLockOptimistic
    Set Me.Grid_Incidencias.DataSource = rsinci
    Me.Grid_Incidencias.Refresh

End Sub

