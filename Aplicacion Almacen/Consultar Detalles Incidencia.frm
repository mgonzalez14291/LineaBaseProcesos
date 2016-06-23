VERSION 5.00
Begin VB.Form Consultar_Detalles_Incidencia 
   Caption         =   "Detalles de Incidencia"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "Consultar Detalles Incidencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de Empleado "
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox cNif 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox cNomEmpleado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label16 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Datos_incidencia 
      Caption         =   "Datos de la Incidencia "
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      Begin VB.TextBox cIncidencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox cFechaInci 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox cCod_pedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Código de Incidencia"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Incidencia"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lCodigo 
         Caption         =   "Código del Pedido"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones de la Incidencia "
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   7095
      Begin VB.TextBox cObservaciones 
         Height          =   855
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5760
      Picture         =   "Consultar Detalles Incidencia.frx":0442
      Top             =   3600
      Width           =   480
   End
End
Attribute VB_Name = "Consultar_Detalles_Incidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bSalir_Click()
    Unload Me
    Consultar_Incidencias.Visible = True
End Sub

Sub cargar_detalles_incidencia(codinci As String)
        Dim consulta As String

        Set rsin = New ADODB.Recordset
        consulta = "SELECT * FROM Incidencias WHERE codigo_incidencia = " & codinci & " "
        rsin.Open consulta, MiConexion, adOpenDynamic, adLockOptimistic
        Consultar_Detalles_Incidencia.cIncidencia = rsin("codigo_incidencia")
        Consultar_Detalles_Incidencia.cCod_pedido = rsin("codigo_pedido")
        Consultar_Detalles_Incidencia.cFechaInci = rsin("fecha_incidencia")
        Consultar_Detalles_Incidencia.cNif = rsin("nif_creador")
        Consultar_Detalles_Incidencia.cNomEmpleado = rsin("creador")
        Consultar_Detalles_Incidencia.cObservaciones = rsin("observaciones")

End Sub
