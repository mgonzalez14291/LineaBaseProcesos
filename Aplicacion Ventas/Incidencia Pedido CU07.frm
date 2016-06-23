VERSION 5.00
Begin VB.Form Incidencia_pedido 
   Caption         =   "Incidencia de pedido"
   ClientHeight    =   4680
   ClientLeft      =   4650
   ClientTop       =   2835
   ClientWidth     =   7335
   Icon            =   "Incidencia Pedido CU07.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bConsult_incidencias 
      Caption         =   "Consultar Incidencias"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones de la Incidencia "
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   7095
      Begin VB.TextBox cObservaciones 
         Height          =   855
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Datos_incidencia 
      Caption         =   "Datos de la Incidencia "
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7095
      Begin VB.TextBox cCod_pedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox cFechaInci 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox cIncidencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lCodigo 
         Caption         =   "Código del Pedido"
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Incidencia"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Código de Incidencia"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de Empleado "
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox cNomEmpleado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox cNif 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton bSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton bGuardar 
      Caption         =   "&Guardar "
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "Incidencia Pedido CU07.frx":014A
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5760
      Picture         =   "Incidencia Pedido CU07.frx":058C
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "Incidencia Pedido CU07.frx":09CE
      Top             =   3600
      Width           =   480
   End
End
Attribute VB_Name = "Incidencia_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bConsult_incidencias_Click()
    Me.Visible = False
    Consultar_Incidencias.Cargar_incidencias
    Consultar_Incidencias.Show vbModal
End Sub

Private Sub bGuardar_Click()
    guardar_incidencia
    MiConexion.CommitTrans
    Unload Me
    Elaborar_Pedido_Nuevo_Modif.Visible = True
End Sub

Private Sub bSalir_Click()
    MiConexion.RollbackTrans
    Unload Me
    Elaborar_Pedido_Nuevo_Modif.Visible = True
End Sub

Sub cargar_datos_incidencia(cod_ped As Integer)
    
    Dim rs As ADODB.Recordset
    Dim consulta As String
    
    Me.cNif = Elaborar_pedido_CU02.cCodigoRepres
    Me.cNomEmpleado = Elaborar_pedido_CU02.cNombreRepres
    Me.cCod_pedido = cod_ped
    Me.cFechaInci = Date
        
    'tambien se le ha de dar ya un codigo a la incidencia
    consulta = "SELECT max(codigo_incidencia)+1 as ultimo FROM Incidencias"
    crea_record rs, consulta, False
    If IsNull(rs("ultimo")) Then
        Me.cIncidencia = 1
    Else
        Me.cIncidencia.Text = rs("ultimo")
    End If
    
End Sub

Private Sub guardar_incidencia()
            If Len(Me.cObservaciones) = 0 Then
                MsgBox "El campo de observaciones no puede estar vacío, la incidencia no se registrará", vbExclamation, "Aviso de Error"
            Else
                MiCommand.CommandText = "INSERT INTO Incidencias (codigo_incidencia,codigo_pedido,fecha_incidencia,nif_creador,creador,observaciones) VALUES ('" & Me.cIncidencia.Text & "','" & Me.cCod_pedido.Text & "',#" & Me.cFechaInci.Text & "#, '" & Me.cNif.Text & "','" & Me.cNomEmpleado.Text & "','" & Me.cObservaciones & "')"
                MiCommand.Execute
                
            End If
End Sub

Private Sub Form_Load()
    MiConexion.BeginTrans
End Sub
