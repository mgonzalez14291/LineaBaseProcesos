VERSION 5.00
Begin VB.Form Identificacion 
   Caption         =   "Identificacion"
   ClientHeight    =   2415
   ClientLeft      =   5265
   ClientTop       =   4920
   ClientWidth     =   6540
   Icon            =   "Identificacion.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Identificacion.frx":030A
   ScaleHeight     =   2415
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Base de Datos "
      Height          =   1095
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox Check1 
         Caption         =   "Oracle"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Access (por defecto)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton bCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox cContraseña 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox cUsuario 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1560
      Picture         =   "Identificacion.frx":074C
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4680
      Picture         =   "Identificacion.frx":0B8E
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Identificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bAceptar_Click()
    Dim id, cargoid As String
    Dim claveid, consulta As String
    Dim num As Integer
    'definimos el objeto necesario para realizar consultas
    Dim adRst As Recordset
    
    On Error GoTo situacion_de_error
    
    Set adRst = New ADODB.Recordset
        
    oracle = Me.Check1.Value
    conecta
    
    id = cUsuario.Text
    claveid = cContraseña.Text
    'COMPROBACION DEL LOGIN Y CONTRASEÑA
    consulta = "SELECT nif,password,cargo,nombre FROM empleado WHERE login='" & id & "'"
    crea_record adRst, consulta, False
    
    If adRst.RecordCount = 0 Then
    MsgBox "Error , este usuario no existe."
    Else
     If claveid <> adRst.Fields("password") Then
     MsgBox "Error, contraseña incorrecta."
     Else
            Me.Visible = False
            cargoid = adRst.Fields("cargo")
            Select Case cargoid
                Case "Tecnico"
                    Tecnico_Almacen.cargar_almacen adRst("nif")
                    Tecnico_Almacen.Caption = Tecnico_Almacen.Caption + ": " + adRst.Fields("nombre")
                    Tecnico_Almacen.Show
        
                Case "Representante"
                    Elaborar_pedido_CU02.Show
                    
            
                Case "Operadora"
                    Elaborar_pedido_CU02.Show
                        
                Case Else
                    MsgBox "Error ud. no esta identificado"
                    Me.Visible = True
            End Select
     End If
    GoTo fin_de_identificacion
    End If
situacion_de_error:
    MsgBox ("Error de conexión, la base de datos puede no estar disponible o se trata de un usuario de otro subsistema")
    Identificacion.Show
    
fin_de_identificacion:
    'se activa el formulario pricipal del subsistema
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    desconecta
End Sub
