VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Elegir_nombre 
   Caption         =   "Selección cliente"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "Elegir nombre.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid gridnom 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3201
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
         DataField       =   "nombre"
         Caption         =   "Nombre"
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
         DataField       =   "codigo"
         Caption         =   "Código"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   3135,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1950,236
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1200
      Picture         =   "Elegir nombre.frx":030A
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4320
      Picture         =   "Elegir nombre.frx":074C
      Top             =   2400
      Width           =   480
   End
End
Attribute VB_Name = "Elegir_nombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bAceptar_Click()

Dim cod_cli As String
Dim consulta As String
Dim adRst As Recordset
Dim RstGrid As Recordset
Dim RstGrid2 As Recordset
gridnom.Col = 1
cod_cli = gridnom.Text

consulta = "SELECT * FROM cliente WHERE codigo = '" & cod_cli & "'"
crea_record adRst, consulta, False


Elaborar_pedido_CU02.cNombre.Text = adRst("nombre")
Elaborar_pedido_CU02.cDni.Text = adRst("nif_cif")
Elaborar_pedido_CU02.cCodigoCli.Text = adRst("codigo")
Elaborar_pedido_CU02.cPerContacto.Text = IIf(adRst("persona_contacto") <> "", adRst("persona_contacto"), "")
Elaborar_pedido_CU02.cTlfContacto.Text = IIf(adRst("tlf_pers_contacto") <> "", adRst("tlf_pers_contacto"), "")
Elaborar_pedido_CU02.cCalle.Text = IIf(adRst("calle") <> "", adRst("calle"), "")
Elaborar_pedido_CU02.cNum.Text = IIf(adRst("numero") <> "", adRst("numero"), "")
Elaborar_pedido_CU02.cPta.Text = IIf(adRst("puerta") <> "", adRst("puerta"), "")
Elaborar_pedido_CU02.cLocalidad.Text = IIf(adRst("localidad") <> "", adRst("localidad"), "")
Elaborar_pedido_CU02.cProvincia.Text = IIf(adRst("provincia") <> "", adRst("provincia"), "")
Elaborar_pedido_CU02.cCP.Text = IIf(adRst("CP") <> "", adRst("CP"), "")
Elaborar_pedido_CU02.cTelefono.Text = IIf(adRst("telefono") <> "", adRst("telefono"), "")
Elaborar_pedido_CU02.cFax.Text = IIf(adRst("fax") <> "", adRst("fax"), "")
Elaborar_pedido_CU02.cEmail.Text = IIf(adRst("email") <> "", adRst("email"), "")
Elaborar_pedido_CU02.cPais.Text = IIf(adRst("pais") <> "", adRst("pais"), "")
                
                'Para llenar el grid
consultagrid = "SELECT codigo_pedido, fecha_elaboracion FROM orden_pedido WHERE cliente ='" & cod_cli & "' and fecha_llegada_almacen is null"
Set RstGrid = New ADODB.Recordset
RstGrid.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
Set Elaborar_pedido_CU02.DataGrid1.DataSource = RstGrid

consultagrid2 = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & cod_cli & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
crea_record RstGrid2, consultagrid2, False
Set Elaborar_pedido_CU02.DataGrid2.DataSource = RstGrid2
                        
Unload Me
          
End Sub

Private Sub bCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim nom As String
Dim consulta As String
Dim nifrepres As String
Dim esoperadora As Boolean
Dim R As Recordset

nom = Elaborar_pedido_CU02.cNombre.Text
nifrepres = Elaborar_pedido_CU02.cCodigoRepres.Text
esoperadora = IIf(Elaborar_pedido_CU02.cEsoperadora = 0, False, True)

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

    
crea_record R, consulta, False
Set gridnom.DataSource = R

End Sub


Private Sub gridnom_DblClick()

Dim cod_cli As String
Dim consulta As String
Dim adRst As Recordset
Dim RstGrid As Recordset
Dim RstGrid2 As Recordset
gridnom.Col = 1
cod_cli = gridnom.Text

consulta = "SELECT * FROM cliente WHERE codigo = '" & cod_cli & "'"
crea_record adRst, consulta, False


Elaborar_pedido_CU02.cNombre.Text = adRst("nombre")
Elaborar_pedido_CU02.cDni.Text = adRst("nif_cif")
Elaborar_pedido_CU02.cCodigoCli.Text = adRst("codigo")
Elaborar_pedido_CU02.cPerContacto.Text = IIf(adRst("persona_contacto") <> "", adRst("persona_contacto"), "")
Elaborar_pedido_CU02.cTlfContacto.Text = IIf(adRst("tlf_pers_contacto") <> "", adRst("tlf_pers_contacto"), "")
Elaborar_pedido_CU02.cCalle.Text = IIf(adRst("calle") <> "", adRst("calle"), "")
Elaborar_pedido_CU02.cNum.Text = IIf(adRst("numero") <> "", adRst("numero"), "")
Elaborar_pedido_CU02.cPta.Text = IIf(adRst("puerta") <> "", adRst("puerta"), "")
Elaborar_pedido_CU02.cLocalidad.Text = IIf(adRst("localidad") <> "", adRst("localidad"), "")
Elaborar_pedido_CU02.cProvincia.Text = IIf(adRst("provincia") <> "", adRst("provincia"), "")
Elaborar_pedido_CU02.cCP.Text = IIf(adRst("CP") <> "", adRst("CP"), "")
Elaborar_pedido_CU02.cTelefono.Text = IIf(adRst("telefono") <> "", adRst("telefono"), "")
Elaborar_pedido_CU02.cFax.Text = IIf(adRst("fax") <> "", adRst("fax"), "")
Elaborar_pedido_CU02.cEmail.Text = IIf(adRst("email") <> "", adRst("email"), "")
Elaborar_pedido_CU02.cPais.Text = IIf(adRst("pais") <> "", adRst("pais"), "")
                
                'Para llenar el grid
consultagrid = "SELECT codigo_pedido, fecha_elaboracion FROM orden_pedido WHERE cliente ='" & cod_cli & "' and fecha_llegada_almacen is null"
Set RstGrid = New ADODB.Recordset
RstGrid.Open consultagrid, MiConexion, adOpenStatic, adLockOptimistic
Set Elaborar_pedido_CU02.DataGrid1.DataSource = RstGrid

consultagrid2 = "SELECT codigo_pedido, fecha_elaboracion, fecha_llegada_almacen FROM orden_pedido WHERE cliente ='" & cod_cli & "' and fecha_llegada_almacen is not null and fecha_salida_almacen is null"
crea_record RstGrid2, consultagrid2, False
Set Elaborar_pedido_CU02.DataGrid2.DataSource = RstGrid2
                        
Unload Me
          
End Sub

