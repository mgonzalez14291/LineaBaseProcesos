VERSION 5.00
Begin VB.Form Consultar_catalogo 
   Caption         =   "Consulta del Catálogo"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "Consultar Catalogo CU04.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bElegir 
      Cancel          =   -1  'True
      Caption         =   "Seleccionar &Producto"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Salir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Producto"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox cReferencia 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox cPrecio 
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox cDescripcion 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox cNom_product 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   2655
         Left            =   7080
         ScaleHeight     =   2595
         ScaleWidth      =   2835
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Descripcion 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del producto"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7080
      Picture         =   "Consultar Catalogo CU04.frx":030A
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2640
      Picture         =   "Consultar Catalogo CU04.frx":074C
      Top             =   3240
      Width           =   480
   End
End
Attribute VB_Name = "Consultar_catalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bElegir_Click()
    'se ha seleccionado el producto por los que hay que pasar
    'los datos al otro formulario
    Elaborar_Pedido_Nuevo_Modif.cCod_Articulo.Text = cReferencia.Text
    Elaborar_Pedido_Nuevo_Modif.cDescripcion.Text = cDescripcion.Text
    Elaborar_Pedido_Nuevo_Modif.cCantidad.Text = "1"
    Elaborar_Pedido_Nuevo_Modif.cPrecioUnidad.Text = cPrecio.Text
End Sub

Private Sub Salir_Click()
    Unload Me
    Elaborar_Pedido_Nuevo_Modif.Visible = True
End Sub
