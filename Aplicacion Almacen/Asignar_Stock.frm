VERSION 5.00
Begin VB.Form Asignar_Stock 
   Caption         =   "Asignar Stock"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "Asignar_Stock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btSalir 
      Caption         =   "A&bandonar"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock Articulo"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7815
      Begin VB.TextBox stock_asignado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox stock_disponible 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox stock_real 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Stock Disponible"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Asignado"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Stock Real"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton bAnyadir 
      Caption         =   "&Asignar Cantidad"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editar línea"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7815
      Begin VB.TextBox cant_solicitada 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox cCod_Articulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox cCantidad 
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox cNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Cant. Solicitada"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Código Artículo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad a Asignar"
         Height          =   255
         Left            =   6120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   840
      Picture         =   "Asignar_Stock.frx":0442
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6720
      Picture         =   "Asignar_Stock.frx":0884
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "Asignar_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Aceptar_cambio As Boolean

Private Sub bAnyadir_Click()
    If CLng(Me.cCantidad.Text) > Atender_pedido.rslinped!cantidad Then
        MsgBox "Esta asignando una cantidad mayor a la solicitado"
        Me.cCantidad.SetFocus
    Else
      If CLng(Me.cCantidad.Text) > CLng(stock_disponible) Then
        MsgBox "Esta asignando una cantidad mayor a la disponible"
        Me.cCantidad.SetFocus
      Else
        If cCantidad < 0 Then
            MsgBox "La cantidad no puede ser negativa"
        Else
            If CLng(Me.cCantidad.Text) >= 0 Then
                Aceptar_cambio = True
                Dim cantidad_asig_anterior As Long
                cantidad_asig_anterior = Atender_pedido.rslinped!cant_asignada
                Atender_pedido.rslinped!cant_asignada = CLng(Me.cCantidad.Text)
                'rslinped("stock_asignado") = rslinped("stock_asignado") + CLng(Me.cCantidad)
                Atender_pedido.rslinped.Update
                MiCommand.CommandText = "UPDATE producto_almacen SET stock_asignado=" & _
                    CLng(Me.stock_asignado) - CLng(cantidad_asig_anterior) + CLng(Me.cCantidad) & " WHERE referencia='" & _
                    Me.cCod_Articulo & "' and almacen = '" & Tecnico_Almacen.cod_almacen & "'"
                
                MiCommand.Execute
                
                Atender_pedido.rslinped.Requery
            End If
            
            Unload Me
        End If
      End If
    End If
End Sub

Private Sub btSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Aceptar_cambio = False
End Sub
