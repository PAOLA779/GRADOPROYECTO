VERSION 5.00
Begin VB.Form FRMMENU 
   Caption         =   "BAZAR JESSICA - Menú General"
   ClientHeight    =   4620
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FRMFACT.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image17 
      Height          =   630
      Left            =   11040
      Picture         =   "FRMFACT.frx":5E39
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   1380
      Left            =   2400
      Picture         =   "FRMFACT.frx":68CB
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona en el menú el formulario que desea entrar."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   1680
      Top             =   2040
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   3720
      Picture         =   "FRMFACT.frx":785B
      Top             =   120
      Width           =   3825
   End
   Begin VB.Menu mnuinventario 
      Caption         =   "Inventario"
   End
   Begin VB.Menu mnuusuarios 
      Caption         =   "Usuarios"
      Begin VB.Menu mnulogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnupropiedades 
         Caption         =   "Propiedades"
      End
   End
   Begin VB.Menu mnuventas 
      Caption         =   "Ventas"
   End
   Begin VB.Menu mnureporte 
      Caption         =   "Reporte"
      Begin VB.Menu mnudatainventario 
         Caption         =   "Data Report de Inventario"
      End
   End
End
Attribute VB_Name = "FRMMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image17.Picture = LoadPicture(App.Path & "\img\X1.jpg")
End Sub

Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image17.Picture = LoadPicture(App.Path & "\img\X0.jpg")
    If MsgBox("Esta seguro que desea cerrar el menú?", vbQuestion + vbYesNo, "Menú") = vbYes Then
        
            End
    End If
End Sub


Private Sub mnudatainventario_Click()
Set rs = CN.Execute("select *from inventario")
    If rs.EOF = False Then
    Set DRINV.DataSource = rs
    DRINV.Show
End If
End Sub

Private Sub mnufactura_Click()
'Set rs = CN.Execute("select *from factura")
    'If rs.EOF = False Then
    'Set DRFACTURA.DataSource = rs
    DRFACTURA.Show
'End If
End Sub

Private Sub mnuinventario_Click()
FRMINV.Show
FRMMENU.Hide

End Sub

Private Sub mnulogin_Click()
FRMLOGIN.Show
FRMMENU.Hide
End Sub

Private Sub mnupropiedades_Click()
FRMNUELO.Show
FRMMENU.Hide
End Sub

Private Sub mnuventas_Click()
FRMVENTAS.Show
FRMMENU.Hide
End Sub
