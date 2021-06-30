VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMNUELO 
   Caption         =   "PROPIEDADES DE USUARIO - Bazar Jessica"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   6480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BASEINV.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BASEINV.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DUENO"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TXTCEDN 
      DataField       =   "CEDULA"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox TXTNOMN 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Image Image3 
      Height          =   1860
      Left            =   360
      Picture         =   "FRMNUELO.frx":0000
      Top             =   -120
      Width           =   3825
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   8280
      Picture         =   "FRMNUELO.frx":231F
      Top             =   5160
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   120
      Picture         =   "FRMNUELO.frx":46CC
      Top             =   5160
      Width           =   1620
   End
   Begin VB.Image Image6 
      Height          =   750
      Left            =   5760
      Picture         =   "FRMNUELO.frx":6935
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   3840
      Picture         =   "FRMNUELO.frx":8970
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   1920
      Picture         =   "FRMNUELO.frx":A5ED
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image Image15 
      Height          =   690
      Left            =   5760
      Picture         =   "FRMNUELO.frx":C62A
      Top             =   3960
      Width           =   780
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   4680
      Picture         =   "FRMNUELO.frx":D648
      Top             =   3960
      Width           =   780
   End
   Begin VB.Image Image13 
      Height          =   690
      Left            =   3600
      Picture         =   "FRMNUELO.frx":E3C8
      Top             =   3960
      Width           =   780
   End
   Begin VB.Image Image16 
      Height          =   690
      Left            =   2520
      Picture         =   "FRMNUELO.frx":F140
      Top             =   3960
      Width           =   780
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   7920
      X2              =   7920
      Y1              =   5040
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   10440
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIOS REGISTRADOS"
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
      Height          =   855
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   10800
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      X1              =   120
      X2              =   10920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "FRMNUELO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub CMDCREAR_Click()
Adodc1.Recordset.AddNew
'copia y pega esto en el nuevo moveup para dar funcionamiento al imagen

End Sub

Private Sub CMDREG_Click()
If MsgBox("Esta seguro que desea regresar al formulario de login?", vbQuestion + vbYesNo, "Propiedades de Usuario") = vbYes Then
FRMNUELO.Hide
FRMLOGIN.Show
    End If
End Sub

Private Sub Command1_Click()
If TXTNOMN.Text = "" Or TXTCEDN.Text = "" Then
    
    MsgBox "Llenar todos los campos del nuevo usuario.", vbInformation, "Dialogo"
    Adodc1.Recordset.Delete
    Else
    MsgBox "El nuevo usuario ha sido registrado.", vbInformation, "Dialogo"
    End If
End Sub

Private Sub Command2_Click()
     
Adodc1.Recordset.Fields("NOMBRE") = TXTNOMN.Text
Adodc1.Recordset.Fields("CEDULA") = TXTCEDN.Text
Adodc1.Recordset.Update
MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command7_Click()
If MsgBox("Esta seguro que desea eliminar este registro del usuario?", vbQuestion + vbYesNo) = vbYes Then
        Adodc1.Recordset.Delete

    End If
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Form_Load()
Dim CN As New ADODB.Connection
Dim rs As New ADODB.Recordset
Adodc1.LockType = adLockReadOnly
rs.LockType = adLockOptimistic
FRMNUELO.Picture = LoadPicture(App.Path & "\IMG\tst.jpg")

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\img\cr2.jpg")
    'convertir el imagen al boton blanco cuando el usuario presiona el boton
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\img\cr1.jpg")
    'resetear al imagen original, buton azul con fuente blanco
    Adodc1.Recordset.AddNew
    'pegar el codigo del comando boton previo al mouse up
    
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri1.jpg")
End Sub
Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri0.jpg")
    Adodc1.Recordset.MovePrevious
End Sub


Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig1.jpg")
End Sub
Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig0.jpg")
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi1.jpg")
End Sub
Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi0.jpg")
    Adodc1.Recordset.MoveLast
End Sub
Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in1.jpg")
End Sub
Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in0.jpg")
    Adodc1.Recordset.MoveFirst
End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\img\reg2.jpg")
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\img\reg1.jpg")
    If MsgBox("Esta seguro que desea regresar al formulario login?", vbQuestion + vbYesNo) = vbYes Then
    FRMNUELO.Hide
    FRMLOGIN.Show
    End If
    
End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\img\gua1.jpg")

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\img\gua0.jpg")
        If TXTNOMN.Text = "" Or TXTCEDN.Text = "" Then
    
    MsgBox "Llenar todos los campos de datos de los productos", vbInformation, "Dialogo"
    Adodc1.Recordset.Delete
    Else
    MsgBox "El registro ha sido guardado.", vbInformation, "Dialogo"
    End If
End Sub


Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\img\ed1.jpg")

End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image5.Picture = LoadPicture(App.Path & "\img\ed0.jpg")
     
    Adodc1.Recordset.Fields("NOMBRE") = TXTNOMN.Text
    Adodc1.Recordset.Fields("CEDULA") = TXTCEDN.Text
    
    Adodc1.Recordset.Update
    MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli1.jpg")
End Sub



Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli0.jpg")
    If MsgBox("Esta seguro que desea eliminar este registro del usuario?", vbQuestion + vbYesNo) = vbYes Then
        Adodc1.Recordset.Delete
        End If
        
    
End Sub

