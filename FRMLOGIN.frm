VERSION 5.00
Begin VB.Form FRMLOGIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN - Bazar Jessica"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRMLOGIN.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "AGREGAR NUEVO USUARIO + PROPIEDADES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox TXTNOM 
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
      Left            =   3840
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox TXTCED 
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
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   600
      Shape           =   3  'Circle
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   600
      Left            =   1440
      Picture         =   "FRMLOGIN.frx":5E39
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   1440
      Picture         =   "FRMLOGIN.frx":A5A9
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   2400
      Picture         =   "FRMLOGIN.frx":C44A
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   4920
      Picture         =   "FRMLOGIN.frx":CE11
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   2520
      Picture         =   "FRMLOGIN.frx":D774
      Top             =   0
      Width           =   3825
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
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
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
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[!] INGRESE EL NOMBRE Y NUMERO DE CEDULA DEL PROPIETARIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   8895
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub CMDLOGIN_Click()
     
    If TXTNOM = "" And TXTCED = "" Then
    MsgBox "Llenar todos los campos indicados.", vbInformation, "Dialogo"
    ElseIf TXTNOM = "" Then
    MsgBox "Llenar el campo de nombre", vbInformation, "Dialogo"
    ElseIf TXTCED = "" Then
    MsgBox "Llenar el campo de cedula", vbInformation, "Dialogo"
    ElseIf Not (IsNumeric(TXTCED.Text)) Then
    MsgBox "Llenar el campo de cedula correcta con numeros", vbInformation, "Dialogo"
    TXTCED = ""
    Else
    
    rs.Requery 'Refrescar la tabla
    rs.Find "NOMBRE='" & (TXTNOM.Text) & "'", , , 1
    'Validad que el usuario exista para poder borrarlo
        If rs.EOF Then
            MsgBox "Los datos ingresados no se encuentra dentro del base de datos.", vbInformation, "Eliminar registro"
            Exit Sub 'Termina el procedimiento
        ElseIf rs!CEDULA = TXTCED.Text Then
            FRMINV.Show
            FRMLOGIN.Hide
            a = TXTCED.Text
        End If
    End If
End Sub

Private Sub Command1_Click()

If TXTNOM = "" And TXTCED = "" Then
    MsgBox "Llenar todos los campos indicados, para poder crear un nuevo usuario.", vbInformation, "Dialogo"
    ElseIf TXTNOM = "" Then
    MsgBox "Llenar el campo de nombre, para poder crear un nuevo usuario.", vbInformation, "Dialogo"
    ElseIf TXTCED = "" Then
    MsgBox "Llenar el campo de cedula, para poder crear un nuevo usuario.", vbInformation, "Dialogo"
    ElseIf Not (IsNumeric(TXTCED.Text)) Then
    MsgBox "Llenar el campo de cedula correcta con numeros, para poder crear un nuevo usuario.", vbInformation, "Dialogo"
    TXTCED = ""
    Else
    
    rs.Requery 'Refrescar la tabla
    rs.Find "NOMBRE='" & (TXTNOM.Text) & "'", , , 1
    'Validad que el usuario exista para poder borrarlo
        If rs.EOF Then
            MsgBox "Los datos ingresados no se encuentra dentro del base de datos. Ingresar los datos necesarios para poder crear un nuevo usuario.", vbInformation, "Eliminar registro"
            Exit Sub 'Termina el procedimiento
        ElseIf rs!CEDULA = TXTCED.Text Then
            FRMNUELO.Show
            FRMLOGIN.Hide
        End If
    End If
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
CON.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\BASEINV.mdb;Persist Security Info=False"


rs.Source = "DUENO"
rs.Open "select * from DUENO", CON
rs.MoveFirst
End Sub

Private Sub CMDSALIR_Click()
End
End Sub



Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img0\sal1.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img0\sal0.jpg")
End
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\img0\log1.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\img0\log0.jpg")
If TXTNOM = "" And TXTCED = "" Then
    MsgBox "Llenar todos los campos indicados.", vbInformation, "Dialogo"
    ElseIf TXTNOM = "" Then
    MsgBox "Llenar el campo de nombre", vbInformation, "Dialogo"
    ElseIf TXTCED = "" Then
    MsgBox "Llenar el campo de cedula", vbInformation, "Dialogo"
    ElseIf Not (IsNumeric(TXTCED.Text)) Then
    MsgBox "Llenar el campo de cedula correcta con numeros", vbInformation, "Dialogo"
    TXTCED = ""
    Else
    
    rs.Requery 'Refrescar la tabla
    rs.Find "NOMBRE='" & (TXTNOM.Text) & "'", , , 1
    'Validad que el usuario exista para poder borrarlo
        If rs.EOF Then
            MsgBox "Los datos ingresados no se encuentra dentro del base de datos.", vbInformation, "Eliminar registro"
            Exit Sub 'Termina el procedimiento
        ElseIf rs!CEDULA = TXTCED.Text Then
            FRMMENU.Show
            FRMLOGIN.Hide
            a = TXTCED.Text
        End If
    End If
End Sub
