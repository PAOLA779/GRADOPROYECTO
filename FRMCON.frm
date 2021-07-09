VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMCON 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "BUSCAR"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Hasta 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox desde 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FRMCON.frx":0000
      Left            =   960
      List            =   "FRMCON.frx":000D
      TabIndex        =   3
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "FRMCON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer
Private Sub Combo1_Click()
  If Combo1.Text = "Desde" Then
        desde.Visible = True
        Hasta.Visible = False
    End If
    If Combo1.Text = "Hasta" Then
        desde.Visible = False
        Hasta.Visible = True
    End If
    If Combo1.Text = "Desde/Hasta" Then
        desde.Visible = True
        Hasta.Visible = True
    End If
End Sub

Private Sub Command1_Click()
With RSVENTAS_ELIMINADAS
    .AddNew
    !IDVENTAS = DataGrid1.Columns(0).Text
    !FECHA = DataGrid1.Columns(1).Text
    !CEDULACLIENTE = DataGrid1.Columns(2).Text
    !CEDULADUENO = DataGrid1.Columns(3).Text
    .UpdateBatch
    End With
    q = rsFactura.RecordCount
    For X = 1 To q
    With RSFACTURA_ELIMINADAS
    .Requery
    .AddNew
    !IDFACTURA = DataGrid2.Columns(0).Text
    !IDPRODUCTO = DataGrid2.Columns(1).Text
    !CANTIDAD = DataGrid2.Columns(2).Text
    !PRECIO = DataGrid2.Columns(3).Text
    !IDVENTAS = DataGrid2.Columns(4).Text
    .UpdateBatch
    End With
    rsFactura.MoveNext
    Next
    With RSVEN
    .Delete
    .MoveFirst
    End With
    For X = 1 To q
    With rsFactura
 
    .Requery
    .Delete
    .MoveNext
    If .EOF Then Exit Sub
    End With
    Next
    
End Sub

Private Sub DataGrid1_Click()

  label1 = DataGrid1.Columns(0).Text
    With rsFactura
        Dim s As String
        s = "%" & label1.Caption & "%"
        If .State = 1 Then .Close
        .Open "Select * From FACTURA Where [IDVENTAS] Like '" & s & "'"
        Set DataGrid2.DataSource = rsFactura
    End With
End Sub


Private Sub Form_Load()
FACTURA_ELIMINADA
VENTAS_ELIMINADAS
tablaVENTAS
factura
Set DataGrid1.DataSource = RSVEN
End Sub
