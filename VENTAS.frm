VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMVENTAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTAS - Bazar Jessica"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "VENTAS.frx":0000
   ScaleHeight     =   11085
   ScaleWidth      =   20400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTCEDULACLIENTE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2640
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc ADODCVENTAS 
      Height          =   735
      Left            =   16680
      Top             =   10800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
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
      RecordSource    =   "VENTAS"
      Caption         =   "VENTAS"
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
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "FINALIZAR VENTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox TXTID 
      Height          =   495
      Left            =   16680
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CMDCREARVEN 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "CREAR VENTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc ADODCFACTURAS 
      Height          =   615
      Left            =   16680
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      RecordSource    =   "FACTURA"
      Caption         =   "FACTURAS"
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
   Begin VB.TextBox TXTCAN2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   14
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton CMDB1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "BUSCAR ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "VENTAS.frx":5E39
      Height          =   3015
      Left            =   480
      TabIndex        =   12
      Top             =   8040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.TextBox TXTTOT 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox TXTIDPRO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TXTCAN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox TXTCOS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CMDAGR 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton CMDELI 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label IVA 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   13800
      TabIndex        =   31
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image15 
      Height          =   90
      Left            =   5520
      Picture         =   "VENTAS.frx":5E55
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   4545
   End
   Begin VB.Image Image14 
      Height          =   90
      Left            =   840
      Picture         =   "VENTAS.frx":6FE2
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   9345
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUB TOTAL"
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
      Height          =   495
      Left            =   5400
      TabIndex        =   30
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label LBLTOTGLOBAL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   29
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA DEL CLIENTE:"
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
      Height          =   855
      Left            =   360
      TabIndex        =   28
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image ImageM 
      Height          =   675
      Left            =   14280
      Picture         =   "VENTAS.frx":816F
      Top             =   4320
      Width           =   1725
   End
   Begin VB.Label LBLHORA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   24
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HORA"
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
      Height          =   495
      Left            =   840
      TabIndex        =   23
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label LBLFECHA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
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
      Height          =   495
      Left            =   840
      TabIndex        =   21
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label LBLIDVENTA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   20
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID VENTAS"
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
      Height          =   495
      Left            =   840
      TabIndex        =   19
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   5640
      X2              =   11040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image13 
      Height          =   1200
      Left            =   4680
      Picture         =   "VENTAS.frx":89E5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   600
      Left            =   16680
      Picture         =   "VENTAS.frx":8FF8
      Top             =   9360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Image12 
      Height          =   1110
      Left            =   14520
      Picture         =   "VENTAS.frx":9BA0
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Image Image11 
      Height          =   210
      Left            =   13800
      Picture         =   "VENTAS.frx":10314
      Stretch         =   -1  'True
      Top             =   9720
      Width           =   2745
   End
   Begin VB.Image Image10 
      Height          =   4455
      Left            =   16320
      Picture         =   "VENTAS.frx":114A1
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   225
   End
   Begin VB.Image Image9 
      Height          =   3270
      Left            =   16320
      Picture         =   "VENTAS.frx":116DE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   225
   End
   Begin VB.Image Image8 
      Height          =   210
      Left            =   13800
      Picture         =   "VENTAS.frx":1191B
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2625
   End
   Begin VB.Image Image7 
      Height          =   1020
      Left            =   14160
      Picture         =   "VENTAS.frx":12AA8
      Top             =   6000
      Width           =   2040
   End
   Begin VB.Image Image6 
      Height          =   1020
      Left            =   14160
      Picture         =   "VENTAS.frx":13DEA
      Top             =   7680
      Width           =   2040
   End
   Begin VB.Label LBLP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   17
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   3270
      Left            =   13800
      Picture         =   "VENTAS.frx":15046
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   225
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   13920
      Picture         =   "VENTAS.frx":15283
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2505
   End
   Begin VB.Image Image3 
      Height          =   4455
      Left            =   13800
      Picture         =   "VENTAS.frx":16410
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   0
      Picture         =   "VENTAS.frx":1664D
      Top             =   0
      Width           =   3825
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6000
      TabIndex        =   9
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID DEL PRODUCTO:"
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
      Height          =   735
      Left            =   6240
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO/u"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMACIÓN SOBRE LOS PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   5055
   End
End
Attribute VB_Name = "FRMVENTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim CN As New ADODB.Connection
 Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        Check1.Value = vbUnchecked
        Exit Sub
    End If
    cmddeli.ForeColor = RGB(255, 255, 255)
    
End Sub
 
Private Sub CMDB1_Click()
TXTCAN2.Enabled = True

rs.Requery
rs.Find "idproducto=" & Val(TXTIDPRO.Text)

If rs.EOF Then
            MsgBox "No se encontro ningun registro", vbInformation, "Eliminar registro"
            Exit Sub 'Termina el procedimiento
        Else
           TXTCAN.Text = rs.Fields("CANTIDAD")
           TXTCOS.Text = rs.Fields("PRECIO")
           TXTID.Text = Val(TXTIDPRO.Text)
           LBLP.Caption = rs.Fields("nombre")
            Command1.Enabled = True
            CMDAGR.Enabled = True
            
End If
End Sub

Private Sub CMDAGR_Click()

If Val(TXTCAN2.Text) > Val(TXTCAN.Text) Then
    MsgBox "El cantidad de productos pedidos es mayor al stock en este momento.", vbInformation, "Dialogo"
    Exit Sub
Else
    rs.Fields("CANTIDAD") = Val(TXTCAN.Text) - Val(TXTCAN2.Text)
    rs.Update
End If



TXTTOT.Text = Val(TXTCAN2.Text) * Val(TXTCOS.Text)
ADODCFACTURAS.Recordset.AddNew

'Le añadi estas lineas, como te dije al momento de poner un AddNew debo especificar los campos y con que informacion voy _
a llenarlos. El problema de porque no nos salio antes es que en tu proyecto tienes como datasource de los textbox el adodc _
por lo que al momento de poner AddNew los textbox borran su contenido y ya no podemos extraer la infomacion de ahi, solo le quite eso _
y le añadi estas lineas.


ADODCFACTURAS.Recordset("IDPRODUCTO") = (TXTIDPRO.Text)
ADODCFACTURAS.Recordset("CANTIDAD") = (TXTCAN2.Text)
ADODCFACTURAS.Recordset("PRECIO") = (TXTTOT.Text)
ADODCFACTURAS.Recordset("IDVENTAS") = (LBLIDVENTA.Caption)


'LBLIDVENTA.Caption
ADODCFACTURAS.Recordset.Update
ADODCFACTURAS.Refresh
ADODCFACTURAS.Recordset.MoveLast

RSVEN.Update

'FRMINV.ADODCINV.Refresh
TXTCAN.Text = Val(TXTCAN.Text) - Val(TXTCAN2.Text)

LBLTOTGLOBAL.Caption = Val(LBLTOTGLOBAL.Caption) + Val(TXTTOT.Text)

CMDELI.Enabled = True
Command2.Enabled = True
buscarfactura
End Sub

Private Sub CMDCREARVEN_Click()
    If TXTCEDULACLIENTE.Text = "" Then
        Exit Sub
    Else
    

    CMDB1.Enabled = True
    TXTIDPRO.Enabled = True
    TXTIDPRO.SetFocus
    LBLFECHA.Caption = (Date)
    LBLHORA = (Time)
    
    RSVEN.AddNew
    RSVEN("CEDULADUENO") = (a)
    RSVEN("CEDULACLIENTE") = (TXTCEDULACLIENTE.Text)
    RSVEN("FECHA") = (LBLFECHA.Caption)
    RSVEN("HORA") = (LBLHORA.Caption)
    RSVEN.Update
    
    
    LBLIDVENTA.Caption = RSVEN.Fields("IDVENTAS")
    buscarfactura
    CMDCREARVEN.Enabled = False
    End If
End Sub
Sub buscarfactura()
    With ADODCFACTURAS.Recordset
        If .State = 1 Then .Close
        .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
        .Open "select * from FACTURA where idventas = '" & LBLIDVENTA.Caption & "'", CN
        If .EOF And .BOF Then
        MsgBox ("Venta Creada"), vbInformation, "Dialogo"
        Else
        
        .MoveLast
        End If
        'DataGrid1.Refresh
        Set DataGrid1.DataSource = ADODCFACTURAS.Recordset
        
    End With
End Sub

Private Sub CMDELI_Click()
If MsgBox("Esta seguro que desea eliminar un registro?", vbQuestion + vbYesNo, "Ventas") = vbYes Then
        rs.Fields("CANTIDAD") = Val(TXTCAN.Text) + Val(TXTCAN2.Text)
        rs.Update
        ADODCFACTURAS.Recordset.Delete
        ADODCFACTURAS.Refresh
        LBLTOTGLOBAL.Caption = Val(LBLTOTGLOBAL.Caption) - Val(TXTTOT.Text)
        buscarfactura
    End If
End Sub

Private Sub CMDREG_Click()
If MsgBox("Esta seguro que desea regresar al formulario de inventario?", vbQuestion + vbYesNo, "Ventas") = vbYes Then
    FRMINV.Show
    rs.Close
    
    Unload Me
    
    End If
End Sub

Private Sub Command1_Click()
TXTIDPRO.Text = ""
TXTID.Text = ""
TXTCAN.Text = ""
TXTCOS.Text = ""
TXTCAN2.Text = ""
TXTTOT.Text = ""
LBLP.Caption = "Nombre del Producto"
'LBLIDVENTA.Caption = ""
'LBLFECHA.Caption = ""
'LBLHORA.Caption = ""
End Sub

Private Sub Command2_Click()

'Set DRFACTURA.DataSource = ADODCFACTURAS.Recordset

'Set rs = CN.Execute("select *from ventas")
    'If rs.EOF = False Then
    Dim IVA, TOTAL As Double
    IVA = CDbl(LBLTOTGLOBAL.Caption) * 0.12
    TOTAL = CDbl(LBLTOTGLOBAL.Caption) + IVA
    
    Set DataReport1.DataSource = ADODCFACTURAS.Recordset
    DataReport1.Sections("Section3").Controls("LBLSUBTOTAL").Caption = (LBLTOTGLOBAL.Caption)
    DataReport1.Sections("Section2").Controls("LBLCEDULACLIENTE").Caption = (TXTCEDULACLIENTE.Text)
    DataReport1.Sections("Section2").Controls("LBLCEDULADUE").Caption = (a)
    DataReport1.Sections("Section3").Controls("LBLIVA").Caption = (IVA)
    DataReport1.Sections("Section3").Controls("TXTTOTALFINAL").Caption = (TOTAL)
    DataReport1.Show
    CMDCREARVEN.Enabled = True
'End If

End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
    'Abrimos la base de datos "agenda.mdb".
    If CN.State = 0 Then CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\BASEINV.mdb;Persist Security Info=False"
    
    '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\JULIO\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
    rs.Source = "INVENTARIO" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open "select * from INVENTARIO", CN 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    'Cargamos los datos en las cajas de texto.
    rs.MoveFirst 'Nos movemos al principio del Recordset.
    TXTCAN.Enabled = False
    TXTCOS.Enabled = False
    TXTCAN2.Enabled = False
    TXTTOT.Enabled = False
    TXTIDPRO.Enabled = False
     Command1.Enabled = False
      CMDAGR.Enabled = False
    CMDELI.Enabled = False
    CMDB1.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    
    tablaVENTAS
End Sub


Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\img0\reg1.jpg")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\img0\reg0.jpg")
If MsgBox("Esta seguro que desea regresar al formulario de inventario?", vbQuestion + vbYesNo, "Ventas") = vbYes Then
    FRMINV.Show
    rs.Close
    
    Unload Me
    
    End If
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = LoadPicture(App.Path & "\img0\cer1.jpg")
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = LoadPicture(App.Path & "\img0\cer0.jpg")
FRMINV.Hide
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Picture = LoadPicture(App.Path & "\img0\mos1.jpg")
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Picture = LoadPicture(App.Path & "\img0\mos0.jpg")
FRMINV.Show
End Sub

Private Sub ImageM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu1.jpg")
End Sub

Private Sub ImageM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu0.jpg")
If MsgBox("Esta seguro que desea regresar al menu?", vbQuestion + vbYesNo, "Ventas") = vbYes Then
    FRMMENU.Show
    rs.Close
    
    Unload Me
    
    End If
End Sub

