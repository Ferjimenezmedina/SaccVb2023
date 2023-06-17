VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Requi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REQUISICIÓN"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   10755
      TabIndex        =   43
      Top             =   2760
      Width           =   10815
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   7320
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   7440
         TabIndex        =   62
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Calcular"
         Height          =   375
         Left            =   9240
         TabIndex        =   61
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         DataField       =   "TOTAL"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   7440
         TabIndex        =   60
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton btnSAlir 
         BackColor       =   &H80000009&
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2040
         TabIndex        =   50
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   8520
         TabIndex        =   49
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   9480
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   9240
         TabIndex        =   47
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mandar a Orden de compra"
         Height          =   495
         Left            =   8880
         TabIndex        =   46
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   9240
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   0
         TabIndex        =   45
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3625
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Iva"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   480
         Top             =   2760
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
         OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "REQUISICION_PRODUCTO"
         Caption         =   "Adodc2"
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
      Begin VB.Label Label18 
         Caption         =   "Total:"
         Height          =   255
         Left            =   6840
         TabIndex        =   64
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "IVA:"
         Height          =   255
         Left            =   6960
         TabIndex        =   63
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   1080
         TabIndex        =   58
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   2400
         TabIndex        =   57
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         Height          =   195
         Left            =   7560
         TabIndex        =   56
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   8760
         TabIndex        =   55
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   9720
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   4680
      TabIndex        =   34
      Top             =   2160
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      DataField       =   "nombre"
      DataSource      =   "Adodc5"
      Height          =   315
      Left            =   2520
      TabIndex        =   33
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "nombre"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      DataField       =   "nombre"
      DataSource      =   "Adodc4"
      Height          =   315
      Left            =   2520
      TabIndex        =   32
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "nombre"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4680
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select nombre from cliente"
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4680
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select nombre from proveedor"
      Caption         =   "Adodc4"
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
   Begin VB.TextBox Text13 
      Height          =   195
      Left            =   720
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text12 
      Height          =   195
      Left            =   600
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text11 
      Height          =   195
      Left            =   480
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text10 
      Height          =   195
      Left            =   360
      TabIndex        =   28
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text9 
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text8 
      Height          =   195
      Left            =   840
      TabIndex        =   26
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text7 
      Height          =   195
      Left            =   720
      TabIndex        =   25
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4680
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ORDEN_COMPRA"
      Caption         =   "Adodc3"
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
   Begin VB.TextBox Text1 
      DataField       =   "CLIENTE"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   15
      Left            =   8280
      TabIndex        =   14
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   14
      Left            =   5880
      TabIndex        =   13
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   13
      Left            =   3360
      TabIndex        =   12
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   12
      Left            =   1080
      TabIndex        =   11
      Top             =   6960
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49610753
      CurrentDate     =   38684
   End
   Begin VB.TextBox Text1 
      DataField       =   "EJECUTIVO"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "FECHA"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "NUM_ORDEN"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_REQUISICION"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "PROVEEDOR"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4680
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "REQUISICION"
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
   Begin VB.TextBox Text2 
      DataField       =   "ID_REQUISICION"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10440
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text14 
      DataField       =   "ID_PRODUCTO"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10560
      TabIndex        =   37
      Text            =   "Text14"
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text15 
      DataField       =   "DESCRIPCION"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10680
      TabIndex        =   38
      Text            =   "Text15"
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text16 
      DataField       =   "CANTIDAD"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10800
      TabIndex        =   39
      Text            =   "Text16"
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text17 
      DataField       =   "PRECIO"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10440
      TabIndex        =   40
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text18 
      DataField       =   "IVA"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10560
      TabIndex        =   41
      Text            =   "Text18"
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text19 
      DataField       =   "IMPORTE"
      DataSource      =   "Adodc2"
      Height          =   195
      Left            =   10680
      TabIndex        =   42
      Text            =   "Text19"
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Entregó a compras"
      Height          =   195
      Left            =   3720
      TabIndex        =   18
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Recibió"
      Height          =   195
      Left            =   6600
      TabIndex        =   17
      Top             =   6720
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Recibió de compras"
      Height          =   195
      Left            =   8640
      TabIndex        =   16
      Top             =   6720
      Width           =   1410
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Aprobado por:"
      Height          =   195
      Left            =   1560
      TabIndex        =   15
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ejecutivo"
      Height          =   195
      Left            =   6240
      TabIndex        =   10
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No. de Orden"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Requisicion"
      Height          =   195
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Requi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub Agregar_Item()
    Dim X As ListItem
    Set X = ListView1.ListItems.Add(, , Text1(6).Text)
    X.Tag = Text1(6).Text
    X.SubItems(1) = Text1(7).Text
    X.SubItems(2) = Text1(8).Text
    X.SubItems(3) = Text22.Text
    X.SubItems(4) = Text1(10).Text
    X.SubItems(5) = Text1(11).Text
End Sub
Private Sub vaciar_item()
    Text1(6).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = ""
    Text22.Text = ""
    Text1(10).Text = "0"
    Text1(11).Text = ""
End Sub
Private Sub cmdAgregar_Click()
    Dim pre, CANT, tot As String
    pre = Text22.Text
    CANT = Text1(6).Text
    tot = CDbl(CANT) * CDbl(pre)
    Text1(11).Text = tot
    Call Agregar_Item
    Call vaciar_item
End Sub
Private Sub cmdNuevo_Click()
    Adodc1.Recordset.AddNew
End Sub
Private Sub Command1_Click()
With Adodc1.Recordset
    If Not (.BOF) And Not (.EOF) Then
        .MoveFirst
        Do While Not (.BOF) And Not (.EOF)
            Adodc2.Recordset.AddNew
            If Not IsNull(.Fields(ID_PROVEEDOR)) Then
                Text1(0).Text = (.Fields(ID_PROVEEDOR))
            End If
            If Not IsNull(.Fields(ID_REQUISICION)) Then
                Text1(1).Text = (.Fields(ID_REQUISICION))
            End If
            .MoveFirst
            If Not IsNull(.Fields(NUM_ORDEN)) Then
                Text1(2).Text = (.Fields(NUM_ORDEN))
            End If
            .MoveFirst
            If Not IsNull(.Fields(Cliente)) Then
                Text1(4).Text = (.Fields(Cliente))
            End If
            .MoveFirst
            If Not IsNull(.Fields(EJECUTIVO)) Then
                Text1(5).Text = (.Fields(EJECUTIVO))
            End If
            .MoveFirst
            If Not IsNull(.Fields(cantidad)) Then
                Text1(6).Text = (.Fields(cantidad))
            End If
            .MoveFirst
            If Not IsNull(.Fields(Clave)) Then
                Text1(7).Text = (.Fields(Clave))
            End If
            .MoveFirst
            If Not IsNull(.Fields(DESCRIPCION)) Then
                Text1(8).Text = (.Fields(DESCRIPCION))
            End If
            If Not IsNull(.Fields(PRECIO)) Then
                Text1(9).Text = (.Fields(PRECIO))
            End If
            If Not IsNull(.Fields(iva)) Then
                Text1(10).Text = (.Fields(iva))
            End If
            Adodc3.Recordset.Update
          Loop
          .MoveNext
        End If
    End With
End Sub
Private Sub Command2_Click()
    Adodc1.Recordset.Update
    Picture1.Left = 120
    Adodc1.Recordset.MoveLast
End Sub
Private Sub Command3_Click()
    Dim X, y, sum, sum2 As Integer
    X = ListView1.ListItems.Count
    sum = 0
    sum2 = 0
    For y = 1 To X
        Text18.Text = ListView1.ListItems(y).SubItems(4)
        Text19.Text = ListView1.ListItems(y).SubItems(5)
        sum = Text19.Text + sum
        sum2 = Text18.Text + sum2
    Next y
    Text21.Text = sum2
    sum = sum2 + sum
    Text20.Text = sum
End Sub
Private Sub Command4_Click()
    Dim X, y As Integer
    X = ListView1.ListItems.Count
    For y = 1 To X
        Adodc2.Recordset.AddNew
        Text2.Text = Text1(1).Text
        Text14.Text = ListView1.ListItems.Item(y)
        Text15.Text = ListView1.ListItems(y).SubItems(1)
        Text16.Text = ListView1.ListItems(y).SubItems(2)
        Text17.Text = ListView1.ListItems(y).SubItems(3)
        Text18.Text = ListView1.ListItems(y).SubItems(4)
        Text19.Text = ListView1.ListItems(y).SubItems(5)
        Text20.Text = ListView1.ListItems(y).SubItems(5)
        Adodc2.Recordset.Update
    Next y
End Sub
Private Sub Command5_Click()
    Adodc2.Recordset.AddNew
End Sub
Private Sub DataCombo1_Change()
    Text1(0).Text = DataCombo1.Text
End Sub
Private Sub DataCombo2_Change()
    Text1(4).Text = DataCombo2.Text
End Sub
Private Sub DTPicker1_Change()
Text1(3).Text = DTPicker1.Value
End Sub
Private Sub Form_Load()
    Dim i As Long
    For i = 0 To 5
        Set Text1(i).DataSource = Adodc1
    Next
    Text1(0).DataField = "PROVEEDOR"
    Text1(1).DataField = "ID_REQUISICION"
    Text1(2).DataField = "NUM_ORDEN"
    Text1(3).DataField = "FECHA"
    Text1(4).DataField = "CLIENTE"
    Text1(5).DataField = "EJECUTIVO"
    DTPicker1.Value = Date
    Text1(3).Text = DTPicker1.Value
End Sub
