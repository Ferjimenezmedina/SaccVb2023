VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form xMENU 
   Caption         =   "SISTEMA AP TONER"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11265
   ClipControls    =   0   'False
   Icon            =   "MENU2.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   8040
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   63
      Left            =   2280
      TabIndex        =   78
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   62
      Left            =   2160
      TabIndex        =   77
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   61
      Left            =   2040
      TabIndex        =   76
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   60
      Left            =   1920
      TabIndex        =   75
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   59
      Left            =   1800
      TabIndex        =   74
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   58
      Left            =   1680
      TabIndex        =   73
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   57
      Left            =   1560
      TabIndex        =   72
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      Picture         =   "MENU2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   1440
      TabIndex        =   69
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   55
      Left            =   2520
      TabIndex        =   68
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   54
      Left            =   2400
      TabIndex        =   67
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   2280
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   52
      Left            =   2160
      TabIndex        =   65
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   51
      Left            =   2040
      TabIndex        =   64
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   1920
      TabIndex        =   63
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   1800
      TabIndex        =   62
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   1680
      TabIndex        =   61
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   1560
      TabIndex        =   60
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   1440
      TabIndex        =   59
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   2520
      TabIndex        =   58
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   2400
      TabIndex        =   57
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   2280
      TabIndex        =   56
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   2160
      TabIndex        =   55
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   2040
      TabIndex        =   54
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   1920
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   1800
      TabIndex        =   52
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   1680
      TabIndex        =   51
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   1560
      TabIndex        =   50
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   1440
      TabIndex        =   49
      Top             =   5640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   2520
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   2400
      TabIndex        =   47
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   2280
      TabIndex        =   46
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   2160
      TabIndex        =   45
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   44
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   1920
      TabIndex        =   43
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   1800
      TabIndex        =   42
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   1680
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   1560
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   1440
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   2520
      TabIndex        =   38
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   2400
      TabIndex        =   37
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   2280
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   2160
      TabIndex        =   35
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   34
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   1920
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   1800
      TabIndex        =   32
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   1680
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   1560
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   1440
      TabIndex        =   29
      Top             =   5160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   2520
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   2400
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   2280
      TabIndex        =   26
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   2160
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1920
      Top             =   0
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SUCURSALES"
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   7710
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "27/09/2006"
            Object.ToolTipText     =   "FECHA"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "03:48 p.m."
            Object.ToolTipText     =   "HORA"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4940
            MinWidth        =   4940
            Text            =   "SISTEMA AP TONER 1.1"
            TextSave        =   "SISTEMA AP TONER 1.1"
            Object.ToolTipText     =   "Sistema Ap Toner 1.0"
         EndProperty
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
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H80000013&
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del usuario"
      Height          =   2055
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   7695
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NOMBRE"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   0
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USUARIOS"
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
   Begin VB.PictureBox Picture1 
      Height          =   7695
      Left            =   0
      Picture         =   "MENU2.frx":2234
      ScaleHeight     =   7635
      ScaleWidth      =   11235
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
   End
   Begin VB.Menu Ventillas 
      Caption         =   "Ventas"
      Begin VB.Menu Comandas 
         Caption         =   "Comandas"
      End
      Begin VB.Menu Venta 
         Caption         =   "Venta al Cotado"
      End
      Begin VB.Menu Facturar 
         Caption         =   "Factura"
      End
      Begin VB.Menu NotaCred 
         Caption         =   "Nota Credito"
      End
      Begin VB.Menu Garantias 
         Caption         =   "Garantias"
      End
      Begin VB.Menu Cotizacion2 
         Caption         =   "Cotizacion"
         Begin VB.Menu CotiRapi 
            Caption         =   "Rapida"
         End
      End
      Begin VB.Menu AsTec 
         Caption         =   "Asistencia Tecnica"
      End
      Begin VB.Menu Corte 
         Caption         =   "Corte de Caja"
      End
      Begin VB.Menu OrCOmp 
         Caption         =   "Ventas Programadas"
         Begin VB.Menu Captu 
            Caption         =   "Capturar"
         End
         Begin VB.Menu CerrVent 
            Caption         =   "Cerrar como Venta"
         End
      End
      Begin VB.Menu Vencred 
         Caption         =   "Venta Credito"
         Begin VB.Menu AbonVenCredito 
            Caption         =   "Abonar"
         End
         Begin VB.Menu HisVenCred 
            Caption         =   "Historial"
         End
      End
      Begin VB.Menu VenEsp 
         Caption         =   "Venta Especial"
      End
      Begin VB.Menu Domis 
         Caption         =   "Domicilios"
         Begin VB.Menu AgreDomi 
            Caption         =   "Agregar"
         End
         Begin VB.Menu RepDomi 
            Caption         =   "Reporte"
         End
      End
   End
   Begin VB.Menu Busca 
      Caption         =   "Consultar"
      Begin VB.Menu DatProv 
         Caption         =   "Datos de Proveedor"
      End
      Begin VB.Menu VerFaltantes 
         Caption         =   "Ver Faltantes"
      End
      Begin VB.Menu RasPed 
         Caption         =   "Rastrear Pedido"
      End
      Begin VB.Menu ProdPed 
         Caption         =   "Producto en pedido"
      End
      Begin VB.Menu BusExi 
         Caption         =   "Existencia"
      End
      Begin VB.Menu verExi 
         Caption         =   "Ver Existencia en Bodega"
      End
      Begin VB.Menu BusPro 
         Caption         =   "Producto"
      End
   End
   Begin VB.Menu Pedido 
      Caption         =   "Pedidos"
      Begin VB.Menu PedidosSuc 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu Requisicion 
         Caption         =   "Requisicion"
      End
   End
   Begin VB.Menu Almacen 
      Caption         =   "Almacen"
      Begin VB.Menu Entradas 
         Caption         =   "Entradas"
         Begin VB.Menu EntAlm1 
            Caption         =   "Entrada Almacen 1"
         End
         Begin VB.Menu EntAl2 
            Caption         =   "Entrada Almacen 2"
         End
         Begin VB.Menu EntAl3 
            Caption         =   "Entrada Almacen 3"
         End
      End
      Begin VB.Menu TrasInvent 
         Caption         =   "Traspasos de Inventario"
      End
      Begin VB.Menu BusEntr 
         Caption         =   "Buscar Entrada"
      End
      Begin VB.Menu VerPerSuc 
         Caption         =   "Ver Pedidos de Sucursales"
      End
      Begin VB.Menu VnProg 
         Caption         =   "Ventas Programadas"
         Begin VB.Menu SurOrd 
            Caption         =   "Surtir Venta Programada"
         End
         Begin VB.Menu AjuPedMan 
            Caption         =   "Ajustar Venta Programada"
         End
      End
      Begin VB.Menu Inventarios 
         Caption         =   "Inventarios"
      End
      Begin VB.Menu SalInt 
         Caption         =   "Salida Interna"
      End
      Begin VB.Menu SurtirPedidos 
         Caption         =   "Surtir Pedidos"
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Begin VB.Menu OrCompr 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu Cotizar 
         Caption         =   "Cotizar"
      End
      Begin VB.Menu Revisar 
         Caption         =   "Revisar"
      End
      Begin VB.Menu Atorizar_Cotizacines 
         Caption         =   "Autorizar Cotizaciones"
      End
      Begin VB.Menu GuardarOrden 
         Caption         =   "Guardar Ordenes"
      End
      Begin VB.Menu ImpOrdCom 
         Caption         =   "Imprimir Orden de Compra"
      End
   End
   Begin VB.Menu AsTecMen 
      Caption         =   "Asistencia Tecnica"
      Begin VB.Menu VerAsTec 
         Caption         =   "Ver"
      End
   End
   Begin VB.Menu Recursos 
      Caption         =   "Administrador"
      Begin VB.Menu RegProd 
         Caption         =   "Registrar Producto"
         Begin VB.Menu Alm2 
            Caption         =   "Almacen 1 y 2"
         End
         Begin VB.Menu Alm3 
            Caption         =   "Almacen 3"
         End
      End
      Begin VB.Menu Agregar 
         Caption         =   "Agregar"
         Begin VB.Menu Agente2 
            Caption         =   "Agente"
         End
         Begin VB.Menu Cliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu Sucursal 
            Caption         =   "Sucursal"
         End
         Begin VB.Menu ProveedorMen 
            Caption         =   "Proveedor"
         End
         Begin VB.Menu AgMen 
            Caption         =   "Mensajero"
         End
      End
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
         Begin VB.Menu ElAgente 
            Caption         =   "Agente"
         End
         Begin VB.Menu ElCliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu ElProv 
            Caption         =   "Proveedor"
         End
      End
      Begin VB.Menu Modificar 
         Caption         =   "Modificar"
         Begin VB.Menu ModPrecio 
            Caption         =   "Precio"
         End
      End
      Begin VB.Menu Excel 
         Caption         =   "Pasar a Excel"
      End
      Begin VB.Menu DPVent 
         Caption         =   "Dar Permiso Especiales"
      End
   End
   Begin VB.Menu Produccion 
      Caption         =   "Produccion"
      Begin VB.Menu JuegRep 
         Caption         =   "Nuevo Juego de Reparación"
      End
      Begin VB.Menu Ver 
         Caption         =   "Ver pedidos"
      End
      Begin VB.Menu VerCl 
         Caption         =   "Ver pedidos de Clientes"
      End
      Begin VB.Menu VerSuc 
         Caption         =   "Ver pedidos de Sucursales"
      End
      Begin VB.Menu VJR 
         Caption         =   "Ver Juegos de Reparación"
      End
   End
   Begin VB.Menu Utilerias 
      Caption         =   "Utilerias"
      Begin VB.Menu Promociones 
         Caption         =   "Promociones"
      End
      Begin VB.Menu AgLicita 
         Caption         =   "Licitación"
      End
      Begin VB.Menu Marcas 
         Caption         =   "Marcas"
      End
      Begin VB.Menu dolar2 
         Caption         =   "Dolar"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "xMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Y As Integer
Private Sub AbonVenCredito_Click()
    FrmAbonoCuenta.Show vbModal
End Sub
Private Sub Agente2_Click()
    frmPermisos.Show vbModal
End Sub
Private Sub AgLicita_Click()
    FrmLicitacion.Show vbModal
End Sub
Private Sub AgMen_Click()
    FrmNueRep.Show vbModal
End Sub
Private Sub AgreDomi_Click()
    FrmRegDomi.Show vbModal
End Sub
Private Sub AjuPedMan_Click()
    PermisoAjuste.Show vbModal
End Sub
Private Sub Alm2_Click()
    frmAlmacen2.Show vbModal
End Sub
Private Sub Alm3_Click()
    frmAlmacen3.Show vbModal
End Sub
Private Sub ASTec_Click()
    AsisTec.Show vbModal
End Sub
Private Sub Atorizar_Cotizacines_Click()
    frmAutorizarCotizaciones.Show vbModal
End Sub
Private Sub BusEntr_Click()
    BuscaEntrada.Show vbModal
End Sub
Private Sub BusExi_Click()
    BuscaExist.Show vbModal
End Sub
Private Sub BusPro_Click()
    BuscaProd.Show vbModal
End Sub
Private Sub Captu_Click()
    frmClien.Show vbModal
End Sub
Private Sub CerrVent_Click()
    FrmPasPedVent.Show vbModal
End Sub
Private Sub Cliente_Click()
    AltaClien.Show vbModal
End Sub
Private Sub Comandas_Click()
    frmComandas.Show vbModal
End Sub
Private Sub Command1_Click()
    Checar
End Sub
Private Sub Command2_Click()
    MsgAPToner.Show
End Sub
Private Sub Corte_Click()
    CorteCaja.Show vbModal
End Sub
Private Sub CotiRapi_Click()
    FrmCotizaRapida.Show vbModal
End Sub
Private Sub Cotizar_Click()
    frmRequisiciones.Show vbModal
End Sub
Private Sub DatProv_Click()
    frmProveedores.Show vbModal
End Sub
Private Sub dolar2_Click()
    Dolar.Show vbModal
End Sub
Private Sub DPVent_Click()
    DarPerVenta.Show vbModal
End Sub
Private Sub ElAgente_Click()
    EliAgente.Show vbModal
End Sub
Private Sub ElCliente_Click()
    EliCliente.Show vbModal
End Sub
Private Sub ElProv_Click()
    EliProveedor.Show vbModal
End Sub
Private Sub EntAl2_Click()
    EntradaProd2.Show vbModal
End Sub
Private Sub EntAl3_Click()
    EntradaProd3.Show vbModal
End Sub
Private Sub EntAlm1_Click()
    EntradaProd.Show vbModal
End Sub
Private Sub Excel_Click()
    BajaExcel.Show vbModal
End Sub
Private Sub Exit_Click()
    Me.Adodc1.Recordset.Close
    Me.Adodc2.Recordset.Close
    Unload Me
End Sub
Private Sub Facturar_Click()
    frmFactura.Show vbModal
End Sub
Private Sub Form_Load()
    Command2.Visible = False
    Comandas.Enabled = False
    Venta.Enabled = False
    Facturar.Enabled = False
    NotaCred.Enabled = False
    Garantias.Enabled = False
    AsTec.Enabled = False
    Corte.Enabled = False
    BusPro.Enabled = False
    BusExi.Enabled = False
    PedidosSuc.Enabled = False
    Requisicion.Enabled = False
    VerFaltantes.Enabled = False
    EntAlm1.Enabled = False
    EntAl2.Enabled = False
    EntAl3.Enabled = False
    Alm2.Enabled = False
    Alm3.Enabled = False
    TrasInvent.Enabled = False
    BusEntr.Enabled = False
    VerPerSuc.Enabled = False
    Inventarios.Enabled = False
    SalInt.Enabled = False
    OrCompr.Enabled = False
    DatProv.Enabled = False
    VerAsTec.Enabled = False
    Agente2.Enabled = False
    AgMen.Enabled = False
    Cliente.Enabled = False
    Sucursal.Enabled = False
    ProveedorMen.Enabled = False
    ElAgente.Enabled = False
    ElCliente.Enabled = False
    ElProv.Enabled = False
    ModPrecio.Enabled = False
    Excel.Enabled = False
    JuegRep.Enabled = False
    Ver.Enabled = False
    VerCl.Enabled = False
    VerSuc.Enabled = False
    VJR.Enabled = False
    Promociones.Enabled = False
    AgLicita.Enabled = False
    Marcas.Enabled = False
    dolar2.Enabled = False
    DPVent.Enabled = False
    VenEsp.Enabled = False
    Captu.Enabled = False
    CerrVent.Enabled = False
    RasPed.Enabled = False
    ProdPed.Enabled = False
    SurOrd.Enabled = False
    verExi.Enabled = False
    AjuPedMan.Enabled = False
    AgreDomi.Enabled = False
    RepDomi.Enabled = False
    AbonVenCredito.Enabled = False
    HisVenCred.Enabled = False
    SurtirPedidos.Enabled = False
    Cotizar.Enabled = False
    CotiRapi.Enabled = False
    Atorizar_Cotizacines.Enabled = False
    GuardarOrden.Enabled = False
    ImpOrdCom.Enabled = False
    Revisar.Enabled = False
    Dim i As Long
    For i = 0 To 63
        Set Text1(i).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_USUARIO"
    Text1(1).DataField = "NOMBRE"
    Text1(2).DataField = "APELLIDOS"
    Text1(3).DataField = "PUESTO"
    Text1(4).DataField = "PASSWORD"
    Text1(5).DataField = "ID_SUCURSAL"
    Text1(6).DataField = "Pe1"
    Text1(7).DataField = "Pe2"
    Text1(8).DataField = "Pe3"
    Text1(9).DataField = "Pe4"
    Text1(10).DataField = "Pe5"
    Text1(11).DataField = "Pe6"
    Text1(12).DataField = "Pe7"
    Text1(13).DataField = "Pe8"
    Text1(14).DataField = "Pe9"
    Text1(15).DataField = "Pe10"
    Text1(16).DataField = "Pe11"
    Text1(17).DataField = "Pe12"
    Text1(18).DataField = "Pe13"
    Text1(19).DataField = "Pe14"
    Text1(20).DataField = "Pe15"
    Text1(21).DataField = "Pe16"
    Text1(22).DataField = "Pe17"
    Text1(23).DataField = "Pe18"
    Text1(24).DataField = "Pe19"
    Text1(25).DataField = "Pe20"
    Text1(26).DataField = "Pe21"
    Text1(27).DataField = "Pe22"
    Text1(28).DataField = "Pe23"
    Text1(29).DataField = "Pe24"
    Text1(30).DataField = "Pe25"
    Text1(31).DataField = "Pe26"
    Text1(32).DataField = "Pe27"
    Text1(33).DataField = "Pe28"
    Text1(34).DataField = "Pe29"
    Text1(35).DataField = "Pe30"
    Text1(36).DataField = "Pe31"
    Text1(37).DataField = "Pe32"
    Text1(38).DataField = "Pe33"
    Text1(39).DataField = "Pe34"
    Text1(40).DataField = "Pe35"
    Text1(41).DataField = "Pe36"
    Text1(42).DataField = "Pe37"
    Text1(43).DataField = "Pe38"
    Text1(44).DataField = "Pe39"
    Text1(45).DataField = "Pe40"
    Text1(46).DataField = "Pe41"
    Text1(47).DataField = "Pe42"
    Text1(48).DataField = "Pe43"
    Text1(49).DataField = "Pe44"
    Text1(50).DataField = "Pe45"
    Text1(51).DataField = "Pe46"
    Text1(52).DataField = "Pe47"
    Text1(53).DataField = "Pe48"
    Text1(54).DataField = "Pe49"
    Text1(55).DataField = "Pe50"
    Text1(56).DataField = "Pe51"
    Text1(57).DataField = "Pe52"
    Text1(58).DataField = "Pe53"
    Text1(59).DataField = "Pe54"
    Text1(60).DataField = "Pe55"
    Text1(61).DataField = "Pe56"
    Text1(62).DataField = "Pe57"
    Text1(63).DataField = "Pe58"
    Dim X As Long
    For X = 0 To 5
        Set Text4(X).DataSource = Adodc2
    Next
    Text4(0).DataField = "NOMBRE"
    Text4(1).DataField = "CALLE"
    Text4(2).DataField = "COLONIA"
    Text4(3).DataField = "CIUDAD"
    Text4(4).DataField = "ESTADO"
    Text4(5).DataField = "TELEFONO"
End Sub

Private Sub Form_Resize()
    
    'Me.Picture1.Move (Me.Width - Me.Picture1.Width) / 2, (Me.Height - Me.Picture1.Height) / 2
    'Me.Frame1.Move (Me.Width - Me.Picture1.Width) / 2, (Me.Height - Me.Picture1.Height) / 2

End Sub

Private Sub Garantias_Click()
    frmGarantias.Show vbModal
End Sub
Private Sub GuardarOrden_Click()
    frmOrden_Compra.Show vbModal
End Sub
Private Sub HisVenCred_Click()
    FrmBusHisCred.Show vbModal
End Sub
Private Sub ImpOrdCom_Click()
    frmOrdenCompra.Show vbModal
End Sub
Private Sub Inventarios_Click()
    frmSucInv.Show vbModal
End Sub
Private Sub JuegRep_Click()
    JuegoRep.Show vbModal
End Sub
Private Sub Marcas_Click()
    Marca.Show vbModal
End Sub
Private Sub ModPrecio_Click()
    CambioPRe.Show vbModal
End Sub
Private Sub NotaCred_Click()
    NotaCredito.Show vbModal
End Sub
Private Sub OrCompr_Click()
    Orden.Show vbModal
End Sub
Private Sub PedidosSuc_Click()
    Pedidos.Show vbModal
End Sub
Private Sub ProdPed_Click()
    FrmBusProdPed.Show vbModal
End Sub
Private Sub Promociones_Click()
    frmPromos.Show vbModal
End Sub
Private Sub ProveedorMen_Click()
    Proveedor.Show vbModal
End Sub
Private Sub RasPed_Click()
    FrmRastrearPed.Show vbModal
End Sub
Private Sub RepDomi_Click()
    FrmRevDomi.Show vbModal
End Sub
Private Sub Requisicion_Click()
    frmRequisicion.Show vbModal
End Sub
Private Sub Revisar_Click()
    frmVerCotizaciones.Show vbModal
End Sub
Private Sub SalInt_Click()
    Salidas.Show vbModal
End Sub
Private Sub Sucursal_Click()
    AltaSucu.Show vbModal
End Sub
Private Sub SurOrd_Click()
    frmShowPediC.Show vbModal
End Sub
Private Sub SurtirPedidos_Click()
    frmSurtir.Show vbModal
End Sub
Private Sub Text1_Change(Index As Integer)
    If Index = 5 Then
        Dim nReg2 As String
        Dim sADOBuscar2 As String
        On Error Resume Next
        nReg2 = Val(Text1(5).Text)
        sADOBuscar2 = "ID_SUCURSAL = " & nReg2
        Adodc2.Recordset.MoveFirst
        Adodc2.Recordset.Find sADOBuscar2
    End If
End Sub
Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" And Text3.Text <> "" Then
        Checar
    End If
End Sub
Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" And Text3.Text <> "" Then
        Checar
    End If
End Sub
Private Sub TrasInvent_Click()
    Transfe.Show vbModal
End Sub
Private Sub VenEsp_Click()
    PermisoVenta.Show vbModal
End Sub
Private Sub Venta_Click()
    Ventas.Show vbModal
End Sub
Private Sub Ver_Click()
    frmProd.Show vbModal
End Sub
Private Sub VerAsTec_Click()
    frmAStec.Show vbModal
End Sub
Private Sub VerCl_Click()
    frmRevCom.Show vbModal
End Sub
Private Sub verExi_Click()
    FrmVerExisBodega.Show vbModal
End Sub
Private Sub VerFaltantes_Click()
    Faltantes.Show vbModal
End Sub
Private Sub VerPerSuc_Click()
    frmRevPed.Show vbModal
End Sub
Private Sub VerSuc_Click()
    frmRevComSuc.Show vbModal
End Sub
Private Sub VJR_Click()
    VerJuegoRep.Show vbModal
End Sub
Private Sub Checar()
    If Text2.Text <> "" And Text3.Text <> "" Then
        Dim nReg As String
        Dim vBookmark As Variant
        Dim sADOBuscar As String
        On Error Resume Next
        nReg = Text2.Text
        sADOBuscar = "NOMBRE LIKE '" & nReg & "'"
        vBookmark = Adodc1.Recordset.Bookmark
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Find sADOBuscar
        Text2.Text = Replace(Text2.Text, " ", "")
        Text1(1).Text = Replace(Text1(1).Text, " ", "")
        Text1(4).Text = Replace(Text1(4).Text, " ", "")
        Text3.Text = Replace(Text3.Text, " ", "")
        If Text2.Text = Text1(1).Text And Text3.Text = Text1(4).Text Then
            Me.Picture1.Visible = True
            Me.Frame1.Visible = False
            Me.Command1.Visible = False
            Me.Ventillas.Enabled = True
            Me.Utilerias.Enabled = True
            Me.Produccion.Enabled = True
            Me.Recursos.Enabled = True
            Me.Almacen.Enabled = True
            Me.Compras.Enabled = True
            Me.Pedido.Enabled = True
            Me.AsTecMen.Enabled = True
            Command2.Visible = True
            If Text1(6).Text = "S" Then
                Comandas.Enabled = True
            End If
            If Text1(7).Text = "S" Then
                Venta.Enabled = True
            End If
            If Text1(8).Text = "S" Then
                Facturar.Enabled = True
            End If
            If Text1(9).Text = "S" Then
                NotaCred.Enabled = True
            End If
            If Text1(10).Text = "S" Then
                Garantias.Enabled = True
            End If
            If Text1(11).Text = "S" Then
                 CotiRapi.Enabled = True
            End If
            If Text1(12).Text = "S" Then
                AsTec.Enabled = True
            End If
            If Text1(13).Text = "S" Then
                Corte.Enabled = True
            End If
            If Text1(14).Text = "S" Then
                BusPro.Enabled = True
            End If
            If Text1(15).Text = "S" Then
                BusExi.Enabled = True
            End If
            If Text1(16).Text = "S" Then
                PedidosSuc.Enabled = True
            End If
            If Text1(17).Text = "S" Then
                Requisicion.Enabled = True
                VerFaltantes.Enabled = True
            End If
            If Text1(18).Text = "S" Then
                EntAlm1.Enabled = True
            End If
            If Text1(19).Text = "S" Then
                EntAl2.Enabled = True
            End If
            If Text1(20).Text = "S" Then
                EntAl3.Enabled = True
            End If
            If Text1(21).Text = "S" Then
                Alm2.Enabled = True
            End If
            If Text1(22).Text = "S" Then
                Alm3.Enabled = True
            End If
            If Text1(23).Text = "S" Then
                TrasInvent.Enabled = True
            End If
            If Text1(24).Text = "S" Then
                BusEntr.Enabled = True
            End If
            If Text1(25).Text = "S" Then
                VerPerSuc.Enabled = True
            End If
            If Text1(26).Text = "S" Then
                Inventarios.Enabled = True
            End If
            If Text1(27).Text = "S" Then
                OrCompr.Enabled = True
                DatProv.Enabled = True
            End If
            If Text1(28).Text = "S" Then
                VerAsTec.Enabled = True
            End If
            If Text1(29).Text = "S" Then
                Agente2.Enabled = True
                AgMen.Enabled = True
            End If
            If Text1(30).Text = "S" Then
                Cliente.Enabled = True
            End If
            If Text1(31).Text = "S" Then
                Sucursal.Enabled = True
            End If
            If Text1(32).Text = "S" Then
                ProveedorMen.Enabled = True
            End If
            If Text1(33).Text = "S" Then
                ElAgente.Enabled = True
            End If
            If Text1(34).Text = "S" Then
                ElCliente.Enabled = True
            End If
            If Text1(35).Text = "S" Then
                ElProv.Enabled = True
            End If
            If Text1(36).Text = "S" Then
                ModPrecio.Enabled = True
            End If
            If Text1(37).Text = "S" Then
                Excel.Enabled = True
            End If
            If Text1(38).Text = "S" Then
                JuegRep.Enabled = True
            End If
            If Text1(39).Text = "S" Then
                Ver.Enabled = True
            End If
            If Text1(40).Text = "S" Then
                VerCl.Enabled = True
            End If
            If Text1(41).Text = "S" Then
                VerSuc.Enabled = True
            End If
            If Text1(42).Text = "S" Then
                VJR.Enabled = True
            End If
            If Text1(43).Text = "S" Then
                Promociones.Enabled = True
                AgLicita.Enabled = True
            End If
            If Text1(44).Text = "S" Then
                Marcas.Enabled = True
            End If
            If Text1(45).Text = "S" Then
                dolar2.Enabled = True
            End If
            If Text1(46).Text = "S" Then
                SalInt.Enabled = True
            End If
            If Text1(47).Text = "S" Then
                DPVent.Enabled = True
            End If
            If Text1(48).Text = "S" Then
                VenEsp.Enabled = True
            End If
            If Text1(49).Text = "S" Then
                Captu.Enabled = True
            End If
            If Text1(50).Text = "S" Then
                CerrVent.Enabled = True
            End If
            If Text1(51).Text = "S" Then
                RasPed.Enabled = True
            End If
            If Text1(52).Text = "S" Then
                ProdPed.Enabled = True
            End If
            If Text1(53).Text = "S" Then
                SurOrd.Enabled = True
            End If
            If Text1(54).Text = "S" Then
                verExi.Enabled = True
            End If
            If Text1(55).Text = "S" Then
                AjuPedMan.Enabled = True
            End If
            If Text1(56).Text = "S" Then
                AgreDomi.Enabled = True
                RepDomi.Enabled = True
            End If
            If Text1(57).Text = "S" Then
                AbonVenCredito.Enabled = True
                HisVenCred.Enabled = True
            End If
            If Text1(58).Text = "S" Then
                SurtirPedidos.Enabled = True
            End If
            If Text1(59).Text = "S" Then
                Revisar.Enabled = True
            End If
            If Text1(60).Text = "S" Then
                Atorizar_Cotizacines.Enabled = True
            End If
            If Text1(61).Text = "S" Then
               GuardarOrden.Enabled = True
            End If
            If Text1(62).Text = "S" Then
                ImpOrdCom.Enabled = True
            End If
            If Text1(63).Text = "S" Then
                Cotizar.Enabled = True
                DatProv.Enabled = True
            End If
        Else
            If Adodc1.Recordset.EOF Or Text2.Text <> Text1(1).Text Or Text3.Text <> Text1(4).Text Then
                Err.Clear
                MsgBox "El Nombre o el Password son incorrectos."
                Text2.SetFocus
                Text3.SetFocus
            End If
        End If
    End If
End Sub
