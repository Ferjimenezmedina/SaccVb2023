VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ventas"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   169
      Top             =   4080
      Width           =   975
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   170
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepVentas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentas.frx":030A
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   2
      Top             =   5280
      Width           =   975
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepVentas.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentas.frx":0BA3
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   0
      Top             =   6480
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepVentas.frx":26E5
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentas.frx":29EF
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "General de Ventas"
      TabPicture(0)   =   "FrmRepVentas.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSTab3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame29"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Detalles de Ventas"
      TabPicture(1)   =   "FrmRepVentas.frx":4AED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(1)=   "ListView4"
      Tab(1).Control(2)=   "ListView3"
      Tab(1).Control(3)=   "Frame16"
      Tab(1).Control(4)=   "Text20"
      Tab(1).Control(5)=   "Frame15"
      Tab(1).Control(6)=   "Check11"
      Tab(1).Control(7)=   "Frame14"
      Tab(1).Control(8)=   "Check7"
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(10)=   "Frame13"
      Tab(1).Control(11)=   "Check14"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Mas Vendidos"
      TabPicture(2)   =   "FrmRepVentas.frx":4B09
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Text12"
      Tab(2).Control(2)=   "Command4"
      Tab(2).Control(3)=   "Command5"
      Tab(2).Control(4)=   "Frame17"
      Tab(2).Control(5)=   "ListView7"
      Tab(2).Control(6)=   "ListView6"
      Tab(2).Control(7)=   "Cliente"
      Tab(2).Control(8)=   "Line1"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Reporte por Agente"
      TabPicture(3)   =   "FrmRepVentas.frx":4B25
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text18"
      Tab(3).Control(1)=   "Command9"
      Tab(3).Control(2)=   "Frame27"
      Tab(3).Control(3)=   "Frame26"
      Tab(3).Control(4)=   "Frame23"
      Tab(3).Control(5)=   "Command3"
      Tab(3).Control(6)=   "Frame22"
      Tab(3).Control(7)=   "Frame21"
      Tab(3).Control(8)=   "Option11"
      Tab(3).Control(9)=   "Option10"
      Tab(3).Control(10)=   "ListView9"
      Tab(3).Control(11)=   "Frame19"
      Tab(3).Control(12)=   "Label26"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Sucursal"
      TabPicture(4)   =   "FrmRepVentas.frx":4B41
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Check19"
      Tab(4).Control(1)=   "Check18"
      Tab(4).Control(2)=   "Frame25"
      Tab(4).Control(3)=   "Frame24"
      Tab(4).Control(4)=   "ListView10"
      Tab(4).Control(5)=   "Frame20"
      Tab(4).Control(6)=   "Frame4"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Totales"
      TabPicture(5)   =   "FrmRepVentas.frx":4B5D
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command10"
      Tab(5).Control(1)=   "Frame34"
      Tab(5).Control(2)=   "Frame33"
      Tab(5).Control(3)=   "Frame32"
      Tab(5).Control(4)=   "Frame31"
      Tab(5).Control(5)=   "Frame30"
      Tab(5).Control(6)=   "ListView11"
      Tab(5).ControlCount=   7
      Begin VB.Frame Frame29 
         Caption         =   "Tipo de Venta"
         Height          =   1215
         Left            =   3360
         TabIndex        =   171
         Top             =   2160
         Width           =   1335
         Begin VB.OptionButton Option25 
            Caption         =   "Todo"
            Height          =   195
            Left            =   120
            TabIndex        =   174
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option24 
            Caption         =   "Contado"
            Height          =   195
            Left            =   120
            TabIndex        =   173
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option23 
            Caption         =   "Crédito"
            Height          =   195
            Left            =   120
            TabIndex        =   172
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69000
         Picture         =   "FrmRepVentas.frx":4B79
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame Frame34 
         Caption         =   "Sucursal"
         Height          =   1215
         Left            =   -72240
         TabIndex        =   165
         Top             =   2040
         Width           =   2535
         Begin VB.ComboBox Combo9 
            Height          =   315
            Left            =   120
            TabIndex        =   166
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "Marca"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   163
         Top             =   2040
         Width           =   2535
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   120
            TabIndex        =   164
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame32 
         Caption         =   "Clasificación"
         Height          =   1215
         Left            =   -68400
         TabIndex        =   161
         Top             =   720
         Width           =   2535
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   120
            TabIndex        =   162
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "Tipo"
         Height          =   1215
         Left            =   -70920
         TabIndex        =   159
         Top             =   720
         Width           =   2415
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   120
            TabIndex        =   160
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   153
         Top             =   720
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPicker11 
            Height          =   375
            Left            =   600
            TabIndex        =   156
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39926
         End
         Begin VB.CheckBox Check20 
            Caption         =   "Ignorar fechas"
            Height          =   195
            Left            =   1200
            TabIndex        =   154
            Top             =   840
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPicker12 
            Height          =   375
            Left            =   2400
            TabIndex        =   158
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39926
         End
         Begin VB.Label Label33 
            Caption         =   "Al :"
            Height          =   255
            Left            =   2040
            TabIndex        =   157
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label32 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   480
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   152
         Top             =   3360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   -70560
         TabIndex        =   151
         Text            =   "Text18"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   375
         Left            =   -68640
         TabIndex        =   150
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Producciones"
         Height          =   255
         Left            =   -72240
         TabIndex        =   148
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Ventas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   147
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame Frame27 
         Caption         =   "Id-Producto"
         Height          =   1575
         Left            =   -71040
         TabIndex        =   143
         Top             =   480
         Width           =   2655
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   240
            TabIndex        =   146
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Picture         =   "FrmRepVentas.frx":754B
            Style           =   1  'Graphical
            TabIndex        =   145
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Buscar-Prod"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Picture         =   "FrmRepVentas.frx":9F1D
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame26 
         Height          =   615
         Left            =   -74160
         TabIndex        =   137
         Top             =   6840
         Width           =   7335
         Begin VB.TextBox Text16 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            TabIndex        =   139
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   138
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "CANTIDAD"
            Height          =   375
            Left            =   1080
            TabIndex        =   141
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "TOTAL:"
            Height          =   375
            Left            =   3840
            TabIndex        =   140
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Producto"
         Height          =   1455
         Left            =   -69000
         TabIndex        =   128
         Top             =   2160
         Width           =   2895
         Begin VB.CommandButton Command8 
            Caption         =   "Buscar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Picture         =   "FrmRepVentas.frx":C8EF
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            Height          =   375
            Left            =   360
            TabIndex        =   129
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Juegos de Reparacion"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   123
         Top             =   2160
         Width           =   4575
         Begin VB.OptionButton Option20 
            Caption         =   "Original"
            Height          =   255
            Left            =   2280
            TabIndex        =   134
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option19 
            Caption         =   "Remanofactura"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option18 
            Caption         =   "Comap"
            Height          =   255
            Left            =   2280
            TabIndex        =   126
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option17 
            Caption         =   "Camap"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option16 
            Caption         =   "Recargas"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Juegos de Reparacion"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   118
         Top             =   480
         Width           =   2055
         Begin VB.OptionButton Option21 
            Caption         =   "Original"
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Recargas"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Camap"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Comap"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Remanofactura"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   116
         Top             =   3720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72480
         Picture         =   "FrmRepVentas.frx":F2C1
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1740
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Frame Frame22 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   110
         Top             =   600
         Visible         =   0   'False
         Width           =   15
         Begin VB.ComboBox Combo4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            TabIndex        =   112
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            TabIndex        =   111
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Agente de Ventas"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Busqueda"
         Height          =   1575
         Left            =   -72840
         TabIndex        =   106
         Top             =   480
         Width           =   4455
         Begin VB.CheckBox Check15 
            Caption         =   "Asignar Agente"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   109
            Top             =   1200
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Reporte por Agente"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   108
            Top             =   1320
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Ventas"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   -69000
         TabIndex        =   105
         Top             =   720
         Width           =   2895
         Begin MSComCtl2.DTPicker DTPicker10 
            Height          =   375
            Left            =   720
            TabIndex        =   131
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39724
         End
         Begin MSComCtl2.DTPicker DTPicker9 
            Height          =   375
            Left            =   720
            TabIndex        =   130
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39724
         End
         Begin VB.Label Label21 
            Caption         =   "Al :"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sucursal"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   104
         Top             =   720
         Width           =   4695
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1200
            TabIndex        =   117
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccione Una Sucursal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   135
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Bus-Cliente"
         Height          =   255
         Left            =   -66360
         TabIndex        =   103
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Bus-Agente"
         Height          =   255
         Left            =   -66120
         TabIndex        =   102
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   101
         Top             =   2640
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame19 
         Caption         =   "Fechas"
         Height          =   1575
         Left            =   -68400
         TabIndex        =   96
         Top             =   480
         Width           =   2535
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   375
            Left            =   600
            TabIndex        =   97
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39640
         End
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   375
            Left            =   600
            TabIndex        =   98
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39600
         End
         Begin VB.Label Label17 
            Caption         =   "Del"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label18 
            Caption         =   "Al"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   -67200
         TabIndex        =   95
         Top             =   2880
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ordenado por"
         Height          =   1215
         Left            =   1680
         TabIndex        =   85
         Top             =   2160
         Width           =   1575
         Begin VB.OptionButton Option22 
            Caption         =   "No. Venta"
            Height          =   195
            Left            =   240
            TabIndex        =   168
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sucursal"
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   240
            TabIndex        =   86
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         Picture         =   "FrmRepVentas.frx":11C93
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Facturado y  No"
         Height          =   195
         Left            =   2640
         TabIndex        =   83
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Forma de Pago"
         Height          =   975
         Left            =   120
         TabIndex        =   79
         Top             =   2160
         Width           =   1455
         Begin VB.CheckBox Check4 
            Caption         =   "Tarjeta"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Cheque"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Efectivo"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Facturados"
         Height          =   195
         Left            =   2640
         TabIndex        =   78
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sucursal"
         Height          =   735
         Left            =   2520
         TabIndex        =   76
         Top             =   480
         Width           =   2175
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5400
         TabIndex        =   75
         Top             =   480
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Condiciones"
         Height          =   1455
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox Check6 
            Caption         =   "Por Fecha"
            Height          =   195
            Left            =   600
            TabIndex        =   70
            Top             =   1200
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   71
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39314
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   720
            TabIndex        =   72
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39314
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Agruar por"
         Height          =   1095
         Left            =   -73200
         TabIndex        =   65
         Top             =   1920
         Width           =   1575
         Begin VB.OptionButton Option5 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Sucursal"
            Height          =   195
            Left            =   240
            TabIndex        =   67
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Producto"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         Picture         =   "FrmRepVentas.frx":14665
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Solo Ventas de Credito"
         Height          =   255
         Left            =   -72360
         TabIndex        =   63
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Frame Frame14 
         Caption         =   "Forma de Pago"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   59
         Top             =   1920
         Width           =   1575
         Begin VB.CheckBox Check8 
            Caption         =   "Contado"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Tarjeta"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Solo Facturados"
         Height          =   255
         Left            =   -72360
         TabIndex        =   58
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame Frame15 
         Caption         =   "Sucursal"
         Height          =   855
         Left            =   -72480
         TabIndex        =   56
         Top             =   480
         Width           =   2175
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   -69600
         TabIndex        =   55
         Top             =   480
         Width           =   3495
      End
      Begin VB.Frame Frame16 
         Caption         =   "Condiciones"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox Check12 
            Caption         =   "Por Fecha"
            Height          =   195
            Left            =   600
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   720
            TabIndex        =   51
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39314
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   720
            TabIndex        =   52
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39314
         End
         Begin VB.Label Label23 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   -69240
         TabIndex        =   14
         Top             =   600
         Width           =   3135
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   375
            Left            =   960
            TabIndex        =   15
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39380
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39380
         End
         Begin VB.Label Label15 
            Caption         =   "Del :"
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label16 
            Caption         =   "Al :"
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -74160
         TabIndex        =   13
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70560
         Picture         =   "FrmRepVentas.frx":17037
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67320
         Picture         =   "FrmRepVentas.frx":19A09
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame17 
         Caption         =   "Ordenar"
         Height          =   975
         Left            =   -69240
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
         Begin VB.OptionButton Option7 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Por Cantidad"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check13 
         Caption         =   "No Facturadas"
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   10
         Top             =   3720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   4
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   3975
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Compras"
         TabPicture(0)   =   "FrmRepVentas.frx":1C3DB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame6"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Abonos"
         TabPicture(1)   =   "FrmRepVentas.frx":1C3F7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ListView5"
         Tab(1).Control(1)=   "Frame18"
         Tab(1).Control(2)=   "Frame9"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Abonos Clientes"
         TabPicture(2)   =   "FrmRepVentas.frx":1C413
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListView8"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame6 
            Caption         =   "Totales"
            Height          =   1335
            Left            =   5880
            TabIndex        =   39
            Top             =   2520
            Width           =   2775
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label7 
               Caption         =   "Total :"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label6 
               Caption         =   "IVA :"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label5 
               Caption         =   "Subtotal :"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Creditos"
            Height          =   1335
            Left            =   3000
            TabIndex        =   32
            Top             =   2520
            Width           =   2775
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label8 
               Caption         =   "Total :"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label9 
               Caption         =   "IVA :"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label10 
               Caption         =   "Subtotal :"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Contado"
            Height          =   1335
            Left            =   120
            TabIndex        =   25
            Top             =   2520
            Width           =   2775
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox Text10 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label11 
               Caption         =   "Total :"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label12 
               Caption         =   "IVA :"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "Subtotal :"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Total de Abonos"
            Height          =   855
            Left            =   -72480
            TabIndex        =   23
            Top             =   3000
            Width           =   2175
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   360
               Width           =   1935
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Total de Deuda"
            Height          =   855
            Left            =   -74760
            TabIndex        =   21
            Top             =   3000
            Width           =   2175
            Begin VB.TextBox Text13 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   360
               Width           =   1935
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView5 
            Height          =   2415
            Left            =   -74760
            TabIndex        =   47
            Top             =   480
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4260
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView8 
            Height          =   3255
            Left            =   -74880
            TabIndex        =   48
            Top             =   480
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5741
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   4800
         TabIndex        =   89
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1935
         Left            =   -70200
         TabIndex        =   90
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   91
         Top             =   3120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7646
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label26 
         Height          =   375
         Left            =   -73320
         TabIndex        =   136
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   4800
         TabIndex        =   94
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   -70200
         TabIndex        =   93
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Cliente 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   92
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -66240
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9600
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmRepVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim IdClien As String
Dim IdClien1 As String
Dim IdClieRep As String
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim StrRep5 As String
Dim StrRep8 As String
Dim StrRep9 As String
Dim parcial As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
    Me.Check13.Value = 0
End Sub
Private Sub Check10_Click()
    If IdClien1 <> "" And Check10.Value = 1 Then
        Check10.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Check12_Click()
    If Check12.Value = 1 Then
        DTPicker3.Enabled = True
        DTPicker4.Enabled = True
    Else
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
    End If
End Sub
Private Sub Check13_Click()
    Me.Check1.Value = 0
End Sub
Private Sub Check14_Click()
    frmrepventasdetalle.Show vbModal
End Sub
Private Sub Check15_Click()
    Combo3.Enabled = True
    Command3.Visible = True
    Combo4.Enabled = True
    Check16.Value = 0
    Check17.Value = 0
    Text14.Visible = False
    Command7.Visible = False
    Option12.Visible = False
    Option13.Visible = False
    Option14.Visible = False
    Option15.Visible = False
End Sub
Private Sub Check16_Click()
    Combo3.Enabled = True
    Combo4.Enabled = True
    Command6.Visible = True
    Command3.Visible = False
    Command7.Visible = False
    Check15.Value = 0
    Check17.Value = 0
    Text14.Visible = True
    Option10.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option15.Visible = True
End Sub
Private Sub Check17_Click()
    Command6.Visible = False
    Command7.Visible = True
    Option12.Visible = True
    Option13.Visible = True
    Option14.Visible = True
    Option15.Visible = True
    Text14.Visible = True
    Check16.Value = 0
    Check15.Value = 0
End Sub
Private Sub Check18_Click()
    Check19.Value = 0
    Frame11.Visible = True
End Sub
Private Sub Check19_Click()
    Check18.Value = 0
    'Frame28.Visible = True
    Frame11.Visible = False
End Sub
Private Sub Check2_Click()
    If IdClien <> "" And Check2.Value = 1 Then
        Check2.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Check3_Click()
    If IdClien <> "" And Check3.Value = 1 Then
        Check3.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Check4_Click()
    If IdClien <> "" And Check4.Value = 1 Then
        Check4.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Check6_Click()
    If Check6.Value = 1 Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
End Sub
Private Sub Check8_Click()
    If IdClien1 <> "" And Check8.Value = 1 Then
        Check8.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Check9_Click()
    If IdClien1 <> "" And Check9.Value = 1 Then
        Check9.Value = 0
        MsgBox "NO SE PUEDE SELECCIONAR FORMA DE PAGO SI SELECCIONA UN CLIENTE", vbInformation, "SACC"
    End If
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim AndVar As String
    Dim SubTotVar As String
    Dim IvaVar As String
    Dim TotalVar As String
    Dim SubTotVarCred As String
    Dim IvaVarCred As String
    Dim TotalVarCred As String
    Dim SubTotVarCont As String
    Dim IvaVarCont As String
    Dim TotalVarCont As String
    Dim sWhere As String
    SubTotVar = "0"
    Text11.Text = "0"
    IvaVar = "0"
    TotalVar = "0"
    SubTotVarCred = "0"
    IvaVarCred = "0"
    TotalVarCred = "0"
    SubTotVarCont = "0"
    TotalVarCont = "0"
    IvaVarCont = "0"
    AndVar = "0"
    sBuscar = "SELECT * FROM VENTAS WHERE "
    If Combo1.Text <> "" Then
        If Combo1.Text <> "<TODAS>" Then
            sBuscar = sBuscar & " SUCURSAL = '" & Combo1.Text & "' "
            AndVar = "1"
        End If
    End If
    If IdClien <> "" Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "ID_CLIENTE IN (" & IdClien & ") "
        AndVar = "1"
    End If
    If Option23.Value Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "UNA_EXIBICION = 'N'"
        AndVar = "1"
    Else
        If Option24.Value Then
            If AndVar = "1" Then
                sBuscar = sBuscar & "AND "
            End If
            sBuscar = sBuscar & "UNA_EXIBICION = 'S'"
            AndVar = "1"
        End If
    End If
    If Check1.Value = 1 Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
       sBuscar = sBuscar & " FACTURADO = '1' "
       AndVar = "1"
   Else
       If Check13.Value = 1 Then
            If AndVar = "1" Then
                sBuscar = sBuscar & "AND "
            End If
            sBuscar = sBuscar & " FACTURADO = '0' "
            AndVar = "1"
        Else
            If Check5.Value = "1" Then
                If AndVar = "1" Then
                    sBuscar = sBuscar & "AND "
                End If
                sBuscar = sBuscar & " FACTURADO IN (0, 1) "
                AndVar = "1"
            End If
        End If
    End If
    If Not (Check2.Value = 1 And Check3.Value = 1 And Check3.Value = 1) Then
        If Check2.Value = 1 Then
            If AndVar = "1" Then
                sBuscar = sBuscar & "AND "
            End If
            sBuscar = sBuscar & "TIPO_PAGO = 'C' "
            AndVar = "1"
        End If
        If Check3.Value = 1 Then
            If AndVar = "1" Then
                If Check2.Value = 1 Then
                    sBuscar = sBuscar & "OR "
                Else
                    sBuscar = sBuscar & "AND "
                End If
            End If
            sBuscar = sBuscar & "TIPO_PAGO = 'H' "
            AndVar = "1"
        End If
        If Check4.Value = 1 Then
            If AndVar = "1" Then
                If Check2.Value = 1 Or Check3.Value = 1 Then
                    sBuscar = sBuscar & "OR "
                Else
                    sBuscar = sBuscar & "AND "
                End If
            End If
            sBuscar = sBuscar & "TIPO_PAGO = 'T' "
            AndVar = "1"
        End If
    End If
    If Check6.Value = 1 Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997' "
    End If
    If Option1.Value = True Then
        sBuscar = sBuscar & "ORDER BY FECHA"
    Else
        If Option2.Value = True Then
            sBuscar = sBuscar & "ORDER BY NOMBRE"
        Else
            If Option3.Value = True Then
                sBuscar = sBuscar & "ORDER BY SUCURSAL"
            Else
                sBuscar = sBuscar & "ORDER BY ID_VENTA"
            End If
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("FECHA"))
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(1) = tRs.Fields("SUCURSAL")
            If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(2) = tRs.Fields("ID_VENTA")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(3) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("FORMA_PAGO")) Then tLi.SubItems(4) = tRs.Fields("FORMA_PAGO")
            If Not IsNull(tRs.Fields("UUID")) Then tLi.SubItems(10) = tRs.Fields("UUID")
            If Not IsNull(tRs.Fields("UNA_EXIBICION")) Then
                If tRs.Fields("UNA_EXIBICION") = "S" Then
                    tLi.SubItems(5) = "CONTADO"
                Else
                    tLi.SubItems(5) = "CREDITO"
                End If
            End If
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(6) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then
                tLi.SubItems(7) = tRs.Fields("SUBTOTAL")
                SubTotVar = CDbl(SubTotVar) + CDbl(tRs.Fields("SUBTOTAL"))
                If tRs.Fields("UNA_EXIBICION") = "S" Then
                    SubTotVarCont = CDbl(SubTotVarCont) + CDbl(tRs.Fields("SUBTOTAL"))
                Else
                    SubTotVarCred = CDbl(SubTotVarCred) + CDbl(tRs.Fields("SUBTOTAL"))
                End If
            End If
            If Not IsNull(tRs.Fields("IVA")) Then
                tLi.SubItems(8) = tRs.Fields("IVA")
                IvaVar = CDbl(IvaVar) + CDbl(tRs.Fields("IVA"))
                If tRs.Fields("UNA_EXIBICION") = "S" Then
                    IvaVarCont = CDbl(IvaVarCont) + CDbl(tRs.Fields("IVA"))
                Else
                    IvaVarCred = CDbl(IvaVarCred) + CDbl(tRs.Fields("IVA"))
                End If
            End If
            If Not IsNull(tRs.Fields("TOTAL")) Then
                tLi.SubItems(9) = tRs.Fields("TOTAL")
                TotalVar = CDbl(TotalVar) + CDbl(tRs.Fields("TOTAL"))
                If tRs.Fields("UNA_EXIBICION") = "S" Then
                    TotalVarCont = CDbl(TotalVarCont) + CDbl(tRs.Fields("TOTAL"))
                Else
                    TotalVarCred = CDbl(TotalVarCred) + CDbl(tRs.Fields("TOTAL"))
                End If
            End If
            tRs.MoveNext
        Loop
        ListView5.ListItems.Clear
        StrRep = sBuscar
        sBuscar = Replace(sBuscar, "SELECT * FROM VENTAS WHERE ", "SELECT * FROM VsAbonosCXC WHERE ")
        sBuscar = Replace(sBuscar, "UNA_EXIBICION = 'N' ", "UNA_EXIBICION = 'S' ")
        sBuscar = Replace(sBuscar, "AND TIPO_PAGO = 'C' OR TIPO_PAGO = 'H' OR TIPO_PAGO = 'T'", "")
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView5.ListItems.Add(, , tRs.Fields("FECHA"))
                If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(1) = tRs.Fields("ID_VENTA")
                If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                    tLi.SubItems(4) = tRs.Fields("CANT_ABONO")
                    Text11.Text = Format(CDbl(Text11.Text) + CDbl(tRs.Fields("CANT_ABONO")), "###,###,##0.00")
                End If
                If Not IsNull(tRs.Fields("NO_CHEQUE")) Then tLi.SubItems(5) = tRs.Fields("NO_CHEQUE")
                If Not IsNull(tRs.Fields("BANCO")) Then tLi.SubItems(6) = tRs.Fields("BANCO")
                If Not IsNull(tRs.Fields("FECHA_CHEQUE")) Then tLi.SubItems(7) = tRs.Fields("FECHA_CHEQUE")
                If Not IsNull(tRs.Fields("REFERENCIA")) Then tLi.SubItems(8) = tRs.Fields("REFERENCIA")
                tRs.MoveNext
            Loop
        End If
        Text2.Text = Format(SubTotVar, "###,###,##0.00")
        Text3.Text = Format(IvaVar, "###,###,##0.00")
        Text4.Text = Format(TotalVar, "###,###,##0.00")
        Text7.Text = Format(SubTotVarCred, "###,###,##0.00")
        Text6.Text = Format(IvaVarCred, "###,###,##0.00")
        Text5.Text = Format(TotalVarCred, "###,###,##0.00")
        Text8.Text = Format(SubTotVarCont, "###,###,##0.00")
        Text9.Text = Format(IvaVarCont, "###,###,##0.00")
        Text10.Text = Format(TotalVarCont, "###,###,##0.00")
    End If
    Exit Sub
ManejaError:
    Err.Clear
    MsgBox "DEBE SELECCIONAR AL MENOS UN METODO DE FLITRACIÓN", vbInformation, "SACC"
End Sub
Private Sub Command10_Click()
    VentasPorMes
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim AndVar As String
    AndVar = "0"
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_MAXIMO, FECHA, NOMBRE, PRECIO_MINIMO, ID_VENTA, FOLIO FROM VsRepVentas WHERE "
    If Combo2.Text <> "" Then
        If Combo2.Text <> "<TODAS>" Then
            sBuscar = sBuscar & " SUCURSAL = '" & Combo2.Text & "' "
            AndVar = "1"
        End If
    End If
    If IdClien1 <> "" Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "ID_CLIENTE = " & IdClien1 & " "
        AndVar = "1"
    End If
   If Check14 = 1 Then
            If Combo3.Text <> "<TODAS>" Then
            sBuscar = sBuscar & " ID_PRODUCTO = '" & Combo3.Text & "' "
            AndVar = "1"
        End If
    End If
    If Check11.Value = 1 Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "FACTURADO = '1' "
        AndVar = "1"
    Else
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "FACTURADO IN (1, 0) "
        AndVar = "1"
    End If
    If Not (Check8.Value = 1 And Check9.Value = 1 And Check10.Value = 1) Then
        If Check8.Value = 1 Then
            If AndVar = "1" Then
                sBuscar = sBuscar & "AND "
            End If
            sBuscar = sBuscar & "FORMA_PAGO = 'C' "
            AndVar = "1"
        End If
        If Check9.Value = 1 Then
            If AndVar = "1" Then
                If Check8.Value = 1 Then
                    sBuscar = sBuscar & "OR "
                Else
                    sBuscar = sBuscar & "AND "
                End If
            End If
            sBuscar = sBuscar & "FORMA_PAGO = 'H' "
            AndVar = "1"
        End If
        If Check10.Value = 1 Then
            If AndVar = "1" Then
                If Check8.Value = 1 Or Check9.Value = 1 Then
                    sBuscar = sBuscar & "OR "
                Else
                    sBuscar = sBuscar & "AND "
                End If
            End If
            sBuscar = sBuscar & "FORMA_PAGO = 'T' "
            AndVar = "1"
        End If
    End If
    If Check7.Value = 1 Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "UNA_EXIBICION = 'N' "
        AndVar = "1"
    End If
    If Check12.Value = 1 Then
        If AndVar = "1" Then
            sBuscar = sBuscar & "AND "
        End If
        sBuscar = sBuscar & "FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " 23:59:59.997'"
    End If
    If Option4.Value = True Then
        sBuscar = sBuscar & "GROUP BY ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_MAXIMO, PRECIO_MINIMO, FECHA, NOMBRE, ID_VENTA, FOLIO"
    End If
    If Option5.Value = True Then
        sBuscar = sBuscar & "GROUP BY ID_CLIENTE, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_MAXIMO, FECHA, PRECIO_MINIMO, NOMBRE, ID_VENTA, FOLIO"
    End If
    If Option6.Value = True Then
        sBuscar = sBuscar & "GROUP BY SUCURSAL, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_MAXIMO, PRECIO_MINIMO, FECHA, NOMBRE, ID_VENTA, FOLIO"
    End If
    StrRep9 = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    ListView4.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("ID_PRODUCTO") <> "                         " Then
                Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("PRECIO_MAXIMO")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_MAXIMO")
                If Not IsNull(tRs.Fields("PRECIO_MINIMO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_MINIMO")
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(6) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(7) = tRs.Fields("ID_VENTA")
                If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(8) = tRs.Fields("FOLIO")
            End If
            tRs.MoveNext
        Loop
        StrRep2 = sBuscar
    End If
    Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbInformation, "SACC"
    Err.Clear
    
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE CLIENTE SET AGENTE= '" & Combo3.Text & "' WHERE  NOMBRE='" & Combo4.Text & "'"
    cnn.Execute (sBuscar)
    sBuscar = "UPDATE CLIENTE SET ASIG= 'S' WHERE  NOMBRE= '" & Combo4.Text & "'"
    cnn.Execute (sBuscar)
    Combo3.Text = ""
    Combo4.Text = ""
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView6.ListItems.Clear
    sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Replace(Text12.Text, " ", "%") & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView6.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command5_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView7.ListItems.Clear
    If IdClieRep <> "" Then
        sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, SUM(VENTAS_DETALLE.CANTIDAD) AS CANTIDAD, ALMACEN3.Clasificacion, ALMACEN3.Marca FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE VENTAS.ID_CLIENTE = " & IdClieRep & " AND VENTAS.FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & " 23:59:59.997' GROUP BY VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, ALMACEN3.CLASIFICACION, ALMACEN3.MARCA"
    Else
        sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, SUM(VENTAS_DETALLE.CANTIDAD) AS CANTIDAD, ALMACEN3.Clasificacion, ALMACEN3.Marca FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE VENTAS.FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & " 23:59:59.997' GROUP BY VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, ALMACEN3.CLASIFICACION, ALMACEN3.MARCA"
    End If
    If Option7.Value = True Then
       sBuscar = sBuscar & " ORDER BY VENTAS_DETALLE.ID_PRODUCTO"
       End If
    If Option8.Value = True Then
            sBuscar = sBuscar & " ORDER BY CANTIDAD DESC "
       End If
    If Option9.Value = True Then
        sBuscar = sBuscar & " ORDER BY VENTAS_DETALLE.DESCRIPCION"
    End If
      '"CANTIDAD"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
              Set tLi = ListView7.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("CLASIFICACION")) Then tLi.SubItems(3) = tRs.Fields("CLASIFICACION")
            If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(4) = tRs.Fields("MARCA")
           tRs.MoveNext
        Loop
    End If
    StrRep3 = sBuscar
    StrRep3 = Replace(StrRep3, "SUM(CANTIDAD) AS CANTIDAD", "CANTIDAD")
    StrRep3 = Replace(StrRep3, "GROUP BY ID_PRODUCTO, Descripcion", "")
    IdClieRep = ""
End Sub
Private Sub Command6_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Frame9.Visible = True
    ListView9.ListItems.Clear
    If Option11 = True Then
        sBuscar = "SELECT ID_CLIENTE,NOMBRE,SUM(TOTAL) AS TOTAL,AGENTE FROM vsrepagente WHERE  NOMBRE LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY NOMBRE,AGENTE,ID_CLIENTE"
    End If
    If Option10 = True Then
        sBuscar = "SELECT ID_CLIENTE,NOMBRE,SUM(TOTAL) AS TOTAL,AGENTE FROM vsrepagente WHERE  AGENTE LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY NOMBRE,AGENTE,ID_CLIENTE"
    End If
    StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView9.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("AGENTE")) Then tLi.SubItems(3) = tRs.Fields("AGENTE")
            tRs.MoveNext
        Loop
        StrRep4 = sBuscar
    End If
End Sub
Private Sub Command7_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim AndVar As String
    Dim parcial As String
    ListView9.ListItems.Clear
    Text17.Text = ""
    Text16.Text = ""
    If MsgBox("QUE REPORTE DESEA GENERAR  ¿DESEA GENERAR REPORTE DE AGENTE?  CLICK SI, AL PRESIONAR CLICKNO, SE GENERA REPORTE GENERALIZADO", vbYesNo + vbCritical + vbDefaultButton1) = vbYes Then
        If Check17.Value = 1 Then
            If Option12 = True Then
                sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, USUARIO, PRECIO_VENTA, ID_VENTA, NOMBRE, FECHA, FACTURADO, UNA_EXIBICION FROM vsrepagente2 WHERE UNA_EXIBICION='S' AND FOlIO NOT IN ('CANCELADO') AND  FACTURADO NOT IN ('2') AND CLASIFICACION='RECARGA' AND USUARIO LIKE '%" & Text14.Text & "%' AND FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997'"
            End If
        End If
        If Check17.Value = 1 Then
            If Option15 = True Then
                sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, USUARIO, PRECIO_VENTA, ID_VENTA, NOMBRE, FECHA, FACTURADO, UNA_EXIBICION FROM vsrepagente2 WHERE UNA_EXIBICION='S' AND FOlIO NOT IN ('CANCELADO') AND  FACTURADO NOT IN ('2') AND CLASIFICACION='REMANUFACTURA' AND USUARIO LIKE '%" & Text14.Text & "%' AND FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' "
            End If
        End If
        If Check17.Value = 1 Then
            If Option21 = True Then
                sBuscar = "SELECT ID_PRODUCTO, SUM(CANTIDAD) AS CANTIDAD, 'TODOS' AS USUARIO, '$' AS PRECIO_VENTA, '#' AS ID_VENTA, 'CLIENTE' AS NOMBRE, 'PERIODO' AS FECHA, 'TODO' AS FACTURADO, 'TODO' AS UNA_EXIBICION FROM vsagente3 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='ORIGINAL' AND ID_PRODUCTO  LIKE '%" & Text14.Text & "%' AND FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        If Check17.Value = 1 Then
            If Option14 = True Then
                sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, USUARIO, PRECIO_VENTA, ID_VENTA, NOMBRE, FECHA, FACTURADO, UNA_EXIBICION FROM vsrepagente2 WHERE UNA_EXIBICION = 'S' AND FOlIO NOT IN ('CANCELADO') AND FACTURADO NOT IN ('2') AND CLASIFICACION = 'COMPATIBLE' AND USUARIO LIKE '%" & Text14.Text & "%' AND FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997'"
            End If
        End If
        If Check17.Value = 1 Then
            If Option13 = True Then
                sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, USUARIO, PRECIO_VENTA, ID_VENTA, NOMBRE, FECHA, FACTURADO, UNA_EXIBICION FROM vsrepagente2 WHERE UNA_EXIBICION = 'S' AND FOlIO NOT IN ('CANCELADO') AND FACTURADO NOT IN ('2') AND CLASIFICACION = 'CAMBIO' AND USUARIO LIKE '%" & Text14.Text & "%' AND FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997'"
            End If
        End If
        If Check17.Value <> 1 Then
            sBuscar = sBuscar & " ORDER BY FECHA DESC"
        End If
        Set tRs = cnn.Execute(sBuscar)
        Dim totcant As Integer
        Dim totales As Double
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView9.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
                tLi.SubItems(3) = tRs.Fields("FECHA")
                tLi.SubItems(4) = tRs.Fields("CANTIDAD")
                 
                If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(5) = tRs.Fields("USUARIO")
                totcant = CDbl(totcant) + CDbl(tRs("CANTIDAD"))
                If Check17.Value <> 1 Then
                    tLi.SubItems(6) = CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD"))
                    totales = CDbl(totales) + CDbl(tLi.SubItems(6))
                Else
                    tLi.SubItems(6) = "$"
                End If
                tRs.MoveNext
            Loop
            StrRep4 = sBuscar
            Text17.Text = totcant
            Text16.Text = totales
        End If
    Else
        If Check17.Value = 1 Then
            If Option12 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsrepagente2 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='RECARGA'  AND ID_PRODUCTO LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        If Check17.Value = 1 Then
            If Option15 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsrepagente2 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='REMANUFACTURA' AND ID_PRODUCTO LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        If Check17.Value = 1 Then
            If Option14 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsrepagente2 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='COMPATIBLE' AND ID_PRODUCTO  LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        If Check17.Value = 1 Then
            If Option21 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsagente3 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='ORIGINAL' AND ID_PRODUCTO  LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        If Check17.Value = 1 Then
            If Option13 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsrepagente2 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='CAMBIO' AND ID_PRODUCTO LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            End If
        End If
        sBuscar = sBuscar & " ORDER BY CANTIDAD DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                 Set tLi = ListView9.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                 tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                 tRs.MoveNext
            Loop
            StrRep4 = sBuscar
        End If
    End If
End Sub
Private Sub Command8_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim AndVar As String
    Dim parcial As String
    Dim Cont As Integer
    ListView10.ListItems.Clear
    Cont = 0
    If Check18.Value = 1 Then
        If Combo5.Text <> "" Then
            If Option16 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD,SUCURSAL FROM vsrepagente2 WHERE CLASIFICACION='RECARGA' AND SUCURSAL='" & Combo5.Text & "' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO,SUCURSAL"
                Command8.Enabled = True
            End If
            If Option19 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD,SUCURSAL FROM vsrepagente2 WHERE CLASIFICACION='REMANUFACTURA' AND SUCURSAL='" & Combo5.Text & "' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO,SUCURSAL"
                Command8.Enabled = True
            End If
            If Option18 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD,SUCURSAL FROM vsrepagente2 WHERE CLASIFICACION='COMPATIBLE'AND SUCURSAL='" & Combo5.Text & "' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO,SUCURSAL"
                Command8.Enabled = True
            End If
            If Option17 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD,SUCURSAL FROM vsrepagente2 WHERE CLASIFICACION='CAMBIO' AND SUCURSAL='" & Combo5.Text & "' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO,SUCURSAL"
                Command8.Enabled = True
            End If
            If Option20 = True Then
                sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD,SUCURSAL FROM vsrepagente2 WHERE CLASIFICACION='ORIGINAL' AND SUCURSAL='" & Combo5.Text & "' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO,SUCURSAL"
                Command8.Enabled = True
            End If
            sBuscar = sBuscar & " ORDER BY CANTIDAD DESC"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView10.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                    tLi.SubItems(2) = tRs.Fields("SUCURSAL")
                    tRs.MoveNext
                Loop
                StrRep5 = sBuscar
            End If
        End If
    End If
    If Check19.Value = 1 Then
        If Option16 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM producciones WHERE CLASIFICACION='RECARGA' AND TIPO='P' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            Command8.Enabled = True
        End If
        If Option19 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM producciones WHERE CLASIFICACION='REMANUFACTURA' AND TIPO='P' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            Command8.Enabled = True
        End If
        If Option18 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM producciones WHERE CLASIFICACION='COMPATIBLE'ANDTIPO='P' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            Command8.Enabled = True
        End If
        If Option17 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM producciones WHERE CLASIFICACION='CAMBIO' AND TIPO='P' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            Command8.Enabled = True
        End If
        If Option20 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM producciones WHERE CLASIFICACION='ORIGINAL' AND TIPO='P' AND ID_PRODUCTO LIKE '%" & Text15.Text & "%' AND  FECHA BETWEEN'" & DTPicker9.Value & " ' AND '" & DTPicker10.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
            Command8.Enabled = True
        End If
        sBuscar = sBuscar & " ORDER BY CANTIDAD DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView10.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                tRs.MoveNext
            Loop
            StrRep5 = sBuscar
        End If
    End If
End Sub
Private Sub Command9_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim AndVar As String
    Dim parcial As String
    ListView9.ListItems.Clear
    Text17.Text = ""
    Text16.Text = ""
    sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsagente3 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='ORIGINAL' AND ID_PRODUCTO  LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
    If Check17.Value = 1 Then
        If Option13 = True Then
            sBuscar = "SELECT ID_PRODUCTO,SUM(CANTIDAD) AS CANTIDAD FROM vsrepagente2 WHERE FOlIO NOT IN ('CANCELADO')  AND CLASIFICACION='CAMBIO' AND ID_PRODUCTO LIKE '%" & Text14.Text & "%' AND  FECHA BETWEEN'" & DTPicker7.Value & " ' AND '" & DTPicker8.Value & " 23:59:59.997' GROUP BY ID_PRODUCTO"
        End If
    End If
    sBuscar = sBuscar & " ORDER BY CANTIDAD DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView9.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
        StrRep4 = sBuscar
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sBu As String
    Dim tRs4 As ADODB.Recordset
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DTPicker3.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker4.Value = Format(Date, "dd/mm/yyyy")
    DTPicker3.Enabled = False
    DTPicker4.Enabled = False
    DTPicker5.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker6.Value = Format(Date, "dd/mm/yyyy")
    DTPicker9.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker10.Value = Format(Date, "dd/mm/yyyy")
    DTPicker7.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker8.Value = Format(Date, "dd/mm/yyyy")
    DTPicker11.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker12.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Nota", 1000
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Forma de Pago", 1500
        .ColumnHeaders.Add , , "Tipo de Venta", 1500
        .ColumnHeaders.Add , , "Cliente", 5500
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "IVA", 1500
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "UUID", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "Direccion", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "Direccion", 1500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio Maximo", 1500
        .ColumnHeaders.Add , , "Precio Minimo", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "Venta", 1500
        .ColumnHeaders.Add , , "Factura", 1500
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Venta", 1000
        .ColumnHeaders.Add , , "Factura", 1000
        .ColumnHeaders.Add , , "Cliente", 5500
        .ColumnHeaders.Add , , "Cantidad Abonada", 1500
        .ColumnHeaders.Add , , "Numero de Cheque", 1000
        .ColumnHeaders.Add , , "Banco", 1500
        .ColumnHeaders.Add , , "Fecha del Cheque", 1500
        .ColumnHeaders.Add , , "Referencia", 1500
    End With
    With ListView6
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 5450
    End With
    With ListView7
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Producto", 1800
        .ColumnHeaders.Add , , "Descripcion", 6000
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "cLA", 0
        .ColumnHeaders.Add , , "mARCA", 0
    End With
    With ListView8
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Nombre", 4000
        .ColumnHeaders.Add , , "Abono", 1500
        .ColumnHeaders.Add , , "Deuda", 1500
        .ColumnHeaders.Add , , "Deuda Actual", 1500
    End With
    With ListView9
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 1500
        .ColumnHeaders.Add , , "Nombre", 3000
        .ColumnHeaders.Add , , "Producto", 1800
        .ColumnHeaders.Add , , "Fecha", 1800
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Asignado", 0
        .ColumnHeaders.Add , , "TOTAL_$", 1000
    End With
    With ListView10
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 1500
        .ColumnHeaders.Add , , "Cantidad   ", 3000
        .ColumnHeaders.Add , , "Sucursal", 1800
    End With
    With ListView11
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 1500
        .ColumnHeaders.Add , , "Descripcion   ", 3000
        .ColumnHeaders.Add , , "Cantidad", 1800
        .ColumnHeaders.Add , , "Sucursal", 1800
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    Combo2.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            Combo2.AddItem tRs.Fields("NOMBRE")
            Combo9.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE,PUESTO FROM USUARIOS WHERE  PUESTO='VENTAS' GROUP BY NOMBRE,PUESTO"
    Set tRs = cnn.Execute(sBuscar)
    Combo3.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo3.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    sBu = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs4 = cnn.Execute(sBu)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs4.EOF
            Combo5.AddItem tRs4.Fields("NOMBRE")
            tRs4.MoveNext
        Loop
    End If
    Combo6.Clear
    sBuscar = "SELECT TIPO FROM ALMACEN3 GROUP BY TIPO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo6.AddItem tRs.Fields("TIPO")
            tRs.MoveNext
        Loop
    End If
    Combo7.Clear
    sBuscar = "SELECT CLASIFICACION FROM ALMACEN3 GROUP BY CLASIFICACION"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo7.AddItem tRs.Fields("CLASIFICACION")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image10_Click()
On Error GoTo ManejaError
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If Ruta <> "" Then
        If StrRep <> "" And SSTab1.Tab = 0 Then
            For Con = 1 To ListView1.ColumnHeaders.Count
                NumColum = ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
        Else
            If StrRep9 <> "" Then
                For Con = 1 To ListView4.ColumnHeaders.Count
                    NumColum = ListView4.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView4.ColumnHeaders(Con).Text & Chr(9)
                Next
            Else
                If StrRep4 <> "" And SSTab1.Tab = 3 Then
                    For Con = 1 To ListView9.ColumnHeaders.Count
                        NumColum = ListView9.ColumnHeaders.Count
                        StrCopi = StrCopi & ListView9.ColumnHeaders(Con).Text & Chr(9)
                    Next
                Else
                    If StrRep2 <> "" And SSTab1.Tab = 1 Then
                        For Con = 1 To ListView4.ColumnHeaders.Count
                            NumColum = ListView4.ColumnHeaders.Count
                            StrCopi = StrCopi & ListView4.ColumnHeaders(Con).Text & Chr(9)
                        Next
                    Else
                        If StrRep5 <> "" And SSTab1.Tab = 4 Then
                            For Con = 1 To ListView10.ColumnHeaders.Count
                                NumColum = ListView10.ColumnHeaders.Count
                                StrCopi = StrCopi & ListView10.ColumnHeaders(Con).Text & Chr(9)
                            Next
                        Else
                            If StrRep3 <> "" And SSTab1.Tab = 2 Then
                                For Con = 1 To ListView7.ColumnHeaders.Count
                                    NumColum = ListView7.ColumnHeaders.Count
                                    StrCopi = StrCopi & ListView7.ColumnHeaders(Con).Text & Chr(9)
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
        StrCopi = StrCopi & Chr(13)
        If StrRep <> "" And SSTab1.Tab = 0 Then
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
        Else
            If StrRep9 <> "" Then
                For Con = 1 To ListView4.ListItems.Count
                    StrCopi = StrCopi & ListView4.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            Else
                If StrRep4 <> "" And SSTab1.Tab = 3 Then
                    For Con = 1 To ListView9.ListItems.Count
                        StrCopi = StrCopi & ListView9.ListItems.Item(Con) & Chr(9)
                        For Con2 = 1 To NumColum - 1
                            StrCopi = StrCopi & ListView9.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                        Next
                        StrCopi = StrCopi & Chr(13)
                    Next
                Else
                    If StrRep2 <> "" And SSTab1.Tab = 1 Then
                        For Con = 1 To ListView4.ListItems.Count
                            StrCopi = StrCopi & ListView4.ListItems.Item(Con) & Chr(9)
                            For Con2 = 1 To NumColum - 1
                                StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                            Next
                            StrCopi = StrCopi & Chr(13)
                        Next
                    Else
                        If StrRep5 <> "" And SSTab1.Tab = 4 Then
                            For Con = 1 To ListView10.ListItems.Count
                                StrCopi = StrCopi & ListView10.ListItems.Item(Con) & Chr(9)
                                For Con2 = 1 To NumColum - 1
                                    StrCopi = StrCopi & ListView10.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                                Next
                                StrCopi = StrCopi & Chr(13)
                            Next
                        Else
                            If StrRep3 <> "" And SSTab1.Tab = 2 Then
                                For Con = 1 To ListView7.ListItems.Count
                                    StrCopi = StrCopi & ListView7.ListItems.Item(Con) & Chr(9)
                                    For Con2 = 1 To NumColum - 1
                                        StrCopi = StrCopi & ListView7.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                                    Next
                                    StrCopi = StrCopi & Chr(13)
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'archivo TXT
        Dim foo As Integer
        foo = FreeFile
        Open Ruta For Output As #foo
            Print #foo, StrCopi
        Close #foo
    End If
    ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim SUMA As String
    Dim Total As Double
    ConPag = 1
    Total = "0"
    SUMA = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Not (ListView1.ListItems.Count = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\VentasPeriodo.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        If Combo1.Text = "" Then
            oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS POR PERIODO", "F3", 8, hCenter
        Else
            oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS POR PERIODO DE SUCURSAL " & Combo1.Text, "F3", 8, hCenter
        End If
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "Venta", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 340, "Cliente", "F2", 8, hCenter
        oDoc.WTextBox Posi, 390, 20, 55, "Fecha", "F2", 8, hCenter
        oDoc.WTextBox Posi, 425, 20, 55, "Folio", "F2", 8, hCenter
        oDoc.WTextBox Posi, 460, 20, 55, "Total", "F2", 8, hCenter
        oDoc.WTextBox Posi, 515, 20, 55, "Tipo", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        For Cont = 1 To ListView1.ListItems.Count
            oDoc.WTextBox Posi, 10, 20, 60, ListView1.ListItems(Cont).SubItems(2), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 340, ListView1.ListItems(Cont).SubItems(6), "F3", 7, hLeft
            oDoc.WTextBox Posi, 390, 20, 55, Format(ListView1.ListItems(Cont), "dd/mm/yyyy"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 435, 20, 55, ListView1.ListItems(Cont).SubItems(3), "F3", 7, hLeft
            oDoc.WTextBox Posi, 460, 20, 50, Format(ListView1.ListItems(Cont).SubItems(9), "###,###,##0.00"), "F3", 7, hRight
            Total = Total + ListView1.ListItems(Cont).SubItems(9)
            oDoc.WTextBox Posi, 515, 20, 55, ListView1.ListItems(Cont).SubItems(5), "F3", 7, hLeft
            Posi = Posi + 12
            If Posi >= 700 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs1 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS POR PERIODO DE SUCURSAL " & Sucursal, "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "Venta", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 340, "Cliente", "F2", 8, hCenter
                oDoc.WTextBox Posi, 390, 20, 55, "Fecha", "F2", 8, hCenter
                oDoc.WTextBox Posi, 425, 20, 55, "Folio", "F2", 8, hCenter
                oDoc.WTextBox Posi, 460, 20, 55, "Subtotal", "F2", 8, hCenter
                oDoc.WTextBox Posi, 515, 20, 55, "Tipo", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
        Next
        ' Linea
        Posi = Posi + 15
        oDoc.WTextBox Posi, 400, 20, 120, Format(Total, "###,###,##0.00"), "F3", 10, hRight
        Posi = Posi + 26
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 205, 100, 175, "COMENTARIOS", "F3", 8, hCenter
        Posi = Posi + 20
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Combo4_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo4.Clear
    sBuscar = "SELECT NOMBRE FROM CLIENTE WHERE ASIG='N'  ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo4.AddItem "<TODAS>"
     If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo4.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub combo()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdClien = Item
    Text1.Text = Item.SubItems(1)
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
End Sub
Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdClien1 = Item
    Text20.Text = Item.SubItems(1)
    Check8.Value = 0
    Check9.Value = 0
    Check10.Value = 0
End Sub
Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
End Sub
Private Sub ListView6_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdClieRep = Item
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Frame28.Visible = True
    Else
        Frame28.Visible = False
    End If
End Sub
Private Sub Text1_Change()
    If Text1.Text = "" Then
        IdClien = ""
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT NOMBRE, ID_CLIENTE, DIRECCION FROM CLIENTE WHERE NOMBRE LIKE '%" & Replace(Text1.Text, " ", "%") & "%' OR NOMBRE_COMERCIAL LIKE '%" & Replace(Text1.Text, " ", "%") & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(2) = tRs.Fields("DIRECCION")
                IdClien = IdClien & tRs.Fields("ID_CLIENTE") & ", "
                tRs.MoveNext
            Loop
        End If
        IdClien = Mid(IdClien, 1, Len(IdClien) - 2)
    Else
        IdClien = ""
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command4.Value = True
    End If
End Sub
Private Sub Text20_Change()
    If Text20.Text = "" Then
        IdClien1 = ""
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    If Check17 = 1 Then
        If (Option12.Value = True) Or (Option21.Value = True) Or (Option13.Value = True) Or (Option14.Value = True) Or (Option15.Value = True) Then
            If KeyAscii = 13 Then
                If Check17.Value = 1 Then
                Me.Command7.Value = True
            Else
                Me.Command6.Value = True
            End If
         
           End If
        Else
            MsgBox "SELECCIONE UN JUEGO DE REPARACION!", vbInformation, "SACC"
        End If
    Else
        MsgBox "FALTA  INFORMACION PARA CONTINUAR LA BUSQUEDA", vbInformation, "SACC"
    End If
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
    If Check18.Value = 1 Then
        If Combo5.Text <> "" Then
            If (Option16 = True) Or (Option17 = True) Or (Option18 = True) Or (Option19 = True) Or (Option20 = True) Then
                If KeyAscii = 13 Then
                    Me.Command8.Value = True
                End If
            Else
                MsgBox ("Selecione  Un Juego De Reparacion")
            End If
        Else
            MsgBox ("Selecione  Una Sucursal")
        End If
    End If
    If Check19.Value = 1 Then
        If (Option16 = True) Or (Option17 = True) Or (Option18 = True) Or (Option19 = True) Or (Option20 = True) Then
            If KeyAscii = 13 Then
                Me.Command8.Value = True
            End If
        Else
            MsgBox ("Selecione  Un Juego De Reparacion")
        End If
    End If
End Sub
Private Sub Text20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text20.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT NOMBRE, ID_CLIENTE, DIRECCION FROM CLIENTE WHERE NOMBRE LIKE '%" & Replace(Text20.Text, " ", "%") & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView3.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(2) = tRs.Fields("DIRECCION")
                tRs.MoveNext
            Loop
        End If
    Else
        IdClien1 = ""
    End If
End Sub
Private Sub VentasPorMes()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sWhere As String
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUCURSAL, SUM(CANTIDAD) AS CANTIDAD FROM VsRepVentas "
    If Check20.Value = 1 Then
        sWhere = " FECHA BETWEEN '" & DTPicker11.Value & "' AND '" & DTPicker12.Value & " 23:59:59.997'"
    End If
    If Combo6.Text <> "" Then
        If sWhere <> "" Then
            sWhere = "AND " & sWhere
        End If
        sWhere = sWhere & " TIPO = '" & Combo6.Text & "'"
    End If
    If Combo7.Text <> "" Then
        If sWhere <> "" Then
            sWhere = "AND " & sWhere
        End If
        sWhere = sWhere & " CLASIFICACION = '" & Combo7.Text & "'"
    End If
    If Combo8.Text <> "" Then
        If sWhere <> "" Then
            sWhere = "AND " & sWhere
        End If
        sWhere = sWhere & " MARCA = '" & Combo8.Text & "'"
    End If
    If Combo9.Text <> "" Then
        If sWhere <> "" Then
            sWhere = "AND " & sWhere
        End If
        sWhere = sWhere & " SUCURSAL = '" & Combo9.Text & "'"
    End If
    If sWhere <> "" Then
        sBuscar = sBuscar & "WHERE" & sWhere
    End If
    sBuscar = sBuscar & " GROUP BY ID_PRODUCTO, DESCRIPCION, SUCURSAL"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView11.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
            tRs.MoveNext
       Loop
       StrRep9 = sBuscar
    End If
End Sub
