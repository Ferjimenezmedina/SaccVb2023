VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Ventas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Punto de Venta"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir Remisión"
      Height          =   255
      Left            =   1080
      TabIndex        =   134
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "0.00"
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "0.00"
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "0.00"
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3120
      TabIndex        =   118
      Top             =   6840
      Width           =   975
      Begin VB.Label Label34 
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
         TabIndex        =   119
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Ventas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Ventas.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   4200
      TabIndex        =   67
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ventas"
      TabPicture(0)   =   "Ventas.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Option3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Comandas"
      TabPicture(1)   =   "Ventas.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "lblEstado"
      Tab(1).Control(3)=   "Label30"
      Tab(1).Control(4)=   "Label31"
      Tab(1).Control(5)=   "lvwNuevaComanda"
      Tab(1).Control(6)=   "lvwProductosComanda"
      Tab(1).Control(7)=   "txtId_Cliente"
      Tab(1).Control(8)=   "Frame12"
      Tab(1).Control(9)=   "cmdAceptarComanda"
      Tab(1).Control(10)=   "cmdQuitarComanda"
      Tab(1).Control(11)=   "txtProductoComanda"
      Tab(1).Control(12)=   "Frame13"
      Tab(1).Control(13)=   "Text12"
      Tab(1).Control(14)=   "Text13"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Domicilios"
      TabPicture(2)   =   "Ventas.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "TxtNomCliente"
      Tab(2).Control(2)=   "TxtDomiCleinte"
      Tab(2).Control(3)=   "TxtNotaDomi"
      Tab(2).Control(4)=   "BtnGuardaDomi"
      Tab(2).Control(5)=   "TxtTelefonoDomi"
      Tab(2).Control(6)=   "TxtNoArticulos"
      Tab(2).Control(7)=   "CmbColonia"
      Tab(2).Control(8)=   "BtnNueColonia"
      Tab(2).Control(9)=   "DTPFechaDomi"
      Tab(2).Control(10)=   "Label27"
      Tab(2).Control(11)=   "Label26"
      Tab(2).Control(12)=   "Label3(2)"
      Tab(2).Control(13)=   "Label25"
      Tab(2).Control(14)=   "Label24"
      Tab(2).Control(15)=   "Label23"
      Tab(2).Control(16)=   "Label22"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Asistencia"
      TabPicture(3)   =   "Ventas.frx":2440
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text14"
      Tab(3).Control(1)=   "TxtNomClienAs"
      Tab(3).Control(2)=   "cmdRegis"
      Tab(3).Control(3)=   "Frame8"
      Tab(3).Control(4)=   "TxtDesPiez"
      Tab(3).Control(5)=   "Frame3"
      Tab(3).Control(6)=   "TxtComTec"
      Tab(3).Control(7)=   "TxtTipoArt"
      Tab(3).Control(8)=   "DtPFechAsi"
      Tab(3).Control(9)=   "Label12"
      Tab(3).Control(10)=   "Label28"
      Tab(3).Control(11)=   "Label3(0)"
      Tab(3).Control(12)=   "Label17"
      Tab(3).Control(13)=   "LblMenu"
      Tab(3).Control(14)=   "Label13"
      Tab(3).Control(15)=   "Label16"
      Tab(3).Control(16)=   "Label15"
      Tab(3).ControlCount=   17
      Begin VB.Frame Frame2 
         Caption         =   "Observaciones"
         Height          =   1575
         Left            =   3600
         TabIndex        =   132
         Top             =   4920
         Width           =   3975
         Begin VB.TextBox Text19 
            Height          =   1095
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   133
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   -69600
         MaxLength       =   9
         TabIndex        =   49
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -70800
         MaxLength       =   20
         TabIndex        =   34
         Top             =   6600
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -71640
         MaxLength       =   100
         TabIndex        =   33
         Top             =   6240
         Width           =   4095
      End
      Begin VB.Frame Frame14 
         Caption         =   "Orden de Compra"
         Height          =   1575
         Left            =   3600
         TabIndex        =   111
         Top             =   4920
         Visible         =   0   'False
         Width           =   2175
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   120
            MaxLength       =   100
            TabIndex        =   16
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   840
            MaxLength       =   15
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Solicito"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Numero"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -70680
         TabIndex        =   110
         Top             =   720
         Width           =   3135
         Begin VB.OptionButton opnCodigoComanda 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   2040
            TabIndex        =   27
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton opnClaveComanda 
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton opnDescripcionComanda 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.TextBox txtProductoComanda 
         Height          =   285
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdQuitarComanda 
         Caption         =   "Quitar"
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
         Picture         =   "Ventas.frx":245C
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarComanda 
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
         Height          =   375
         Left            =   -70320
         Picture         =   "Ventas.frx":4E2E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Frame Frame12 
         Caption         =   "Cantidad"
         Height          =   855
         Left            =   -74760
         TabIndex        =   106
         Top             =   6120
         Width           =   3015
         Begin VB.CommandButton cmdAgregarComanda 
            Caption         =   "Agregar"
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
            Left            =   1560
            Picture         =   "Ventas.frx":7800
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCantidadComanda 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtId_Cliente 
         Height          =   285
         Left            =   -71520
         TabIndex        =   105
         Top             =   6960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame11 
         Caption         =   "Cerrar Venta"
         Height          =   1215
         Left            =   240
         TabIndex        =   104
         Top             =   6480
         Width           =   7335
         Begin VB.OptionButton Option13 
            Caption         =   "No Aplica"
            Height          =   195
            Left            =   5400
            TabIndex        =   125
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option12 
            Caption         =   "T. Debito"
            Height          =   195
            Left            =   3120
            TabIndex        =   124
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option11 
            Caption         =   "T. Electrónica"
            Height          =   195
            Left            =   4080
            TabIndex        =   123
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Usar Vale"
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox IdBenta 
            Height          =   285
            Left            =   1560
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton Option8 
            Caption         =   "T. Credito"
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Cheque"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Efectivo"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Guardar"
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
            Left            =   6000
            Picture         =   "Ventas.frx":A1D2
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtNomClienAs 
         Height          =   285
         Left            =   -74760
         TabIndex        =   48
         Top             =   720
         Width           =   6975
      End
      Begin VB.Frame Frame10 
         Caption         =   "Favor de pasar en el horario"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   93
         Top             =   4440
         Width           =   3615
         Begin VB.TextBox TxtHoraAl 
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtHoraDe 
            Height          =   285
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   43
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Y la Hora :"
            Height          =   255
            Left            =   360
            TabIndex        =   95
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "Entra la Hora :"
            Height          =   255
            Left            =   360
            TabIndex        =   94
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtNomCliente 
         Height          =   285
         Left            =   -73440
         MaxLength       =   100
         TabIndex        =   37
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox TxtDomiCleinte 
         Height          =   285
         Left            =   -73440
         MaxLength       =   100
         TabIndex        =   38
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox TxtNotaDomi 
         Height          =   1575
         Left            =   -70920
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton BtnGuardaDomi 
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
         Height          =   375
         Left            =   -69120
         Picture         =   "Ventas.frx":CBA4
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   6720
         Width           =   1335
      End
      Begin VB.TextBox TxtTelefonoDomi 
         Height          =   285
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   41
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox TxtNoArticulos 
         Height          =   285
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   42
         Text            =   "0"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox CmbColonia 
         Height          =   315
         Left            =   -73440
         TabIndex        =   39
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton BtnNueColonia 
         Caption         =   "Nueva Colonia"
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
         Left            =   -70920
         Picture         =   "Ventas.frx":F576
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   6720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegis 
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
         Height          =   375
         Left            =   -70200
         Picture         =   "Ventas.frx":11F48
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         Caption         =   "Articulo"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   84
         Top             =   2640
         Width           =   7095
         Begin VB.TextBox TxtModelo 
            Height          =   285
            Left            =   1800
            TabIndex        =   52
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox TxtMarcaArt 
            Height          =   285
            Left            =   1800
            TabIndex        =   53
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Left            =   1080
            TabIndex        =   86
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            Height          =   195
            Left            =   1080
            TabIndex        =   85
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.TextBox TxtDesPiez 
         DataField       =   "COMENTARIOS_COTIZACION"
         DataSource      =   "Adodc1"
         Height          =   1095
         Left            =   -74760
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   5880
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   -74760
         TabIndex        =   81
         Top             =   1560
         Width           =   7095
         Begin VB.CheckBox Chk1 
            Caption         =   "A domicilio"
            Height          =   255
            Left            =   2400
            TabIndex        =   50
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "Garantia"
            Height          =   255
            Left            =   4200
            TabIndex        =   51
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Sucursal"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   120
            Width           =   6855
         End
         Begin VB.Label LblMenu2 
            Alignment       =   2  'Center
            Caption         =   "Label12"
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
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   6855
         End
      End
      Begin VB.TextBox TxtComTec 
         DataField       =   "COMENTARIOS_TECNICOS"
         DataSource      =   "Adodc1"
         Height          =   1095
         Left            =   -71160
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox TxtTipoArt 
         DataField       =   "TIPO_ARTICULO"
         DataSource      =   "Adodc1"
         Height          =   1125
         Left            =   -74760
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   4320
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Descripción"
         Height          =   195
         Left            =   4440
         TabIndex        =   68
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Quitar"
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
         Left            =   6240
         Picture         =   "Ventas.frx":1491A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Promoción"
         Height          =   1575
         Left            =   5880
         TabIndex        =   76
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Ventas.frx":172EC
            Left            =   120
            List            =   "Ventas.frx":172EE
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   2040
         TabIndex        =   73
         Top             =   4920
         Width           =   1455
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Ventas.frx":172F0
            Left            =   120
            List            =   "Ventas.frx":17303
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Credito"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Disponible"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Dias de credito"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Extraer"
         Height          =   1575
         Left            =   240
         TabIndex        =   72
         Top             =   4920
         Width           =   1695
         Begin VB.OptionButton Option10 
            Caption         =   "Asistencia "
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Comanda"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtComandas 
            Height          =   285
            Left            =   120
            TabIndex        =   117
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtIDCOMANDA 
            Height          =   285
            Left            =   1440
            TabIndex        =   116
            Text            =   "Text14"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Extraer"
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
            Left            =   240
            Picture         =   "Ventas.frx":1731B
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cantidad"
         Height          =   1095
         Left            =   6120
         TabIndex        =   71
         Top             =   3240
         Width           =   1455
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Agregar"
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
            Left            =   120
            Picture         =   "Ventas.frx":19CED
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Clave"
         Height          =   195
         Left            =   4440
         TabIndex        =   70
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   4440
         TabIndex        =   69
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1575
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2778
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DtPFechAsi 
         Height          =   375
         Left            =   -70320
         TabIndex        =   57
         Top             =   6000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50003969
         CurrentDate     =   38678
      End
      Begin MSComCtl2.DTPicker DTPFechaDomi 
         Height          =   375
         Left            =   -69120
         TabIndex        =   40
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50003969
         CurrentDate     =   38833
      End
      Begin MSComctlLib.ListView lvwProductosComanda 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   29
         Top             =   1200
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CLAVE"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCIÓN"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GANANCIA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PRECIO COSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PRECIO"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lvwNuevaComanda 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   30
         Top             =   3600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CLAVE"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCIÓN"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "CANTIDAD"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   -69600
         TabIndex        =   120
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "Telefono"
         Height          =   255
         Left            =   -71640
         TabIndex        =   115
         Top             =   6600
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Nombre del Cliente"
         Height          =   255
         Left            =   -71640
         TabIndex        =   114
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label lblEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   109
         Top             =   4560
         Width           =   5175
      End
      Begin VB.Label Label6 
         Caption         =   "Producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   108
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Comanda"
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   103
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   102
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "Domicilio :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   101
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Colonia :"
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   100
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   -69840
         TabIndex        =   99
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Nota :"
         Height          =   255
         Left            =   -70920
         TabIndex        =   98
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   97
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "# Articulos :"
         Height          =   255
         Left            =   -71640
         TabIndex        =   96
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha en que se debe realizar el domicilio"
         Height          =   195
         Index           =   0
         Left            =   -71040
         TabIndex        =   92
         Top             =   5640
         Width           =   2955
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Atiende : "
         Height          =   195
         Left            =   -74760
         TabIndex        =   91
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label LblMenu 
         Caption         =   "Label13"
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
         Left            =   -74760
         TabIndex        =   90
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Articulo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   89
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descipcion de piezas recibidas"
         Height          =   195
         Left            =   -74760
         TabIndex        =   88
         Top             =   5640
         Width           =   3375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Comentarios para los Tecnicos"
         Height          =   195
         Left            =   -71160
         TabIndex        =   87
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Venta"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   77
         Top             =   1920
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   61
      Top             =   720
      Width           =   3975
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   80
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtID_User 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   79
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblEspere 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ESPERE MIENTRAS SE LLENA LA LISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "0.00"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   360
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Retención"
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
      Left            =   1200
      TabIndex        =   131
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Imp. 2"
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
      Left            =   1440
      TabIndex        =   129
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Imp. 1"
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
      Left            =   1440
      TabIndex        =   127
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   122
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   121
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
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
      Left            =   120
      TabIndex        =   66
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   1440
      TabIndex        =   65
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Left            =   1440
      TabIndex        =   64
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
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
      Left            =   1080
      TabIndex        =   63
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Menu atencion 
      Caption         =   "Opciones"
      Begin VB.Menu reim 
         Caption         =   "Reimprimir"
      End
      Begin VB.Menu r2 
         Caption         =   "-"
      End
      Begin VB.Menu BucCOm 
         Caption         =   "Buscar Comanda"
      End
      Begin VB.Menu SubMenConCobCom 
         Caption         =   "Consultar Cobro de Comanda"
      End
      Begin VB.Menu r3 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenCambiaFormPago 
         Caption         =   "Cambiar Forma de Pago"
      End
      Begin VB.Menu CamClien 
         Caption         =   "Cambiar Cliente a Venta"
      End
   End
   Begin VB.Menu acceso 
      Caption         =   "Accesos"
      Begin VB.Menu AgreClien 
         Caption         =   "Agregar Cliente"
      End
      Begin VB.Menu VerClienteaDetalle 
         Caption         =   "Ver Clientes a Detalle"
      End
      Begin VB.Menu r1 
         Caption         =   "-"
      End
      Begin VB.Menu Coti 
         Caption         =   "Cotizar"
      End
      Begin VB.Menu Factu 
         Caption         =   "Facturar"
      End
   End
End
Attribute VB_Name = "Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim Cont As Integer
Dim xA As ListItem
Dim sqlCliente As String
Dim sqlProducto As String
Dim tRs As ADODB.Recordset
Dim tLi As ListItem
Dim bBCli As Boolean, bBAgr As Boolean
Dim ClvProd As String
Dim DesProd As String
Dim PreProd As String
Dim CLVCLIEN As String
Dim NomClien As String
Dim DesClien As String
Dim DelInd As String
Dim DelDes As String
Dim DelCan As String
Dim DelPre As String
Dim DelIVA As String
Dim DelRET As String
Dim DelIMP1 As String
Dim DelIMP2 As String
Dim BanCnn As Boolean
Dim sqlQuery As String
Dim DOMI As String
Dim GTIA As String
Dim IdClien As String
Dim NoAsTec As String
Dim TelCasa As String
Dim TelTrabajo As String
Dim Direc As String
Dim NoExte As String
Dim NoInte As String
Dim COLONIA As String
Dim DelIndex As Integer
Dim ClasProd As String
Dim IdVentAut As String
Dim RFC As String
' Funcion para ejecutar aplicacion externa al sistema (Factura Electronica) 20/Oct/2011 Armando H Valdez Arras
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'-.-. .... .. -. --. .-   - ..-   -- .- -.. .-. .
Private Sub AgreClien_Click()
    If VarMen.Text1(30).Text = "S" Then
        AltaClien.Show vbModal
    Else
        MsgBox "EL USUARIO NO CUENTA CON LOS PERMISOS NECESARIOS!", vbInformation, "SACC"
    End If
End Sub
Private Sub BucCOm_Click()
    FrmBuscaComanda.Show vbModal
End Sub
Private Sub CamClien_Click()
    FrmCamClienVent.Show vbModal
End Sub
Private Sub Check1_Click()
    If Check2.Value = 1 Then
        Check1.Value = 0
        MsgBox "IMPOSIBLE USAR VALE EN VENTA DE CREDITO", vbCritical, "SACC"
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value = 1 Then
        If Text9.Text = "" Then
            MsgBox "NO SE PEUDE HACER VENTA A CREDITO SI NO SE MARCA UN LIMITE DE CREDITO!", vbInformation, "SACC"
            Check2.Value = 0
        Else
            If CDbl(Text9.Text) < CDbl(Text5.Text) Then
                MsgBox "NO SE PEUDE HACER VENTA A CREDITO SI EL LIMITE ES MENOR AL MONTO DE LA VENTA!", vbInformation, "SACC"
                Check2.Value = 0
            Else
                Check1.Value = 0
            End If
        End If
    End If
End Sub
Private Sub CmbColonia_LostFocus()
    CmbColonia.BackColor = &H80000005
End Sub
Private Sub LlenaCombo()
On Error GoTo ManejaError
    Me.Combo1.Clear
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    sBus = "SELECT TIPO FROM PROMOCION ORDER BY TIPO"
    Set tRs = cnn.Execute(sBus)
    Combo1.AddItem "<NINGUNA>"
    Combo1.AddItem "LICITACIÓN"
    If Not (tRs.EOF And tRs.BOF) Then
        Combo1.AddItem "PROMOCION"
        Combo1.Text = "PROMOCION"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HFFE1E1
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &H80000005
End Sub
Private Sub FunRemision()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & IdVentAut & ""
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\Remision.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
        End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 50, "CANTIDAD", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 90, "CLAVE", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 145, 20, 280, "DESCRIPCION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 425, 20, 60, "PRESENTACION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 485, 20, 50, "PRECIO UNITARIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 535, 20, 50, "TOTAL", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20

        ' DETALLE
        sBuscar = "SELECT VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, ALMACEN3.PRESENTACION, VENTAS_DETALLE.PRECIO_VENTA, VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD AS TOTAL FROM ALMACEN3 INNER JOIN VENTAS_DETALLE ON ALMACEN3.ID_PRODUCTO = VENTAS_DETALLE.ID_PRODUCTO WHERE VENTAS_DETALLE.ID_VENTA = " & tRs1.Fields("ID_VENTA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 5, 15, 50, Format(tRs3.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 7, hCenter, , , 1, vbBlack
                
                'oDoc.WTextBox Posi, 55, 15, 90, " " & tRs3.Fields("ID_PRODUCTO"), "F3", 7, hLeft, , , 1, vbBlack
                oDoc.WTextBox Posi, 55, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 1, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 85, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 4, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 115, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 7, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 145, 15, 280, " " & tRs3.Fields("DESCRIPCION"), "F3", 7, hLeft, , , 1, vbBlack
                'oDoc.WTextBox Posi, 425, 15, 60, tRs3.Fields("PRESENTACION"), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 425, 15, 60, Mid(tRs3.Fields("ID_PRODUCTO"), 11, 7), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 485, 15, 50, Format(CDbl(tRs3.Fields("PRECIO_VENTA")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                oDoc.WTextBox Posi, 535, 15, 50, Format(CDbl(tRs3.Fields("TOTAL")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 600 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & tRs1.Fields("ID_VENTA")
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        oDoc.WImage 70, 40, 43, 161, "Logo"
                        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
                        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
                
                        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
                        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
                        Posi = Posi + 15
                        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
                        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
                            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
                        End If
                        Posi = 210
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        'oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Impuesto 1:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Impuesto 2:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Retencion:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO1")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("IMPUESTO1"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO2")) Then oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("IMPUESTO2"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("RETENCION")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("RETENCION"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IVA")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TOTAL")) Then oDoc.WTextBox 720, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 720, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        'If tRs1.Fields("CONFIRMADA") = "E" Then
        '    oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        'End If
        'oDoc.WTextBox 620, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 620, 20, 100, 275, "OBSERVACIONES:", "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 640, 60, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 700, 20, 100, 275, "RESPONSABLE : ", "F3", 8, hLeft
                oDoc.WTextBox 720, 20, 100, 275, tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hLeft
            End If
        End If
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se enc ontro la orden de compra solicitada, puede ser que este cancelda o aun no se genere el folio", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    On Error GoTo ManejaError
    Command2.Enabled = False
    Dim tRs8 As ADODB.Recordset
    Dim tRs7 As ADODB.Recordset
    Dim Item As MSComctlLib.ListItem
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    Dim NRegistros As Integer
    Dim Con As Integer
    Dim POSY As Integer
    Dim FechaVence As String
    Dim sBuscar As String
    Dim P_COSTO As String
    Dim cant As String
    Dim CanProd As String
    Dim P_ven As String
    Dim Ganan As String
    Dim IdCta As String
    Dim TPago As String
    Dim TotDeuda As Double
    Dim Vale As Double
    Dim MosLey As String
    Dim ID_VALE As Integer
    Dim continuar As Boolean
    Dim AbonoClien As String
    Dim sExibicion As String
    Dim FormaPagoSAT As String
    If Check3.Value = 1 Then
        SaveSetting "APTONER", "ConfigSACC", "Remision", "S"
    Else
        SaveSetting "APTONER", "ConfigSACC", "Remision", "N"
    End If
    If Check2.Value = 1 Then
        If CDbl(Text9.Text) < CDbl(Text5.Text) Then
            MsgBox "EL MONTO DE LA VENTA SUPERO EL LIMITE DE CREDITO DISPONIBLE!", vbInformation, "SACC"
            Check2.Value = 0
            Exit Sub
        Else
            Check1.Value = 0
        End If
    End If
    If Text7.Text <> "" Then
        sBuscar = "SELECT NO_PEDIDO FROM PED_CLIEN WHERE NO_ORDEN = '" & Text7.Text & "' AND ID_CLIENTE  = " & Me.txtId_Cliente.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            MsgBox "Este numero de orden de compra esta asignado a la venta programada numero " & tRs.Fields("NO_PEDIDO"), vbExclamation, "SACC"
            Exit Sub
        End If
     End If
    If Val(Replace(Text9.Text, ",", "")) > CDbl(Text5.Text) Then
        If MsgBox("EL CLIENTE MANEJA CREDITO" & Chr(13) & "DESEA QUE ESTA VENTA SEA A CREDITO?", vbYesNo, "SACC") = vbYes Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
    Else
        If Check2.Value = 1 Then
            If MsgBox("EL CLIENTE NO CUENTA CON SUFICIENTE CREDITO, DESEA HACER LA VENTA DE CONTADO?", vbYesNo, "SACC") = vbYes Then
                Check2.Value = 0
            Else
                Exit Sub
            End If
        End If
    End If
    If Option6.Value = True Then
        TPago = "C"
        FormaPagoSAT = "001"
    Else
        If Option7.Value = True Then
            TPago = "H"
            FormaPagoSAT = "002"
        Else
            If Option8.Value = True Then
                TPago = "T"
                FormaPagoSAT = "004"
            Else
                If Option11.Value = True Then
                    TPago = "E"
                    FormaPagoSAT = "003"
                Else
                    If Option13.Value = True Then
                        TPago = "N"
                        FormaPagoSAT = "099"
                    Else
                        TPago = "D"
                        FormaPagoSAT = "028"
                    End If
                End If
            End If
        End If
    End If
    If Option6.Value = False And Option11.Value = False And Option7.Value = False And Option8.Value = False And Option12.Value = False And Option13.Value = False Then
        MsgBox "DEBE MARCAR UNA FORMA DE PAGO!", vbExclamation, "SACC"
        Exit Sub
    End If
    continuar = True
    Vale = 0
    If Check1.Value = 1 Then
        sBuscar = "SELECT * FROM VALE_CAJA WHERE APLICADO = 'N' AND ID_VALE = " & InputBox("INTRODUSCA FOLIO DEL VALE", "SACC")
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            If MsgBox("EL VALE DE CAJA NO EXISTE O YA FUE APLICADO" & Chr(13) & "DESEA CONTINUAR SIN USAR NINGUN VALE?", vbYesNo, "SACC") = vbNo Then
                continuar = False
            End If
        Else
            Vale = tRs.Fields("IMPORTE")
            ID_VALE = tRs.Fields("ID_VALE")
        End If
    End If
    If continuar Then
        If Check2.Value = 0 Then
            Text3.Text = Replace(Text3.Text, ",", "")
            Text4.Text = Replace(Text4.Text, ",", "")
            Text5.Text = Replace(Text5.Text, ",", "")
            If InStr(1, Text5.Text, ".") > 0 Then Text5.Text = Right(Text5.Text, InStr(1, Text5.Text, ".") + 2)
            If InStr(1, Text3.Text, ".") > 0 Then Text3.Text = Right(Text3.Text, InStr(1, Text3.Text, ".") + 2)
            If InStr(1, Text4.Text, ".") > 0 Then Text4.Text = Right(Text4.Text, InStr(1, Text4.Text, ".") + 2)
            If DesClien = "" Then DesClien = "0.00"
            DesClien = Replace(DesClien, ",", "")
            sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, DESCUENTO, FECHA, TOTAL, SUCURSAL, ID_USUARIO, IVA, SUBTOTAL, TIPO_PAGO, UNA_EXIBICION, NOOC, COMENTARIO, FORMA_PAGO, IMPUESTO1, IMPUESTO2, RETENCION, FormaPagoSAT) VALUES (" & CLVCLIEN & ", '" & NomClien & "', " & DesClien & ",  SYSDATETIME(), " & Text5.Text & ", '" & VarMen.Text4(0).Text & "', '" & VarMen.Text1(0).Text & "', " & Text4.Text & ", " & Text3.Text & ", '" & TPago & "', 'S', '" & Text7.Text & "', '" & Text19.Text & "', 'PAGO EN UNA SOLA EXHIBICION', " & Text15.Text & ", " & Text16.Text & ", " & Text17.Text & ", '" & FormaPagoSAT & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_VENTA, UNA_EXIBICION FROM VENTAS WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "' ORDER BY ID_VENTA DESC"
            Set tRs = cnn.Execute(sBuscar)
            sExibicion = tRs.Fields("UNA_EXIBICION")
            IdVentAut = tRs.Fields("ID_VENTA")
            NumeroRegistros = ListView3.ListItems.Count
            For Conta = 1 To NumeroRegistros
                cant = Replace(ListView3.ListItems.Item(Conta).SubItems(2), ",", "")
                P_ven = Replace(ListView3.ListItems.Item(Conta).SubItems(3), ",", "")
                If Val(cant) > 0 Then
                    P_ven = Format(Val(P_ven) / Val(cant), "0.00")
                    P_ven = Replace(P_ven, ",", "")
                Else
                    P_ven = "0.00"
                End If
                sBuscar = "INSERT INTO VENTAS_DETALLE (ID_PRODUCTO, DESCRIPCION, PRECIO_VENTA, CANTIDAD, ID_VENTA, NO_COM_AT, IMPORTE, IVA, IMPUESTO1, IMPUESTO2, RETENCION) VALUES ('" & ListView3.ListItems.Item(Conta) & "', '" & ListView3.ListItems.Item(Conta).SubItems(1) & "', " & P_ven & ", " & cant & ", " & IdVentAut & ", '" & ListView3.ListItems.Item(Conta).SubItems(4) & "', " & Format(Val(P_ven) * Val(cant), "0.00") & ", '" & ListView3.ListItems(Conta).SubItems(5) & "', '" & ListView3.ListItems(Conta).SubItems(6) & "', '" & ListView3.ListItems(Conta).SubItems(7) & "', '" & ListView3.ListItems(Conta).SubItems(8) & "');"
                cnn.Execute (sBuscar)
            Next Conta
            
            sBuscar = "SELECT * FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "' AND ELIMINADO = 'N'"
            Set tRs8 = cnn.Execute(sBuscar)
            '********************************IMPRIMIR TICKET********************************************
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA"))) / 2
            Printer.Print tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA")
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP"))) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP")
            Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
            Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
            Printer.Print "TELEFONO SUCURSAL : " & tRs8.Fields("TELEFONO")
            Printer.Print "No. DE VENTA : " & IdVentAut
            If Option6.Value = True Then
                Printer.Print "FORMA DE PAGO : EFECTIVO"
            Else
                If Option7.Value = True Then
                    Printer.Print "FORMA DE PAGO : CHEQUE"
                Else
                    If Option8.Value = True Then
                        Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                    Else
                        If Option11.Value = True Then
                            Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                        Else
                            Printer.Print "FORMA DE PAGO : NO INDICADO"
                        End If
                    End If
                End If
            End If
            Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " "; VarMen.Text1(2).Text
            Printer.Print "CLIENTE : " & Text1.Text
            If sExibicion = "N" Then
                Printer.Print "VENTA A CREDITO"
            Else
                Printer.Print "VENTA DE CONTADO"
            End If
            Printer.Print "--------------------------------------------------------------------------------"
            Printer.Print "                          NOTA DE FACTURA"
            Printer.Print "--------------------------------------------------------------------------------"
            NRegistros = ListView3.ListItems.Count
            POSY = 2900
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print "Cant."
            Printer.CurrentY = POSY
            Printer.CurrentX = 3000
            Printer.Print "Precio unitario"
            For Con = 1 To NRegistros
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print Mid(ListView3.ListItems(Con).Text, 1, 25)
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print ListView3.ListItems(Con).SubItems(2)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                If CDbl(ListView3.ListItems(Con).SubItems(2)) > 0 Then
                    Printer.Print Format(CDbl(ListView3.ListItems(Con).SubItems(3)) / CDbl(ListView3.ListItems(Con).SubItems(2)), "###,###,##0.00")
                Else
                    Printer.Print Format(CDbl(ListView3.ListItems(Con).SubItems(3)), "###,###,##0.00")
                End If
            Next Con
            Printer.Print ""
            Printer.Print "SUBTOTAL : " & Format(Text3.Text, "###,###,##0.00")
            Printer.Print "IVA              : " & Format(Text4.Text, "###,###,##0.00")
            If Text17.Text <> "0.00" Then
                        Printer.Print "RETENCION: " & Format(Text17.Text, "###,###,##0.00")
                    End If
                    If Text15.Text <> "0.00" Then
                        Printer.Print "IMPUESTO1: " & Format(Text15.Text, "###,###,##0.00")
                    End If
                    If Text16.Text <> "0.00" Then
                        Printer.Print "IMPUESTO2: " & Format(Text16.Text, "###,###,##0.00")
                    End If
            Printer.Print "TOTAL        : " & Format(Text5.Text, "###,###,##0.00")
            If Vale > 0 Then
                Printer.Print "VALOR DEL VALE:" & Vale
                Printer.Print "TOTAL A PAGAR:" & Val(Replace(Text5.Text, ",", "")) - Vale
            End If
            Printer.Print ""
            Printer.Print "--------------------------------------------------------------------------------"
            Printer.Print "               GRACIAS POR SU COMPRA"
            Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
            Printer.Print "     DESPUES DE HABER EFECTUADO SU "
            Printer.Print "                                COMPRA"
            Printer.Print "--------------------------------------------------------------------------------"
            Printer.EndDoc
            If Vale > 0 Then
                sBuscar = "UPDATE VALE_CAJA SET APLICADO = 'S' WHERE ID_VALE = " & ID_VALE
                cnn.Execute (sBuscar)
            End If
            If txtComandas.Text <> "" Then
                sBuscar = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'I' WHERE ID_COMANDA IN ( " & txtComandas.Text & " ) AND ESTADO_ACTUAL IN ('L','N') "
                cnn.Execute (sBuscar)
                sBuscar = "SELECT COUNT(*) AS CONTA FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA IN ( " & txtComandas.Text & " )"
                Set tRs = cnn.Execute(sBuscar)
                If tRs.Fields("CONTA") <> 0 Then
                    sBuscar = "UPDATE COMANDAS_2 SET ESTADO_ACTUAL = 'I' WHERE ID_COMANDA IN ( " & txtComandas.Text & " ) AND ESTADO_ACTUAL IN ('L','N')"
                    cnn.Execute (sBuscar)
                End If
            End If
        Else
            If InStr(1, Text3.Text, ".") > 0 Then Text3.Text = Right(Text3.Text, InStr(1, Text3.Text, ".") + 2)
            If InStr(1, Text4.Text, ".") > 0 Then Text4.Text = Right(Text4.Text, InStr(1, Text4.Text, ".") + 2)
            sBuscar = "SELECT SUM(TOTAL_COMPRA) AS TOTAL FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If IsNull(tRs.Fields("TOTAL")) Then
                    TotDeuda = 0
                Else
                    TotDeuda = tRs.Fields("TOTAL")
                End If
            Else
                TotDeuda = 0
            End If
            TotDeuda = TotDeuda + Val(Replace(Text5.Text, ",", ""))
            sBuscar = "SELECT SUM(CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If IsNull(tRs.Fields("TOTAL")) Then
                    AbonoClien = 0
                Else
                    AbonoClien = tRs.Fields("TOTAL")
                End If
            Else
                AbonoClien = 0
            End If
            TotDeuda = TotDeuda - CDbl(Replace(AbonoClien, ",", ""))
            If CDbl(Text5.Text) <= CDbl(Replace(Text9.Text, ",", "")) Then
                If Combo2.Text <> "" Then
                    '********************************* PARA CREDITO ********************************
                    'CAMBIAR TODO EL PROCEDIMIENTO, ASEGURARSE QUE INSERTE LA DEUDA
                    NRegistros = ListView3.ListItems.Count
                    FechaVence = Format(Date + CDbl(Combo2.Text), "dd/mm/yyyy")
                    Text3.Text = Replace(Text3.Text, ",", "")
                    Text4.Text = Replace(Text4.Text, ",", "")
                    Text5.Text = Replace(Text5.Text, ",", "")
                    sBuscar = "SELECT LEYENDAS FROM CLIENTE WHERE ID_CLIENTE = " & CLVCLIEN
                    Set tRs2 = cnn.Execute(sBuscar)
                    If Not (tRs2.EOF And tRs2.BOF) Then
                        If tRs2.Fields("LEYENDAS") = "N" Then
                            MosLey = ""
                        Else
                            MosLey = "PAGO EN PARCIALIDADES"
                        End If
                    End If
                    sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, SUCURSAL, ID_USUARIO, FECHA, TOTAL, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, TIPO_PAGO, UNA_EXIBICION, IVA, SUBTOTAL, FORMA_PAGO, IMPUESTO1, IMPUESTO2, RETENCION) VALUES (" & CLVCLIEN & ", '" & NomClien & "', '" & VarMen.Text4(0).Text & "', '" & VarMen.Text1(0).Text & "',  SYSDATETIME(), " & Text5.Text & ", '" & Combo2.Text & "', '" & FechaVence & "', " & DesClien & ", '" & TPago & "', 'N', " & Text4.Text & ", " & Text3.Text & ", '" & MosLey & "', " & Text15.Text & ", " & Text16.Text & ", " & Text17.Text & ");"
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "' ORDER BY ID_VENTA DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    IdVentAut = tRs.Fields("ID_VENTA")
                    sBuscar = "INSERT INTO CUENTAS (PAGADA, ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, SUCURSAL, TOTAL_COMPRA, DEUDA, ID_VENTA) VALUES ( 'N', " & CLVCLIEN & ", '" & VarMen.Text1(0).Text & "',  SYSDATETIME(), " & Combo2.Text & ", '" & FechaVence & "', " & DesClien & ", '" & VarMen.Text4(0).Text & "', " & Text5.Text & ", " & Text5.Text & ", " & IdVentAut & ");"
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT TOP 1 ID_CUENTA FROM CUENTAS ORDER BY ID_CUENTA DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    IdCta = tRs.Fields("ID_CUENTA")
                    IdBenta.Text = IdVentAut
                    sBuscar = "INSERT INTO CUENTA_VENTA (ID_VENTA, ID_CUENTA) VALUES (" & IdVentAut & ", " & IdCta & ");"
                    cnn.Execute (sBuscar)
                    NRegistros = ListView3.ListItems.Count
                    For Conta = 1 To NRegistros
                        CanProd = Replace(ListView3.ListItems(Conta).SubItems(2), ",", "")
                        P_ven = Format((CDbl(ListView3.ListItems(Conta).SubItems(3)) / CDbl(ListView3.ListItems(Conta).SubItems(2))), "0.00")
                        P_ven = Replace(P_ven, ",", "")
                        sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, CANTIDAD, ID_PRODUCTO, PRECIO_VENTA) VALUES (" & IdCta & ", " & CanProd & ", '" & ListView3.ListItems(Conta).Text & "', " & P_ven & ");"
                        cnn.Execute (sBuscar)
                        sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView3.ListItems(Conta).Text & "'"
                        Set tRs = cnn.Execute(sBuscar)
                        P_COSTO = tRs.Fields("PRECIO_COSTO")
                        Ganan = tRs.Fields("GANANCIA")
                        P_COSTO = Replace(P_COSTO, ",", "")
                        Ganan = Replace(Ganan, ",", "")
                        sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, CANTIDAD, ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA, PRECIO_VENTA, IMPORTE, IVA, IMPUESTO1, IMPUESTO2, RETENCION) VALUES (" & IdVentAut & ", " & CanProd & ", '" & ListView3.ListItems(Conta).Text & "', '" & ListView3.ListItems(Conta).SubItems(1) & "', " & P_COSTO & ", " & Ganan & ", " & P_ven & ", " & CDbl(P_ven) * CDbl(CanProd) & ", '" & ListView3.ListItems(Conta).SubItems(5) & "', '" & ListView3.ListItems(Conta).SubItems(6) & "', '" & ListView3.ListItems(Conta).SubItems(7) & "', '" & ListView3.ListItems(Conta).SubItems(8) & "');"
                        cnn.Execute (sBuscar)
                    Next Conta
                    '********************************IMPRIMIR TICKET********************************************
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                    Printer.Print VarMen.Text5(0).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
                    Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
                    Printer.Print "No. DE VENTA : " & IdVentAut
                    Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " "; VarMen.Text1(2).Text
                    If Option6.Value = True Then
                        Printer.Print "FORMA DE PAGO : EFECTIVO"
                    Else
                        If Option7.Value = True Then
                            Printer.Print "FORMA DE PAGO : CHEQUE"
                        Else
                            If Option8.Value = True Then
                                Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                            Else
                                If Option11.Value = True Then
                                    Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                                Else
                                    Printer.Print "FORMA DE PAGO : NO INDICADO"
                                End If
                            End If
                        End If
                    End If
                    Printer.Print "CLIENTE : " & Text1.Text
                    Printer.Print "VENTA A CREDITO"
                    Printer.Print "--------------------------------------------------------------------------------"
                    Printer.Print "                          NOTA DE FACTURA"
                    Printer.Print "--------------------------------------------------------------------------------"
                    NRegistros = ListView3.ListItems.Count
                    POSY = 2600
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 1300
                    Printer.Print "Cant."
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 3000
                    Printer.Print "Precio unitario"
                    For Con = 1 To NRegistros
                        POSY = POSY + 200
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 100
                        Printer.Print ListView3.ListItems(Con).Text
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 1900
                        Printer.Print ListView3.ListItems(Con).SubItems(2)
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 2900
                        Printer.Print CDbl(ListView3.ListItems(Con).SubItems(3)) / CDbl(ListView3.ListItems(Con).SubItems(2))
                    Next Con
                    Printer.Print ""
                    Printer.Print "SUBTOTAL : " & Format(Text3.Text, "###,###,##0.00")
                    Printer.Print "IVA              : " & Format(Text4.Text, "###,###,##0.00")
                    If Text17.Text <> "0.00" Then
                        Printer.Print "RETENCION: " & Format(Text17.Text, "###,###,##0.00")
                    End If
                    If Text15.Text <> "0.00" Then
                        Printer.Print "IMPUESTO1: " & Format(Text15.Text, "###,###,##0.00")
                    End If
                    If Text16.Text <> "0.00" Then
                        Printer.Print "IMPUESTO2: " & Format(Text16.Text, "###,###,##0.00")
                    End If
                    Printer.Print "TOTAL        : " & Format(Text5.Text, "###,###,##0.00")
                    Printer.Print ""
                    Printer.Print "--------------------------------------------------------------------------------"
                    Printer.Print "               GRACIAS POR SU COMPRA"
                    Printer.Print "           PRODUCTO 100% GARANTIZADO"
                    Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
                    Printer.Print "     DESPUES DE HABER EFECTUADO SU "
                    Printer.Print "                                COMPRA"
                    Printer.Print "SIN SU TICKET NO SERA VALIDA LA GARANTIA."
                    Printer.Print "                APLICA RESTRICCIONES"
                    Printer.Print "--------------------------------------------------------------------------------"
                    Printer.EndDoc
                    If txtComandas.Text <> "" Then
                        sBuscar = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'I' WHERE ID_COMANDA IN ( " & txtComandas.Text & " ) AND ESTADO_ACTUAL IN ('L','N') "
                        Set tRs = cnn.Execute(sBuscar)
                        sBuscar = "SELECT COUNT(*) CONTA FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA IN ( " & txtComandas.Text & " ) AND ESTADO_ACTUAL IN ('L','N') "
                        Set tRs = cnn.Execute(sBuscar)
                        If tRs.Fields("CONTA") <> 0 Then
                            sBuscar = "UPDATE COMANDAS_2 SET ESTADO_ACTUAL = 'I' WHERE ID_COMANDA IN ( " & txtComandas.Text & " ) AND ESTADO_ACTUAL IN ('L','N')"
                            Set tRs = cnn.Execute(sBuscar)
                        End If
                    End If
                Else
                    MsgBox "DEBE SELECCIONAR LOS DIAS DE CREDITO!", vbInformation, "SACC"
                End If
            Else
                MsgBox "EL TOTAL ES MAYOR AL CREDITO DISPONIBLE DEL CLIENTE!", vbInformation, "SACC"
                If MsgBox("DESEA CANCELAR LA VENTA", vbYesNo, "SACC") = vbYes Then
                    NRegistros = ListView3.ListItems.Count
                    For Conta = 1 To NRegistros
                        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView3.ListItems(Conta) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                        Set tRs = cnn.Execute(sBuscar)
                        If Not (tRs.EOF And tRs.BOF) Then
                            cant = CDbl(tRs.Fields("CANTIDAD")) + CDbl(ListView3.ListItems(Conta).SubItems(2))
                            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & ListView3.ListItems(Conta) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                        Else
                            sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & ListView3.ListItems(Conta).SubItems(2) & ", '" & ListView3.ListItems(Conta) & "', '" & VarMen.Text4(0).Text & "');"
                        End If
                        cnn.Execute (sBuscar)
                    Next Conta
                Else
                    Exit Sub
                End If
            End If
        End If
        If Check3.Value = 1 Then
            FunRemision
        End If
    Else
        NRegistros = ListView3.ListItems.Count
        For Conta = 1 To NRegistros
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView3.ListItems(Conta) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                cant = CDbl(tRs.Fields("CANTIDAD")) + Val(Replace(ListView3.ListItems(Conta).SubItems(2), ",", ""))
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & ListView3.ListItems(Conta) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Else
                sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & ListView3.ListItems(Conta).SubItems(2) & ", '" & ListView3.ListItems(Conta) & "', '" & VarMen.Text4(0).Text & "');"
            End If
            cnn.Execute (sBuscar)
        Next Conta
    End If
    Finalizar
    sBuscar = "SELECT VENTAS_DETALLE.ID_VENTA, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.IMPORTE * ALMACEN3.IVA AS IVABIEN FROM VENTAS_DETALLE INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (VENTAS_DETALLE.IVA = 0) AND (ALMACEN3.IVA > 0) AND (VENTAS_DETALLE.IMPORTE > 0)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE VENTAS_DETALLE SET IVA = " & Format(tRs.Fields("IVABIEN"), "0.00") & " WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'COREGIR VENTAS CON IMPUESTOS NULL
    sBuscar = "UPDATE VENTAS SET IMPUESTO1 = 0, IMPUESTO2 = 0, RETENCION = 0 WHERE (ID_VENTA IN (SELECT ID_VENTA FROM VENTAS AS VENTAS_1 WHERE (IMPUESTO1 IS NULL) OR (IMPUESTO2 IS NULL) OR (RETENCION IS NULL)))"
    cnn.Execute (sBuscar)
    Command2.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If DelInd <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tRs1 As ADODB.Recordset
        Dim sBuscar As String
        Dim cant As String
        If ListView3.ListItems.Item(DelIndex) <> "" Then
            If Mid(DelInd, Len(DelInd) - 2, Len(DelInd)) <> "REC" And Mid(DelInd, Len(DelInd) - 2, Len(DelInd)) <> "REM" Then
                ' ---------------------------- PRODUCTOS EQUIVALENTES ----------------------------
                ' ------------------------------ 25/05/2021 H VALDEZ -----------------------------
                cant = CDbl(DelCan)
                sBuscar = "SELECT JUEGO_REPARACION.ID_PRODUCTO, JUEGO_REPARACION.CANTIDAD FROM ALMACEN3 INNER JOIN JUEGO_REPARACION ON ALMACEN3.ID_PRODUCTO = JUEGO_REPARACION.ID_REPARACION WHERE (ALMACEN3.TIPO = 'EQUIVALE') AND  ALMACEN3.ID_PRODUCTO = '" & DelInd & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    DelInd = tRs1.Fields("ID_PRODUCTO")
                    DelCan = CDbl(tRs1.Fields("CANTIDAD")) * cant
                End If
                ' ++++++++++++++++++++++++++++ PRODUCTOS EQUIVALENTES ++++++++++++++++++++++++++++
                sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & DelInd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "' AND CANTIDAD > 0"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    cant = CDbl(tRs.Fields("CANTIDAD")) + Val(Replace(DelCan, ",", ""))
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & Replace(cant, ",", "") & " WHERE ID_PRODUCTO = '" & DelInd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                    cnn.Execute (sBuscar)
                Else
                    cant = Val(Replace(DelCan, ",", ""))
                    sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & Replace(cant, ",", "") & ", '" & DelInd & "','" & VarMen.Text4(0).Text & "');"
                    cnn.Execute (sBuscar)
                End If
            Else
                MsgBox "El producto retirado no sera regresado como existencia!", vbInformation, "SACC"
            End If
            Command3.Enabled = False
            Text3.Text = Format(Val(Replace(Text3.Text, ",", "")) - Val(Replace(ListView3.ListItems.Item(DelIndex).SubItems(3), ",", "")), "0.00")
            ListView3.ListItems.Remove DelIndex
            DelIndex = 0
            DelInd = ""
            If ListView3.ListItems.Count = 0 Then
                Text1.Enabled = True
                ListView1.Enabled = True
            End If
        Else
            MsgBox "No ha seleccionado un producto para retirar!", vbExclamation, "SACC"
        End If
    End If
    If Text3.Text = "0.00" Or Val(Replace(Text3.Text, ",", "")) < 0 Then
        Text3.Text = "0.00"
        Text4.Text = "0.00"
        Text17.Text = "0.00"
        Text15.Text = "0.00"
        Text16.Text = "0.00"
        Text5.Text = "0.00"
    Else
        If RFC = "XEXX010101000" Then
            Text4.Text = "0.00"
            Text17.Text = "0.00"
            Text15.Text = "0.00"
            Text16.Text = "0.00"
            Text5.Text = Format(CDbl(Text3.Text), "0.00")
        Else
            Text4.Text = Format(CDbl(Text4.Text) - CDbl(DelIVA), "0.00")
            Text17.Text = Format(CDbl(Text17.Text) - CDbl(DelRET), "0.00")
            Text15.Text = Format(CDbl(Text15.Text) - CDbl(DelIMP1), "0.00")
            Text16.Text = Format(CDbl(Text16.Text) - CDbl(DelIMP2), "0.00")
            Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) - CDbl(Text17.Text) + CDbl(Text15.Text) + CDbl(Text16.Text), "0.00")
        End If
    End If
    DelIndex = 0
    Command4.Visible = True
    Command2.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
    If Text1.Text <> "" And ListView1.ListItems.Count > 0 Then
        If Text8.Text <> "" Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim tRs1 As ADODB.Recordset
            Dim Canti As Double
            'EVITAR DUPLICIDAD DE VENTAS PROGRAMADAS
            'If CLVCLIEN <> "" Then
            '    sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, PED_CLIEN_DETALLE.CANTIDAD_PEDIDA FROM PED_CLIEN_DETALLE, PED_CLIEN WHERE PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO AND PED_CLIEN_DETALLE.ID_PRODUCTO = '" & ClvProd & "' AND PED_CLIEN.ID_CLIENTE = " & CLVCLIEN & " AND  PED_CLIEN.ESTADO IN ('C', 'I')"
            '    Set tRs = cnn.Execute(sBuscar)
            '    If Not (tRs.EOF And tRs.BOF) Then
            '        If CDbl(tRs.Fields("CANTIDAD_PEDIDA")) = CDbl(Text8.Text) Then
            '            MsgBox "EL PRODUCTO YA TIENE UNA VENTA PROGRAMADA CON ESTE CLIENTE POR LA MISMA CANTIDAD DEL PRODUCTO EN EL FOLIO " & tRs.Fields("NO_PEDIDO") & " CIERRE LA VENTA PROGRAMADA POR ESTE PRODUCTO PARA PODER VENDER", vbExclamation, "SACC"
            '            Exit Sub
            '        Else
            '            If MsgBox("EL PRODUCTO YA TIENE UNA VENTA PROGRAMADA CON ESTE CLIENTE POR " & tRs.Fields("CANTIDAD_PEDIDA") & " UNIDADES EN EL FOLIO " & tRs.Fields("NO_PEDIDO") & " ESTA SEGURO QUE QUIERE AGREGAR EL PRODUCTO A ESTA NOTA?", vbYesNo, "SACC") = vbNo Then
            '                Exit Sub
            '            End If
            '        End If
            '    End If
            'End If
            ' ---------------------------- PRODUCTOS EQUIVALENTES ----------------------------
            ' ------------------------------ 25/05/2021 H VALDEZ -----------------------------
            Canti = CDbl(Text8.Text)
            sBuscar = "SELECT JUEGO_REPARACION.ID_PRODUCTO, JUEGO_REPARACION.CANTIDAD FROM ALMACEN3 INNER JOIN JUEGO_REPARACION ON ALMACEN3.ID_PRODUCTO = JUEGO_REPARACION.ID_REPARACION WHERE (ALMACEN3.TIPO = 'EQUIVALE') AND  ALMACEN3.ID_PRODUCTO = '" & ClvProd & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Canti = tRs.Fields("CANTIDAD") * Canti
                sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND CANTIDAD >= " & Canti & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs1.Fields("CANTIDAD")) - Canti & " WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    Agregar2
                    Text1.Enabled = False
                    ListView1.Enabled = False
                    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
                        Me.Command2.Enabled = True
                    Else
                        Me.Command2.Enabled = False
                    End If
                Else
                    MsgBox "LA CANTIDAD ES MAYOR A LA EXISTENCIA!", vbInformation, "SACC"
                End If
            Else
                sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ClvProd & "' AND CANTIDAD >= " & Canti & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - Canti & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    Agregar2
                    Text1.Enabled = False
                    ListView1.Enabled = False
                    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
                        Me.Command2.Enabled = True
                    Else
                        Me.Command2.Enabled = False
                    End If
                Else
                    MsgBox "LA CANTIDAD ES MAYOR A LA EXISTENCIA!", vbInformation, "SACC"
                End If
            End If
            ' ++++++++++++++++++++++++++++ PRODUCTOS EQUIVALENTES ++++++++++++++++++++++++++++
            'sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ClvProd & "' AND CANTIDAD >= " & Canti & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            'Set tRs = cnn.Execute(sBuscar)
            'If Not (tRs.EOF And tRs.BOF) Then
            '    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - Canti & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            '    Set tRs = cnn.Execute(sBuscar)
            '    Agregar2
            '    Text1.Enabled = False
            '    ListView1.Enabled = False
            '    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
            '        Me.Command2.Enabled = True
            '    Else
            '        Me.Command2.Enabled = False
            '    End If
            'Else
            '    MsgBox "LA CANTIDAD ES MAYOR A LA EXISTENCIA!", vbInformation, "SACC"
            'End If
        End If
    Else
        MsgBox "DEBE SELECCIONAR UN CLIENTE PRIMERO", vbCritical, "SACC"
    End If
    If Check2.Value = 1 Then
        If CDbl(Text9.Text) < CDbl(Text5.Text) Then
            MsgBox "EL MONTO DE LA VENTA SUPERO EL LIMITE DE CREDITO DISPONIBLE, LA VENTA SE HARA DE CONTADO SI CONTINUA!", vbInformation, "SACC"
            Check2.Value = 0
        Else
            Check1.Value = 0
        End If
    End If
End Sub
Private Sub Command6_Click()
    If Text1.Text = "" Then
        txtIDCOMANDA.Text = "1"
    End If
    If Option9.Value = True Then
        If Text10.Text <> "" Then
            ExtraeComanda
        End If
    Else
        If Text10.Text <> "" Then
            ExtraerAsistencia
        End If
    End If
End Sub
Private Sub Coti_Click()
    FrmCotizaRapida.Show vbModal
End Sub
Private Sub Factu_Click()
    If VarMen.Text1(8).Text = "S" Then
        If VarMen.TxtEmp(10).Text = "S" Then
            Call ShellExecute(Me.hWnd, "Open", App.Path & "\FactElec", "", "", 1)
        Else
            frmFactura.Show vbModal
        End If
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    If GetSetting("APTONER", "ConfigSACC", "Remision", "S") = "S" Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If
    DTPFechaDomi.Value = Date
    If VarMen.Text1(12).Text = "N" Then
        Me.SSTab1.TabEnabled(3) = False
    End If
    If VarMen.Text1(56).Text = "N" Then
        Me.SSTab1.TabEnabled(2) = False
    End If
    If VarMen.Text1(6).Text = "N" Then
        Me.SSTab1.TabEnabled(1) = False
    End If
    If VarMen.Text1(7).Text = "N" Then
        Me.SSTab1.TabEnabled(0) = False
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL CLIENTE", 600
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DESCUENTO", 1200
        .ColumnHeaders.Add , , "DIAS DE CREDITO", 1200
        .ColumnHeaders.Add , , "LIMITE DE CREDITO", 1200
        .ColumnHeaders.Add , , "TEL. CASA", 1200
        .ColumnHeaders.Add , , "TEL. TRABAJO", 1200
        .ColumnHeaders.Add , , "DIRECCION", 3200
        .ColumnHeaders.Add , , "NO. INT.", 800
        .ColumnHeaders.Add , , "NO. EXT.", 800
        .ColumnHeaders.Add , , "COLONIA", 1200
        .ColumnHeaders.Add , , "CIUDAD", 1200
        .ColumnHeaders.Add , , "TIPO DESCUENTO", 1200
        .ColumnHeaders.Add , , "NOMBRE COMERCIAL", 6100
        .ColumnHeaders.Add , , "RFC", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 3600
        .ColumnHeaders.Add , , "PRECIO", 1000
        .ColumnHeaders.Add , , "EXISTENCIA", 1000
        .ColumnHeaders.Add , , "CLASIFICACION", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 6800
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "COMANDA O ASISTENCIA", 2000
        .ColumnHeaders.Add , , "IVA", 2000
        .ColumnHeaders.Add , , "RETENCIÓN", 2000
        .ColumnHeaders.Add , , "IMPUESTO 1", 2000
        .ColumnHeaders.Add , , "IMPUESTO 2", 2000
    End With
    CLVCLIEN = ""
    txtID_User.Text = VarMen.Text1(0).Text
    txtID_User.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    'asistencia tecnica
    LblMenu.Caption = VarMen.Text1(1).Text
    LblMenu2.Caption = VarMen.Text4(0).Text
    DtPFechAsi.Value = Format(Date, "dd/mm/yyyy")
    ' Domicilios
    Me.BtnGuardaDomi.Enabled = False
    DTPFechaDomi.Value = Format(Date, "dd/mm/yyyy")
    LlenaCombo
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    If ListView3.ListItems.Count = 0 Then
        Unload Me
    Else
        MsgBox "Elimine los articulos para salir", vbInformation, "SACC"
    End If
End Sub
Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        frmVerCliente.Te1.Text = ListView1.SelectedItem
        frmVerCliente.Show vbModal
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    'para domicilios
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim T_VENTA As Double
    Dim T_ABONO As Double
    Me.TxtNomCliente.Text = Item.SubItems(1)
    Me.TxtDomiCleinte.Text = Item.SubItems(7) & " #" & Item.SubItems(9) & " -" & Item.SubItems(8)
    Me.CmbColonia.Text = Item.SubItems(10)
    Me.TxtTelefonoDomi = Item.SubItems(5)
    Text12.Text = Item.SubItems(1)
    Text13.Text = Item.SubItems(5)
    'para comandas
    Me.txtId_Cliente.Text = Item
    bBCli = True
    'para asistencia tecnica
    IdClien = Item
    TxtNomClienAs.Text = Item.SubItems(1)
    Text14.Text = Item.SubItems(5)
    NomClien = Item.SubItems(1)
    TelCasa = Item.SubItems(5)
    TelTrabajo = Item.SubItems(6)
    Direc = Item.SubItems(7)
    NoExte = Item.SubItems(9)
    NoInte = Item.SubItems(8)
    COLONIA = Item.SubItems(10)
    RFC = Item.SubItems(14)
    'para ventas
    Frame5.Enabled = True
    Text1.Text = Item.SubItems(1)
    CLVCLIEN = Item
    NomClien = Item.SubItems(1)
    DesClien = Item.SubItems(2)
    'Combo2.Text = Item.SubItems(3)
    'Text9.Text = Item.SubItems(4)
    If Item.SubItems(3) = "" Then
        Item.SubItems(3) = "0"
    End If
    If Item.SubItems(4) = "" Then
        Item.SubItems(4) = "0"
    End If
    If CDbl(Item.SubItems(3)) > 0 And CDbl(Item.SubItems(4)) > 0 Then
        sBuscar = "SELECT SUM(CUENTAS.TOTAL_COMPRA) AS TOT_VENTAS FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("TOT_VENTAS")) Then
                T_VENTA = CDbl(tRs.Fields("TOT_VENTAS"))
            Else
                T_VENTA = 0
            End If
        End If
        sBuscar = "SELECT SUM(ABONOS_CUENTA.CANT_ABONO) AS TOT_ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("TOT_ABONO")) Then
                T_ABONO = CDbl(tRs.Fields("TOT_ABONO"))
            Else
                T_ABONO = 0
            End If
        End If
        If (CDbl(Item.SubItems(4)) - (CDbl(T_VENTA) - CDbl(T_ABONO))) > 0 Then
            Text9.Text = CDbl(Item.SubItems(4)) - (CDbl(T_VENTA) - CDbl(T_ABONO))
            Combo2.Text = Item.SubItems(3)
        Else
            Text9.Text = 0
            Combo2.Text = 0
            MsgBox "El cliente a excedido su limite de credito!", vbExclamation, "SACC"
        End If
    End If
    If Item.SubItems(12) <> "" Then
        Label32.Caption = Item.SubItems(12)
        Label1.Caption = "DESCUENTO :"
    Else
        Label32.Caption = ""
        Label1.Caption = ""
    End If
    If (Item.SubItems(4) = "") Or (CDbl(Item.SubItems(4)) <= 0) Then
        Frame5.Enabled = False
    End If
    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
        Me.Command2.Enabled = True
    Else
        Me.Command2.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CLVCLIEN <> "" Then
            If Me.ListView1.ListItems.Count <> 0 Then
                If Text1.Text <> "" Then
                    Me.Text2.SetFocus
                End If
            End If
        Else
            Me.Text1.SetFocus
        End If
    End If
End Sub
Private Sub ListView1_LostFocus()
On Error GoTo ManejaError
    If CLVCLIEN <> "" Then
        If Me.ListView1.ListItems.Count <> 0 Then
            If Text1.Text <> "" Then
                Text1.Enabled = False
                ListView1.Enabled = False
            End If
        End If
    Else
        Text1.SetFocus
    End If
Exit Sub
ManejaError:
        Err.Clear
End Sub
Private Sub ListView2_DblClick()
    Text8.SetFocus
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Option4.Value = True Then
        Text2.Text = Item
    Else
        Text2.Text = Item.SubItems(1)
    End If
    ClvProd = Item
    DesProd = Item.SubItems(1)
    PreProd = Item.SubItems(2)
    ClasProd = Item.SubItems(4)
    If Text8.Text <> "" And ClvProd <> "" Then
        Me.Command4.Enabled = True
    Else
        Me.Command4.Enabled = False
    End If
    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
        Me.Command2.Enabled = True
    Else
        Me.Command2.Enabled = False
    End If
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DelInd = Item
    DelDes = Item.SubItems(1)
    DelCan = Val(Replace(Item.SubItems(2), ",", ""))
    DelPre = Item.SubItems(3)
    DelIndex = Item.Index
    If Item.SubItems(5) = "" Then
        DelIVA = 0
    Else
        DelIVA = Item.SubItems(5)
    End If
    If Item.SubItems(6) = "" Then
        DelRET = 0
    Else
        DelRET = Item.SubItems(6)
    End If
    If Item.SubItems(7) = "" Then
        DelIMP1 = 0
    Else
        DelIMP1 = Item.SubItems(7)
    End If
    If Item.SubItems(8) = "" Then
        DelIMP2 = 0
    Else
        DelIMP2 = Item.SubItems(8)
    End If
    Me.Command3.Enabled = True
End Sub
Private Sub lvwProductosComanda_DblClick()
    If Me.lvwProductosComanda.SelectedItem.Selected Then
        Me.txtProductoComanda.Text = Me.lvwProductosComanda.SelectedItem
        Me.txtCantidadComanda.SetFocus
    End If
End Sub
Private Sub reim_Click()
    FrmReImprime.Show vbModal
End Sub
Private Sub SubMenCambiaFormPago_Click()
    FrmModiVenta.Show vbModal
End Sub
Private Sub SubMenConCobCom_Click()
    FrmConsultaComanda.Show vbModal
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim CadClien As String
        If Not IsNumeric(Text1.Text) Then
            CadClien = Text1.Text
            CadClien = Replace(CadClien, " ", "%")
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO, TELEFONO_CASA, TELEFONO_TRABAJO, DIRECCION, NUMERO_EXTERIOR, NUMERO_INTERIOR, COLONIA, CIUDAD, ID_DESCUENTO, NOMBRE_COMERCIAL, RFC FROM CLIENTE WHERE NOMBRE_COMERCIAL LIKE '%" & Trim(CadClien) & "%' AND VALORACION NOT LIKE 'E' OR NOMBRE LIKE '%" & Trim(CadClien) & "%' AND VALORACION NOT LIKE 'E'"
        Else
            CadClien = Text1.Text
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO, TELEFONO_CASA, TELEFONO_TRABAJO, DIRECCION, NUMERO_EXTERIOR, NUMERO_INTERIOR, COLONIA, CIUDAD, ID_DESCUENTO, NOMBRE_COMERCIAL, RFC FROM CLIENTE WHERE NOMBRE_COMERCIAL LIKE '%" & Trim(CadClien) & "%' AND VALORACION NOT LIKE 'E' OR NOMBRE LIKE '%" & Trim(CadClien) & "%' AND VALORACION NOT LIKE 'E' OR ID_CLIENTE = " & Text1.Text & " AND VALORACION NOT LIKE 'E'"
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE")
                If Not IsNull(.Fields("DESCUENTO")) Then
                    tLi.SubItems(2) = .Fields("DESCUENTO")
                Else
                    tLi.SubItems(2) = "0.00"
                End If
                If Not IsNull(.Fields("DIAS_CREDITO")) Then tLi.SubItems(3) = .Fields("DIAS_CREDITO")
                If Not IsNull(.Fields("LIMITE_CREDITO")) Then tLi.SubItems(4) = .Fields("LIMITE_CREDITO")
                If Not IsNull(.Fields("TELEFONO_CASA")) Then tLi.SubItems(5) = .Fields("TELEFONO_CASA")
                If Not IsNull(.Fields("TELEFONO_TRABAJO")) Then tLi.SubItems(6) = .Fields("TELEFONO_TRABAJO")
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(7) = .Fields("DIRECCION")
                If Not IsNull(.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(8) = .Fields("NUMERO_EXTERIOR")
                If Not IsNull(.Fields("NUMERO_INTERIOR")) Then tLi.SubItems(9) = .Fields("NUMERO_INTERIOR")
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(10) = .Fields("COLONIA")
                If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(11) = .Fields("CIUDAD")
                If Not IsNull(.Fields("ID_DESCUENTO")) Then tLi.SubItems(12) = .Fields("ID_DESCUENTO")
                If Not IsNull(.Fields("NOMBRE_COMERCIAL")) Then tLi.SubItems(13) = .Fields("NOMBRE_COMERCIAL")
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(14) = .Fields("RFC")
                .MoveNext
            Loop
            Me.ListView1.SetFocus
        End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Option9.Value = True Then
            ExtraeComanda
        Else
            ExtraerAsistencia
        End If
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text10_GotFocus()
    Text10.BackColor = &HFFE1E1
End Sub
Private Sub Text10_LostFocus()
    Text10.BackColor = &H80000005
End Sub
Private Sub Text19_KeyPress(KeyAscii As Integer)
Dim Valido As String
    Valido = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -/1234567890,_*+=!?()<>."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    Dim sBuscar As String
    Dim CompDolar As Double
    Dim totalreal As Double
    If KeyAscii = 13 And Text2.Text <> "" Then
        SUC = VarMen.Text4(0).Text
        sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("COMPRA")) Then
                CompDolar = tRs.Fields("COMPRA")
            Else
                CompDolar = InputBox("POR FAVOR, DE EL PRECIO DE VENTA DEL DOLAR HOY!", "SACC")
                sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES (DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
                cnn.Execute (sBuscar)
            End If
        Else
            CompDolar = InputBox("POR FAVOR, DE EL PRECIO DE VENTA DEL DOLAR HOY!", "SACC")
            sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES (DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
            cnn.Execute (sBuscar)
        End If
        If Option4.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION, PE, IVA, P_RETENCION, IMPUESTO1, IMPUESTO2 FROM VSVENTAS WHERE ID_PRODUCTO LIKE '%" & Trim(Text2.Text) & "%' AND SUCURSAL = '" & SUC & "'" 'Cambiado 25/09/06
            Set tRs = cnn.Execute(sBus)
        End If                                                                                                                                     'Se cambio Almacen3 por VsVentas
        If Option3.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION, PE, IVA, P_RETENCION, IMPUESTO1, IMPUESTO2 FROM VSVENTAS WHERE Descripcion LIKE '%" & Trim(Text2.Text) & "%' AND SUCURSAL = '" & SUC & "'" 'Cambiado 25/09/06
            Set tRs = cnn.Execute(sBus)
        End If                                                                                                                           'Se cambio Almacen3 por VsVentas
        If Option5.Value = True Then
            sBus = "SELECT ID_PRODUCTO FROM ENTRADA_PRODUCTO WHERE CODIGO_BARAS = '" & Text2.Text & "'"
            Set tRs = cnn.Execute(sBus)
            If Not (tRs.EOF And tRs.BOF) Then
                sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION,MONEDA FROM VSVENTAS WHERE ID_PRODUCTO LIKE '%" & tRs.Fields("ID_PRODUCTO") & "%' AND SUCURSAL = '" & SUC & "'" 'Cambiado 25/09/06
                Set tRs = cnn.Execute(sBus)
            Else
                MsgBox "EL CODIGO DE BARRAS NO ESTA REGISTRADO, INTENTE OTRO MODO DE BUSQUEDA!", vbInformation, "SACC"
            End If
        End If
        If sBus <> "" Then
            With tRs
                ListView2.ListItems.Clear
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                            Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                            If (.Fields("PE")) = "PESOS" Then
                                If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                                    tLi.SubItems(2) = (Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "0.00"))
                                End If
                            Else
                            If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                                    tLi.SubItems(2) = (Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "0.00")) * CompDolar
                                End If
                            End If
                                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = .Fields("CANTIDAD") & ""
                                If Not IsNull(.Fields("CLASIFICACION")) Then tLi.SubItems(4) = .Fields("CLASIFICACION") & ""
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
        Me.ListView2.SetFocus
    End If
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_Change()
    If Text8.Text <> "" And ClvProd <> "" Then
        Me.Command4.Enabled = True
    Else
        Me.Command4.Enabled = False
    End If
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        If ClvProd <> "" And Text8.Text <> "" Then
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ClvProd & "' AND CANTIDAD >= " & Text8.Text & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
        Else
            Exit Sub
        End If
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(Text8.Text) & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            Agregar
            Text2.SetFocus
            Text1.Enabled = False
            ListView1.Enabled = False
            If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
                Me.Command2.Enabled = True
            Else
                Me.Command2.Enabled = False
            End If
        Else
            MsgBox "LA CANTIDAD ES MAYOR A LA EXISTENCIA!", vbInformation, "SACC"
        End If
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Agregar2()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim PreFin As String
    Dim PreFinDes As String
    Dim vDescApli As String
    Dim tLi As ListItem
    Dim iConProd As Double
    Dim iCon As Integer
    Dim RETENCION As String
    Dim IVA As String
    Dim IMPUESTO1 As String
    Dim IMPUESTO2 As String
    Set tLi = ListView3.ListItems.Add(, , ClvProd)
    tLi.SubItems(1) = DesProd
    tLi.SubItems(2) = Text8.Text
    sBuscar = "SELECT P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(ClvProd) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        RETENCION = tRs.Fields("P_RETENCION")
        IVA = tRs.Fields("IVA")
        IMPUESTO1 = tRs.Fields("IMPUESTO1")
        IMPUESTO2 = tRs.Fields("IMPUESTO2")
    Else
        RETENCION = "0.00"
        IVA = VarMen.Text4(7).Text
        IMPUESTO1 = "0.00"
        IMPUESTO2 = "0.00"
    End If
    sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & Trim(ClvProd) & "' AND ID_CLIENTE  = " & IdClien & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        If Label32.Caption <> "" Then
            sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & ClasProd & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            If tRs1.EOF And tRs1.BOF Then
                sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs2.EOF And tRs2.BOF) And Combo1.Text = "PROMOCION" Then
                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                    If DesClien <> "" Then
                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                    End If
                    If Not (tRs2.EOF And tRs2.BOF) Then
                        PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "0.00")
                    Else
                        If DesClien <> "" Then
                            If PreFin > PreFinDes Then
                                PreFin = Format(PreFinDes, "0.00")
                            Else
                                PreFin = Format(PreFin, "0.00")
                            End If
                        End If
                    End If
                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                    tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                    tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                    tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                    tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
                Else
                    sBuscar = "SELECT PORCE_DESC, CANTIDAD FROM PROMO_PORCE WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                    Set tRs2 = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                            If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                If DesClien <> "" Then
                                    PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                End If
                                If Not (tRs2.EOF And tRs2.BOF) Then
                                    vDescApli = CDbl(Text8.Text) \ CDbl(tRs2.Fields("CANTIDAD"))
                                    vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                    PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                Else
                                    If DesClien <> "" Then
                                        If PreFin > PreFinDes Then
                                            PreFin = Format(PreFinDes, "0.00")
                                        Else
                                            PreFin = Format(PreFin, "0.00")
                                        End If
                                    End If
                                End If
                                Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                                tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                                tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                                tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                                tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
                            End If
                        Else
                            sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                            Set tRs2 = cnn.Execute(sBuscar)
                            sBuscar = "SELECT CATEGORIA, PORCE_DESC, CANTIDAD FROM PROMO_CATEGO WHERE CATEGORIA = '" & tRs2.Fields("CATEGORIA") & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                            Set tRs2 = cnn.Execute(sBuscar)
                            If Not (tRs2.EOF And tRs2.BOF) Then
                                For iCon = 1 To ListView3.ListItems.Count
                                    sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView3.ListItems(iCon) & "'"
                                    Set tRs3 = cnn.Execute(sBuscar)
                                    If tRs3.Fields("CATEGORIA") = tRs2.Fields("CATEGORIA") Then
                                        iConProd = iConProd + 1
                                    End If
                                Next iCon
                                iConProd = (iConProd Mod tRs2.Fields("CANTIDAD")) + CDbl(Text8.Text)
                            Else
                                iConProd = CDbl(Text8.Text)
                            End If
                            If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                                If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                    If DesClien <> "" Then
                                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                    End If
                                    If Not (tRs2.EOF And tRs2.BOF) Then
                                        vDescApli = CDbl(iConProd) \ CDbl(tRs2.Fields("CANTIDAD"))
                                        vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                        PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                    Else
                                        If DesClien <> "" Then
                                            If PreFin > PreFinDes Then
                                                PreFin = Format(PreFinDes, "0.00")
                                            Else
                                                PreFin = Format(PreFin, "0.00")
                                            End If
                                        End If
                                    End If
                                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                                    tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                                    tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                                    tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                                    tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
                                End If
                            End If
                        End If
                    Else
                        sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        sBuscar = "SELECT CATEGORIA, PORCE_DESC, CANTIDAD FROM PROMO_CATEGO WHERE CATEGORIA = '" & tRs2.Fields("CATEGORIA") & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            For iCon = 1 To ListView3.ListItems.Count
                                sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView3.ListItems(iCon) & "'"
                                Set tRs3 = cnn.Execute(sBuscar)
                                If tRs3.Fields("CATEGORIA") = tRs2.Fields("CATEGORIA") Then
                                    iConProd = iConProd + 1
                                End If
                            Next iCon
                            iConProd = (iConProd Mod tRs2.Fields("CANTIDAD")) + CDbl(Text8.Text)
                        Else
                            iConProd = CDbl(Text8.Text)
                        End If
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                                If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                    If DesClien <> "" Then
                                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                    End If
                                    If Not (tRs2.EOF And tRs2.BOF) Then
                                        vDescApli = CDbl(iConProd) \ CDbl(tRs2.Fields("CANTIDAD"))
                                        vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                        PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                    Else
                                        If DesClien <> "" Then
                                            If PreFin > PreFinDes Then
                                                PreFin = Format(PreFinDes, "0.00")
                                            Else
                                                PreFin = Format(PreFin, "0.00")
                                            End If
                                        End If
                                    End If
                                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                                    tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                                    tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                                    tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                                    tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
                                End If
                            End If
                        Else
                            If DesClien <> "" Then
                                If PreFin > PreFinDes Then
                                    PreFin = Format(PreFinDes, "0.00")
                                Else
                                    PreFin = Format(PreFin, "0.00")
                                End If
                            End If
                            sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                            Set tRs3 = cnn.Execute(sBuscar)
                            PreFin = Format(CDbl(tRs3.Fields("PRECIO_COSTO") * Text8.Text) * (1 + CDbl(CDbl(tRs3.Fields("GANANCIA")))), "0.00")
                            Text3.Text = Format(CDbl(Text3.Text) + (CDbl(PreFin)), "0.00")
                            tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                            tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                            tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                            tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                            tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
                        End If
                    End If
                End If
            Else
                PreFin = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "0.00")
                Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                tLi.SubItems(3) = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "0.00")
                tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
                tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
                tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
                tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
            End If
            If RFC = "XEXX010101000" Then
                Text4.Text = "0.00"
                Text5.Text = Format(CDbl(Text3.Text), "0.00")
            Else
                Text4.Text = Format(Val(Replace(Text3.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
                Text5.Text = Format(CDbl(Text3.Text) * CDbl(1 + (CDbl(VarMen.Text4(7).Text) / 100)), "0.00")
            End If
        Else
            sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
            If DesClien <> "" Then
                PreFinDes = PreFin * (100 - Val(DesClien)) / 100
            End If
            If Not (tRs2.EOF And tRs2.BOF) And Combo1.Text = "PROMOCION" Then
                PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "0.00")
            End If
            If DesClien <> "" And DesClien <> "0" Then
                If PreFin < PreFinDes Then
                    PreFin = Format(PreFinDes, "0.00")
                Else
                    PreFin = Format(PreFin, "0.00")
                End If
            End If
            Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
            'IMPUESTOS Y TOTALES
            If RFC = "XEXX010101000" Then
                Text4.Text = "0.00"
                Text17.Text = "0.00"
                Text15.Text = "0.00"
                Text16.Text = "0.00"
                Text5.Text = Format(CDbl(Text3.Text), "0.00")
            Else
                Text4.Text = Format(CDbl(Text4.Text) + (CDbl(PreFin) * CDbl(IVA)), "0.00")
                Text17.Text = Format(CDbl(Text17.Text) + (CDbl(PreFin) * CDbl(RETENCION)), "0.00")
                Text15.Text = Format(CDbl(Text15.Text) + (CDbl(PreFin) * CDbl(IMPUESTO1)), "0.00")
                Text16.Text = Format(CDbl(Text16.Text) + (CDbl(PreFin) * CDbl(IMPUESTO2)), "0.00")
                Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) - CDbl(Text17.Text) + CDbl(Text15.Text) + CDbl(Text16.Text), "0.00")
            End If
            tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
            tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
            tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
            tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
            tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
            ClvProd = ""
            DesProd = ""
        End If
    Else
        PreFin = Format(CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text), "0.00")
        Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
        If RFC = "XEXX010101000" Then
            Text4.Text = "0.00"
            Text17.Text = "0.00"
            Text15.Text = "0.00"
            Text16.Text = "0.00"
            Text5.Text = Format(CDbl(Text3.Text), "0.00")
        Else
            Text4.Text = Format(CDbl(Text4.Text) + (CDbl(PreFin) * CDbl(IVA)), "0.00")
            Text17.Text = Format(CDbl(Text17.Text) + (CDbl(PreFin) * CDbl(RETENCION)), "0.00")
            Text15.Text = Format(CDbl(Text15.Text) + (CDbl(PreFin) * CDbl(IMPUESTO1)), "0.00")
            Text16.Text = Format(CDbl(Text16.Text) + (CDbl(PreFin) * CDbl(IMPUESTO2)), "0.00")
            Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) - CDbl(Text17.Text) + CDbl(Text15.Text) + CDbl(Text16.Text), "0.00")
        End If
        tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
        tLi.SubItems(5) = Format(CDbl(PreFin) * CDbl(IVA), "0.00")
        tLi.SubItems(6) = Format(CDbl(PreFin) * CDbl(RETENCION), "0.00")
        tLi.SubItems(7) = Format(CDbl(PreFin) * CDbl(IMPUESTO1), "0.00")
        tLi.SubItems(8) = Format(CDbl(PreFin) * CDbl(IMPUESTO2), "0.00")
        ClvProd = ""
        DesProd = ""
    End If
    If CDbl(Text5.Text) > CDbl(Text9.Text) And Check2.Value = 1 Then
        MsgBox "El limite de credito ha sido sobregirado, debe retirar articulos para cerrar la venta!", vbCritical, "SACC"
        Command4.Visible = False
        Command2.Enabled = False
    End If
    IVA = "0.00"
    IMPUESTO1 = "0.00"
    IMPUESTO2 = "0.00"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Agregar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim PreFin As String
    Dim PreFinDes As String
    Dim vDescApli As String
    Dim tLi As ListItem
    Dim iConProd As Double
    Dim iCon As Integer
    Dim tRs3 As ADODB.Recordset
    Set tLi = ListView3.ListItems.Add(, , ClvProd)
    tLi.SubItems(1) = DesProd
    tLi.SubItems(2) = Text8.Text
    sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & Trim(ClvProd) & "' AND ID_CLIENTE  = " & IdClien & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        If Label32.Caption <> "" Then
            sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & ClasProd & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            If tRs1.EOF And tRs1.BOF Then
                sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs2.EOF And tRs2.BOF) And Combo1.Text = "PROMOCION" Then
                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                    If DesClien <> "" Then
                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                    End If
                    If Not (tRs2.EOF And tRs2.BOF) Then
                        PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "0.00")
                    Else
                        If DesClien <> "" Then
                            If PreFin > PreFinDes Then
                                PreFin = Format(PreFinDes, "0.00")
                            Else
                                PreFin = Format(PreFin, "0.00")
                            End If
                        End If
                    End If
                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                Else
                    sBuscar = "SELECT PORCE_DESC, CANTIDAD FROM PROMO_PORCE WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                    Set tRs2 = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                            If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                If DesClien <> "" Then
                                    PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                End If
                                If Not (tRs2.EOF And tRs2.BOF) Then
                                    vDescApli = CDbl(Text8.Text) \ CDbl(tRs2.Fields("CANTIDAD"))
                                    vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                    PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                Else
                                    If DesClien <> "" Then
                                        If PreFin > PreFinDes Then
                                            PreFin = Format(PreFinDes, "0.00")
                                        Else
                                            PreFin = Format(PreFin, "0.00")
                                        End If
                                    End If
                                End If
                                Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                            End If
                        Else
                            sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                            Set tRs2 = cnn.Execute(sBuscar)
                            sBuscar = "SELECT CATEGORIA, PORCE_DESC, CANTIDAD FROM PROMO_CATEGO WHERE CATEGORIA = '" & tRs2.Fields("CATEGORIA") & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                            Set tRs2 = cnn.Execute(sBuscar)
                            If Not (tRs2.EOF And tRs2.BOF) Then
                                For iCon = 1 To ListView3.ListItems.Count
                                    sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView3.ListItems(iCon) & "'"
                                    Set tRs3 = cnn.Execute(sBuscar)
                                    If tRs3.Fields("CATEGORIA") = tRs2.Fields("CATEGORIA") Then
                                        iConProd = iConProd + 1
                                    End If
                                Next iCon
                                iConProd = (iConProd Mod tRs2.Fields("CANTIDAD")) + CDbl(Text8.Text)
                            Else
                                iConProd = CDbl(Text8.Text)
                            End If
                            If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                                If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                    If DesClien <> "" Then
                                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                    End If
                                    If Not (tRs2.EOF And tRs2.BOF) Then
                                        vDescApli = CDbl(iConProd) \ CDbl(tRs2.Fields("CANTIDAD"))
                                        vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                        PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                    Else
                                        If DesClien <> "" Then
                                            If PreFin > PreFinDes Then
                                                PreFin = Format(PreFinDes, "0.00")
                                            Else
                                                PreFin = Format(PreFin, "0.00")
                                            End If
                                        End If
                                    End If
                                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                                End If
                            End If
                        End If
                    Else
                        sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        sBuscar = "SELECT CATEGORIA, PORCE_DESC, CANTIDAD FROM PROMO_CATEGO WHERE CATEGORIA = '" & tRs2.Fields("CATEGORIA") & "' AND FECHA_VENCE >= '" & Format(Date, "dd/mm/yyyy") & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            For iCon = 1 To ListView3.ListItems.Count
                                sBuscar = "SELECT CATEGORIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView3.ListItems(iCon) & "'"
                                Set tRs3 = cnn.Execute(sBuscar)
                                If tRs3.Fields("CATEGORIA") = tRs2.Fields("CATEGORIA") Then
                                    iConProd = iConProd + 1
                                End If
                            Next iCon
                            iConProd = (iConProd Mod tRs2.Fields("CANTIDAD")) + CDbl(Text8.Text)
                        Else
                            iConProd = CDbl(Text8.Text)
                        End If
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If CDbl(tRs2.Fields("CANTIDAD")) >= CDbl(Text8.Text) Then
                                If tRs2.EOF And tRs2.BOF And Combo1.Text = "PROMOCION" Then
                                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
                                    If DesClien <> "" Then
                                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                                    End If
                                    If Not (tRs2.EOF And tRs2.BOF) Then
                                        vDescApli = CDbl(iConProd) \ CDbl(tRs2.Fields("CANTIDAD"))
                                        vDescApli = ((CDbl(tRs2.Fields("PORCE_DESC")) / 100) * CDbl(tRs.Fields("PRECIO_VENTA"))) * CDbl(vDescApli)
                                        PreFin = Format((CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text)) * CDbl(vDescApli), "0.00")
                                    Else
                                        If DesClien <> "" Then
                                            If PreFin > PreFinDes Then
                                                PreFin = Format(PreFinDes, "0.00")
                                            Else
                                                PreFin = Format(PreFin, "0.00")
                                            End If
                                        End If
                                    End If
                                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                                    tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                                End If
                            End If
                        Else
                            If DesClien <> "" Then
                                If PreFin > PreFinDes Then
                                    PreFin = Format(PreFinDes, "0.00")
                                Else
                                    PreFin = Format(PreFin, "0.00")
                                End If
                            End If
                            sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                            Set tRs3 = cnn.Execute(sBuscar)
                            PreFin = Format(CDbl(tRs3.Fields("PRECIO_COSTO") * Text8.Text) * (1 + CDbl(CDbl(tRs3.Fields("GANANCIA")))), "0.00")
                            Text3.Text = Format(CDbl(Text3.Text) + (CDbl(PreFin)), "0.00")
                            tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
                        End If
                    End If
                End If
            Else
                PreFin = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "0.00")
                Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
                tLi.SubItems(3) = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "0.00")
            End If
            If RFC = "XEXX010101000" Then
                Text4.Text = "0.00"
                Text5.Text = Format(CDbl(Text3.Text), "0.00")
            Else
                Text4.Text = Format(Val(Replace(Text3.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
                Text5.Text = Format(CDbl(Text3.Text) * CDbl(1 + (CDbl(VarMen.Text4(7).Text) / 100)), "0.00")
            End If
        Else
            sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "0.00")
            If DesClien <> "" Then
                PreFinDes = PreFin * (100 - Val(DesClien)) / 100
            End If
            If Not (tRs2.EOF And tRs2.BOF) And Combo1.Text = "PROMOCION" Then
                PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "0.00")
            End If
            If DesClien <> "" And DesClien <> "0" Then
                If PreFin < PreFinDes Then
                    PreFin = Format(PreFinDes, "0.00")
                Else
                    PreFin = Format(PreFin, "0.00")
                End If
            End If
            Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
            If RFC = "XEXX010101000" Then
                Text4.Text = "0.00"
                Text5.Text = Format(CDbl(Text3.Text), "0.00")
            Else
                Text4.Text = Format(Val(Replace(Text3.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
                Text5.Text = Format(CDbl(Text3.Text) * CDbl(1 + (CDbl(VarMen.Text4(7).Text) / 100)), "0.00")
            End If
            tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
            ClvProd = ""
            DesProd = ""
        End If
    Else
        PreFin = Format(CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(Text8.Text), "0.00")
        Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "0.00")
        If RFC = "XEXX010101000" Then
            Text4.Text = "0.00"
            Text5.Text = Format(CDbl(Text3.Text), "0.00")
        Else
            Text4.Text = Format(Val(Replace(Text3.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
            Text5.Text = Format(CDbl(Text3.Text) * CDbl(1 + (CDbl(VarMen.Text4(7).Text) / 100)), "0.00")
        End If
        tLi.SubItems(3) = Format(CDbl(PreFin), "0.00")
        ClvProd = ""
        DesProd = ""
    End If
    If CDbl(Text5.Text) > CDbl(Text9.Text) And Check2.Value = 1 Then
        MsgBox "El limite de credito ha sido sobregirado, debe retirar articulos para cerrar la venta!", vbCritical, "SACC"
        Command4.Visible = False
        Command2.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text9_Change()
    If Text9.Text = "" Then
        Text9.Text = "0.00"
    Else
        If CDbl(Text9.Text) = 0 Then
            Text9.Text = "0.00"
        End If
    End If
End Sub
Private Sub Text9_GotFocus()
    Text9.BackColor = &HFFE1E1
End Sub
Private Sub Text9_LostFocus()
    Text9.BackColor = &H80000005
End Sub
Private Sub ExtraeComanda()
'On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim cant As String
    Dim PreTot As String
    Dim PreTotDes As String
    Dim ClvProd As String
    DesClien = ""
    sBuscar = "SELECT ID_CLIENTE FROM COMANDAS_2 WHERE ID_COMANDA = " & Text10.Text & " AND TIPO = 'C'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, CANTIDAD_NO_SIRVIO FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Text10.Text & " AND ESTADO_ACTUAL IN ('L','N')"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                MsgBox "EL CLIENTE HA SIDO ELIMINADO, DEBE REASIGNAR UN CLIENTE!", vbInformation, "SACC"
            Else
                If txtIDCOMANDA.Text = "1" Then
                    If Not IsNull(tRs.Fields("NOMBRE")) Then Text1.Text = tRs.Fields("NOMBRE")
                    txtIDCOMANDA.Text = "0"
                End If
                If Text1.Text = tRs.Fields("NOMBRE") Then
                    Me.TxtNomCliente.Text = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("NOMBRE")) Then Text1.Text = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("ID_CLIENTE")) Then
                        CLVCLIEN = tRs.Fields("ID_CLIENTE")
                        sBuscar = "SELECT ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If Not IsNull(tRs2.Fields("ID_DESCUENTO")) Then Label32.Caption = tRs2.Fields("ID_DESCUENTO")
                        End If
                    End If
                    If Not IsNull(tRs.Fields("NOMBRE")) Then NomClien = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("DESCUENTO")) Then DesClien = tRs.Fields("DESCUENTO")
                    If Not IsNull(tRs.Fields("DIAS_CREDITO")) Then Combo2.Text = tRs.Fields("DIAS_CREDITO")
                    If Not IsNull(tRs.Fields("LIMITE_CREDITO")) Then Text9.Text = tRs.Fields("LIMITE_CREDITO")
                    If DesClien = "" Then DesClien = "0"
                    Text1.Enabled = False
                    ListView1.Enabled = False
                    If Not (tRs1.BOF And tRs1.EOF) Then
                        If txtComandas.Text = "" Then
                            txtComandas.Text = Text10.Text
                        Else
                            txtComandas.Text = txtComandas.Text & ", " & Text10.Text
                        End If
                        Do While Not tRs1.EOF
                            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA, CLASIFICACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs1.Fields("ID_PRODUCTO") & "'"
                            Set tRs2 = cnn.Execute(sBuscar)
                            If Not (tRs2.BOF And tRs2.EOF) Then
                                ClvProd = tRs2.Fields("ID_PRODUCTO")
                                If CDbl(tRs1.Fields("CANTIDAD")) - CDbl(tRs1.Fields("CANTIDAD_NO_SIRVIO")) <> 0 Then
                                    Set tLi = ListView3.ListItems.Add(, , tRs2.Fields("ID_PRODUCTO"))
                                        If Not IsNull(tRs2.Fields("Descripcion")) Then tLi.SubItems(1) = tRs2.Fields("Descripcion")
                                        If Not IsNull(tRs1.Fields("CANTIDAD")) Then tLi.SubItems(2) = CDbl(tRs1.Fields("CANTIDAD")) - CDbl(tRs1.Fields("CANTIDAD_NO_SIRVIO"))
                                        If Label32.Caption = "" Then
                                        If Combo1.Text = "<NINGUNA>" Or Combo1.Text = "" Then
                                            If Label32.Caption <> "" Then
                                                sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                                Set tRs3 = cnn.Execute(sBuscar)
                                                If Not (tRs3.EOF And tRs3.BOF) Then
                                                    PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "0.00")
                                                Else
                                                    If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                    PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    If DesClien <> "" Then
                                                        PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                    End If
                                                End If
                                            Else
                                                PreTot = Format(CDbl((tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1)) * CDbl(tRs1.Fields("CANTIDAD")), "0.00")
                                            End If
                                        Else
                                            If Combo1.Text = "LICITACIÓN" Then
                                                sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & CLVCLIEN
                                                Set tRs3 = cnn.Execute(sBuscar)
                                                If Not (tRs3.EOF And tRs3.BOF) Then
                                                    PreTot = Format(CDbl(tRs3.Fields("PRECIO_VENTA")) * Val(Replace(tLi.SubItems(2), ",", "")), "0.00")
                                                Else
                                                    PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                    PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                End If
                                            Else
                                                If VarMen.Text4(8).Text = "S" Then
                                                    sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND TIPO = '" & Combo1.Text & "'"
                                                    Set tRs3 = cnn.Execute(sBuscar)
                                                    If Not (tRs3.EOF And tRs3.BOF) Then
                                                        PreTot = CDbl(tRs3.Fields("PRECIO_OFERTA")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTotDes = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                        PreTotDes = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTotDes = Val(Replace(PreTotDes, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                        If Val(Replace(PreTot, ",", "")) > Val(Replace(PreTotDes, ",", "")) Then PreTot = PreTotDes
                                                    Else
                                                        sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                                        Set tRs3 = cnn.Execute(sBuscar)
                                                        If Not (tRs3.EOF And tRs3.BOF) Then
                                                            PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "0.00")
                                                        Else
                                                            If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                            PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                            If DesClien <> "" Then
                                                                PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                                    Set tRs3 = cnn.Execute(sBuscar)
                                                    If Not (tRs3.EOF And tRs3.BOF) Then
                                                        PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "0.00")
                                                    Else
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                        PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                        Set tRs4 = cnn.Execute(sBuscar)
                                        If Not (tRs4.EOF And tRs4.BOF) Then
                                            PreTot = Val(tRs2.Fields("PRECIO_COSTO") * (tRs2.Fields("GANANCIA") + 1))
                                            PreTot = CDbl(PreTot) * (1 - (tRs4.Fields("PORCENTAJE") / 100)) * (tRs1.Fields("CANTIDAD") - tRs1.Fields("CANTIDAD_NO_SIRVIO"))
                                        Else
                                            If Combo1.Text = "<NINGUNA>" Or Combo1.Text = "" Then
                                                If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                If DesClien <> "" Then
                                                    PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                End If
                                            Else
                                                If Combo1.Text = "LICITACIÓN" Then
                                                    sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & CLVCLIEN
                                                    Set tRs3 = cnn.Execute(sBuscar)
                                                    If Not (tRs3.EOF And tRs3.BOF) Then
                                                        PreTot = Format(CDbl(tRs3.Fields("PRECIO_VENTA")) * Val(Replace(tLi.SubItems(2), ",", "")), "0.00")
                                                    Else
                                                        PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                        PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    End If
                                                Else
                                                    If VarMen.Text4(8).Text = "S" Then
                                                        sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND TIPO = '" & Combo1.Text & "'"
                                                        Set tRs3 = cnn.Execute(sBuscar)
                                                        If Not (tRs3.EOF And tRs3.BOF) Then
                                                            PreTot = CDbl(tRs3.Fields("PRECIO_OFERTA")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                            If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTotDes = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                            PreTotDes = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                            If DesClien <> "" Then
                                                                PreTotDes = Val(Replace(PreTotDes, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                            End If
                                                            If Val(Replace(PreTot, ",", "")) > Val(Replace(PreTotDes, ",", "")) Then PreTot = PreTotDes
                                                        Else
                                                            If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                            PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                            If DesClien <> "" Then
                                                                PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                            End If
                                                        End If
                                                    Else
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "0.00")
                                                        PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    tLi.SubItems(3) = Format(CDbl(PreTot), "0.00")
                                    tLi.SubItems(4) = "C" & Text10.Text
                                    tLi.SubItems(5) = Format(CDbl(PreTot) * CDbl(tRs2.Fields("IVA")), "0.00")
                                    tLi.SubItems(6) = Format(CDbl(PreTot) * CDbl(tRs2.Fields("P_RETENCION")), "0.00")
                                    tLi.SubItems(7) = Format(CDbl(PreTot) * CDbl(tRs2.Fields("IMPUESTO1")), "0.00")
                                    tLi.SubItems(8) = Format(CDbl(PreTot) * CDbl(tRs2.Fields("IMPUESTO2")), "0.00")
                                    If Not IsNull(tRs1.Fields("CANTIDAD")) Then Text3.Text = Format(Val(Replace(Text3.Text, ",", "")) + Val(Replace(PreTot, ",", "")), "0.00")
                                 End If
                            End If
                            If RFC = "XEXX010101000" Then
                                Text4.Text = "0.00"
                                Text17.Text = "0.00"
                                Text15.Text = "0.00"
                                Text16.Text = "0.00"
                                Text5.Text = Format(CDbl(Text3.Text), "0.00")
                            Else
                                If PreTot = "" Then PreTot = "0"
                                Text4.Text = Format(CDbl(Text4.Text) + (CDbl(PreTot) * CDbl(tRs2.Fields("IVA"))), "0.00")
                                Text17.Text = Format(CDbl(Text17.Text) + (CDbl(PreTot) * CDbl(tRs2.Fields("P_RETENCION"))), "0.00")
                                Text15.Text = Format(CDbl(Text15.Text) + (CDbl(PreTot) * CDbl(tRs2.Fields("IMPUESTO1"))), "0.00")
                                Text16.Text = Format(CDbl(Text16.Text) + (CDbl(PreTot) * CDbl(tRs2.Fields("IMPUESTO2"))), "0.00")
                                Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) - CDbl(Text17.Text) + CDbl(Text15.Text) + CDbl(Text16.Text), "0.00")
                            End If
                            tRs1.MoveNext
                        Loop
                        Me.Command2.Enabled = True
                    Else
                        MsgBox "NO SE PUEDEN COBRAR JUNTAS COMANDAS DE CLIENTES DISTINTOS!", vbInformation, "SACC"
                    End If
                End If
            End If
        Else
            MsgBox "LA COMANDA NO HA SIDO FINALIZADA O YA FUE ENTREGADA AL CLIENTE!", vbInformation, "SACC"
            txtIDCOMANDA.Text = ""
        End If
    Else
        sBuscar = "SELECT ID_CLIENTE FROM COMANDAS_2 WHERE ID_COMANDA = " & Text10.Text & " AND TIPO = 'P'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            MsgBox "ESTA INTENTANDO EXTRAER UNA ORDEN DE PRODUCCION, ESTA ORDEN DEBE SER TOMADA DE EXISTENCIA", vbInformation, "SACC"
            txtIDCOMANDA.Text = ""
        Else
            MsgBox "NO EXISTE LA COMANDA!", vbInformation, "SACC"
            txtIDCOMANDA.Text = ""
        End If
    End If
    Text10.Text = ""
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCantidadComanda_LostFocus()
    txtCantidadComanda.BackColor = &H80000005
End Sub
Private Sub TxtComTec_GotFocus()
    TxtComTec.BackColor = &HFFE1E1
End Sub
Private Sub TxtComTec_LostFocus()
    TxtComTec.BackColor = &H80000005
End Sub
Private Sub TxtDesPiez_GotFocus()
    TxtDesPiez.BackColor = &HFFE1E1
End Sub
Private Sub TxtDesPiez_LostFocus()
    TxtDesPiez.BackColor = &H80000005
End Sub
Private Sub TxtDomiCleinte_LostFocus()
    TxtDomiCleinte.BackColor = &H80000005
End Sub
Private Sub txtID_User_GotFocus()
    Me.txtID_User.SelStart = 0
    Me.txtID_User.SelLength = Len(Me.txtID_User.Text)
End Sub
Private Sub txtID_User_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++ < Asistencia Tecnica > ++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Chk1_Click()
    DOMI = chk1.Value
End Sub
Private Sub Chk1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chk1.Value = 1
        chk2.SetFocus
    End If
End Sub
Private Sub Chk2_Click()
    GTIA = chk2.Value
End Sub
Private Sub Chk2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chk2.Value = 1
        TxtModelo.SetFocus
    End If
End Sub
Private Sub cmdRegis_Click()
On Error GoTo ManejaError
    Dim sqlComanda As String
    If IdClien <> "" And IdClien <> "0" Then
        If Text14.Text <> "" And TxtNomClienAs.Text <> "" Then
            sqlComanda = "INSERT INTO ASISTENCIA_TECNICA (SUCURSAL, GARANTIA, Descripcion_PIEZAS, FECHA_DEBE_ATENDER, A_DOMICILIO, ATENDIDO, ID_USUARIO, ID_CLIENTE, FECHA_CAPTURA, TIPO_ARTICULO, MODELO, MARCA, COMENTARIOS_TECNICOS, NOMBRE, TELEFONO) VALUES ('" & LblMenu2.Caption & "', '" & GTIA & "', '" & TxtDesPiez.Text & "', '" & DtPFechAsi.Value & "', '" & DOMI & "', 0, '" & VarMen.Text1(0).Text & "', '" & IdClien & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & TxtTipoArt.Text & "', '" & TxtModelo.Text & "', '" & TxtMarcaArt.Text & "', '" & TxtComTec.Text & "', '" & TxtNomClienAs.Text & "', '" & Text14.Text & "');"
            cnn.Execute (sqlComanda)
            Dim tRs As ADODB.Recordset
            sqlComanda = "SELECT ID_AS_TEC FROM ASISTENCIA_TECNICA ORDER BY ID_AS_TEC DESC"
            Set tRs = cnn.Execute(sqlComanda)
            NoAsTec = tRs.Fields("ID_AS_TEC")
            'Imprimir
            FunImpATec
            Me.DtPFechAsi.Value = Format(Date, "dd/mm/yyyy")
        Else
            MsgBox "FALTA INFORMACION NESESARIA PARA EL REGISTRO", vbCritical, "SACC"
        End If
    Else
        MsgBox "DEBE SELECCIONAR UN CLIENTE PARA REGISTRAR", vbCritical, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunImpATec()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE ID_AS_TEC = " & NoAsTec & ""
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\AsTec.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "No. Asistencia : " & tRs.Fields("ID_AS_TEC"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs.Fields("FECHA_CAPTURA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        'sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        'Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 585, "CLIENTE : " & tRs.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 585, "TELEFONO : " & tRs.Fields("TELEFONO"), "F3", 8, hLeft
        'If Not (tRs.EOF And tRs.BOF) Then
        '    If Not IsNull(tRs.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs.Fields("NOMBRE"), "F3", 8, hCenter
        '    If Not IsNull(tRs.Fields("TELEFONO")) Then oDoc.WTextBox 120, 20, 100, 400, tRs.Fields("TELEFONO"), "F3", 8, hCenter
        'End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 10, 50, "MODELO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 10, 80, "MARCA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 10, 280, "TIPO DE ARTICULO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 415, 10, 60, "GARANTIA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 475, 10, 50, "DOMICILIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 525, 10, 65, "F. COMPROMISO", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 20, 50, tRs.Fields("MODELO"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 80, tRs.Fields("MARCA"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 20, 280, tRs.Fields("TIPO_ARTICULO"), "F3", 8, hCenter, , , 1, vbCyan
        If tRs.Fields("GARANTIA") = "1" Then
            oDoc.WTextBox Posi, 415, 20, 60, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 415, 20, 60, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        If tRs.Fields("A_DOMICILIO") = "1" Then
            oDoc.WTextBox Posi, 475, 20, 50, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 475, 20, 50, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        oDoc.WTextBox Posi, 525, 20, 65, Format(tRs.Fields("FECHA_DEBE_ATENDER"), "dd/mm/yyyy"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20
        oDoc.WTextBox Posi, 5, 10, 585, "DESCRIPCION DE PIEZAS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("DESCRIPCION_PIEZAS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        oDoc.WTextBox Posi, 5, 10, 585, "COMENTARIOS TECNICOS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("COMENTARIOS_TECNICOS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontrò la asistencia tècnica solicitada", vbExclamation, "SACC"
    End If
End Sub
Private Sub Imprimir()
On Error GoTo ManejaError
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "SUCURSAL : " & LblMenu2.Caption
    Printer.Print "TELEFONO SUCURSAL : " & VarMen.Text4(5).Text
    Printer.Print "No. DE ASISTENCIA : " & NoAsTec
    Printer.Print "ATENDIDO POR : " & LblMenu.Caption
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                       ASISTENCIA TECNICA"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "Cliente : " & TxtNomClienAs.Text
    Printer.Print "Telefono : " & Text14.Text & " o " & TelTrabajo
    Printer.Print "Calle : " & Direc & " # " & NoExte & "-" & NoInte
    Printer.Print "Colonia : " & COLONIA
    Printer.Print "Fecha a atender : " & DtPFechAsi.Value
    Printer.Print ""
    Printer.Print "Marca : " & TxtModelo.Text
    Printer.Print "Modelo : " & TxtMarcaArt.Text
    Printer.Print "Decripción : " & Mid(TxtDesPiez.Text, 1, 30) & "-"
    If Len(TxtDesPiez.Text) > 30 Then
        Printer.Print "Decripcion : " & Mid(TxtDesPiez.Text, 31, 70) & "-"
        If Len(TxtDesPiez.Text) > 70 Then
            Printer.Print "Decripcion : " & Mid(TxtDesPiez.Text, 71, 111) & "-"
        End If
    End If
    Printer.Print "Comentarios : " & Mid(TxtComTec.Text, 1, 30)
    If Len(TxtComTec.Text) > 30 Then
        Printer.Print "Decripcion : " & Mid(TxtComTec.Text, 31, 70) & "-"
        If Len(TxtComTec.Text) > 70 Then
            Printer.Print "Decripcion : " & Mid(TxtComTec.Text, 71, 111) & "-"
        End If
    End If
    Printer.Print "Articulo : " & Mid(TxtTipoArt.Text, 1, 30)
    If Len(TxtTipoArt.Text) > 30 Then
        Printer.Print "Decripcion : " & Mid(TxtTipoArt.Text, 31, 70) & "-"
        If Len(TxtTipoArt.Text) > 70 Then
            Printer.Print "Decripcion : " & Mid(TxtTipoArt.Text, 71, 111) & "-"
        End If
    End If
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
    Printer.Print "     DESPUES DE HABER EFECTUADO SU "
    Printer.Print "                                COMPRA"
    Printer.Print "SI NO RECOGE SU EQUIPO DESPUES DE 15 "
    Printer.Print "DIAS DE FINALIZADO EL SERVICIO, LA "
    Printer.Print "EMPRESA NO SE HACE RESPONSABLE POR "
    Printer.Print "                               EXTRABIO"
    Printer.Print "                APLICA RESTRICCIONES"
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
    Finalizar
    MsgBox "ASITENCIA REGISTRADA!", vbInformation, "SACC"
    Finalizar
    Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtMarcaArt_LostFocus()
    TxtMarcaArt.BackColor = &H80000005
End Sub
Private Sub TxtModelo_GotFocus()
    TxtModelo.BackColor = &HFFE1E1
    TxtModelo.SetFocus
    TxtModelo.SelStart = 0
    TxtModelo.SelLength = Len(TxtModelo.Text)
End Sub
Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMarcaArt.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtMarcaArt_GotFocus()
    TxtMarcaArt.BackColor = &HFFE1E1
    TxtMarcaArt.SetFocus
    TxtMarcaArt.SelStart = 0
    TxtMarcaArt.SelLength = Len(TxtMarcaArt.Text)
End Sub
Private Sub TxtMarcaArt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDesPiez.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++ < Domicilios > ++++++++++++++++++++++++++++++++
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub CmbColonia_DropDown()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    CmbColonia.Clear
    sBuscar = "SELECT NOMBRE FROM COLONIAS ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then
                CmbColonia.AddItem tRs.Fields("NOMBRE")
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub BtnGuardaDomi_Click()
On Error GoTo ManejaError
    If TxtNomCliente.Text = "" Or TxtDomiCleinte.Text = "" Or CmbColonia.Text = "" Or TxtNoArticulos.Text = "" Then
        MsgBox "FALTA INFORMACION NECESARIA", vbInformation, "SACC"
    Else
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT ZONA FROM COLONIAS WHERE NOMBRE = '" & CmbColonia.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If (tRs.EOF And tRs.BOF) Then
            sBuscar = "INSERT INTO DOMICILIOS (NUM_ARTICULOS, NOM_CLIENTE, DOMICILIO, COLONIA, TELEFONO, FECHA, DE_HORA, A_HORA, NOTA, ZONA, ESTADO, FECHA_ALTA, USUARIO) VALUES ('" & TxtNoArticulos.Text & "', '" & TxtNomCliente.Text & "', '" & TxtDomiCleinte.Text & "', '" & CmbColonia.Text & "','" & TxtTelefonoDomi.Text & "', '" & DTPFechaDomi.Value & "', '" & TxtHoraDe.Text & "', '" & TxtHoraAl.Text & "', '" & TxtNotaDomi.Text & "', 'DESC', 'P', SYSDATETIME(), '" & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text & "');"
        Else
            sBuscar = "INSERT INTO DOMICILIOS (NUM_ARTICULOS, NOM_CLIENTE, DOMICILIO, COLONIA, TELEFONO, FECHA, DE_HORA, A_HORA, NOTA, ZONA, ESTADO, FECHA_ALTA, USUARIO) VALUES ('" & TxtNoArticulos.Text & "', '" & TxtNomCliente.Text & "', '" & TxtDomiCleinte.Text & "', '" & CmbColonia.Text & "','" & TxtTelefonoDomi.Text & "', '" & DTPFechaDomi.Value & "', '" & TxtHoraDe.Text & "', '" & TxtHoraAl.Text & "', '" & TxtNotaDomi.Text & "', '" & tRs.Fields("ZONA") & "', 'P', SYSDATETIME(), '" & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text & "');"
        End If
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_DOMICILIO FROM DOMICILIOS ORDER BY ID_DOMICILIO DESC"
        Set tRs = cnn.Execute(sBuscar)
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "TELEFONO SUCURSAL : " & VarMen.Text4(5).Text
        Printer.Print "No. DE COMANDA : " & tRs.Fields("ID_DOMICILIO") & "-DOMI"
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & VarMen.Text1(2).Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                       RECOLECCION A DOMICILIO"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print "CLIENTE :"
        Printer.Print Mid(TxtNomCliente.Text, 1, 25) & "-"
        If Len(TxtNomCliente.Text) > 26 Then
            Printer.Print Mid(TxtNomCliente.Text, 26, 24) & "-"
        End If
        If Len(TxtNomCliente.Text) > 51 Then
            Printer.Print Mid(TxtNomCliente.Text, 51, 24) & "-"
        End If
        If Len(TxtNomCliente.Text) > 76 Then
            Printer.Print Mid(TxtNomCliente.Text, 76, 24) & "-"
        End If
        Printer.Print "DOMICILIO :"
        Printer.Print Mid(TxtDomiCleinte.Text, 1, 25) & "-"
        If Len(TxtDomiCleinte.Text) > 26 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 26, 24) & "-"
        End If
        If Len(TxtDomiCleinte.Text) > 51 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 51, 24) & "-"
        End If
        If Len(TxtDomiCleinte.Text) > 76 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 76, 24) & "-"
        End If
        Printer.Print "COLONIA : " & CmbColonia.Text
        Printer.Print "TELEFONO : " & TxtTelefonoDomi.Text
        Printer.Print "FECHA : " & DTPFechaDomi.Value
        Printer.Print "ENTRE LAS " & TxtHoraDe.Text & " Y LAS " & TxtHoraAl.Text
        Printer.Print "RECOGER " & TxtNoArticulos.Text & " ARTICULOS"
        Printer.Print "NOTAS :"
        Printer.Print Mid(TxtNotaDomi.Text, 1, 25) & "-"
        If Len(TxtNotaDomi.Text) > 26 Then
            Printer.Print Mid(TxtNotaDomi.Text, 26, 24) & "-"
        End If
        If Len(TxtNotaDomi.Text) > 51 Then
            Printer.Print Mid(TxtNotaDomi.Text, 51, 24) & "-"
        End If
        If Len(TxtNotaDomi.Text) > 76 Then
            Printer.Print Mid(TxtNotaDomi.Text, 76, 24) & "-"
        End If
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    End If
    Finalizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub BtnNueColonia_Click()
    FrmAgrColonia.Show vbModal
End Sub
Private Sub TxtModelo_LostFocus()
    TxtModelo.BackColor = &H80000005
End Sub
Private Sub TxtNoArticulos_LostFocus()
    TxtNoArticulos.BackColor = &H80000005
End Sub
Private Sub TxtNomClienAs_GotFocus()
    TxtNomClienAs.BackColor = &HFFE1E1
End Sub
Private Sub TxtNomClienAs_LostFocus()
    TxtNomClienAs.BackColor = &H80000005
End Sub
Private Sub TxtNomCliente_Change()
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
End Sub
Private Sub TxtNomCliente_GotFocus()
    TxtNomCliente.BackColor = &HFFE1E1
    TxtNomCliente.SelStart = 0
    TxtNomCliente.SelLength = Len(TxtNomCliente.Text)
End Sub
Private Sub TxtNomCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDomiCleinte.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtDomiCleinte_Change()
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
End Sub
Private Sub TxtDomiCleinte_GotFocus()
    TxtDomiCleinte.BackColor = &HFFE1E1
    TxtDomiCleinte.SelStart = 0
    TxtDomiCleinte.SelLength = Len(TxtDomiCleinte.Text)
End Sub
Private Sub TxtDomiCleinte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbColonia.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub CmbColonia_GotFocus()
    CmbColonia.BackColor = &HFFE1E1
    CmbColonia.SelStart = 0
    CmbColonia.SelLength = Len(CmbColonia.Text)
End Sub
Private Sub TxtHoraDe_GotFocus()
    TxtHoraDe.BackColor = &HFFE1E1
    TxtHoraDe.SelStart = 0
    TxtHoraDe.SelLength = Len(TxtHoraDe.Text)
End Sub
Private Sub TxtHoraDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtHoraAl.SetFocus
    End If
    If KeyAscii = 58 And Len(TxtHoraDe.Text) = 1 Then
        TxtHoraDe.Text = "0" & TxtHoraDe.Text
        TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtHoraDe.Text) = 1 And Val(TxtHoraDe.Text) > 2 Then
                TxtHoraDe.Text = "0" & TxtHoraDe.Text
                TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
        Else
            If Len(TxtHoraDe.Text) = 2 Then
                TxtHoraDe.Text = TxtHoraDe.Text & ":"
                TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
            End If
        End If
    End If
    Dim Valido As String
    If Len(TxtHoraDe.Text) = 1 Then
        If TxtHoraDe.Text = "2" Then
            Valido = "12340"
        Else
            Valido = "1234567890"
        End If
    End If
    If Len(TxtHoraDe.Text) = 3 Then
        Valido = "123450"
    End If
    If Len(TxtHoraDe.Text) = 4 Or Len(TxtHoraDe.Text) = 0 Then
        Valido = "1234567890"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtHoraDe_LostFocus()
    TxtHoraDe.BackColor = &H80000005
    If Len(TxtHoraDe.Text) = 1 Then
        TxtHoraDe.Text = "0" & TxtHoraDe.Text & ":00"
    End If
    If Len(TxtHoraDe.Text) = 2 Then
        TxtHoraDe.Text = TxtHoraDe.Text & ":00"
    End If
    If Len(TxtHoraDe.Text) = 3 Then
        TxtHoraDe.Text = TxtHoraDe.Text & "00"
    End If
    If Len(TxtHoraDe.Text) = 4 Then
        TxtHoraDe.Text = TxtHoraDe.Text & "0"
    End If
End Sub
Private Sub TxtHoraAl_GotFocus()
    TxtHoraAl.BackColor = &HFFE1E1
    TxtHoraAl.SetFocus
    TxtHoraAl.SelStart = 0
    TxtHoraAl.SelLength = Len(TxtHoraAl.Text)
End Sub
Private Sub TxtHoraAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNotaDomi.SetFocus
    End If
    If KeyAscii = 58 And Len(TxtHoraAl.Text) = 1 Then
        TxtHoraAl.Text = "0" & TxtHoraAl.Text
        TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtHoraAl.Text) = 1 And Val(TxtHoraAl.Text) > 2 Then
                TxtHoraAl.Text = "0" & TxtHoraAl.Text
                TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
        Else
            If Len(TxtHoraAl.Text) = 2 Then
                TxtHoraAl.Text = TxtHoraAl.Text & ":"
                TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
            End If
        End If
    End If
    Dim Valido As String
    If Len(TxtHoraAl.Text) = 1 Then
        If TxtHoraAl.Text = "2" Then
            Valido = "12340"
        Else
            Valido = "1234567890"
        End If
    End If
    If Len(TxtHoraAl.Text) = 3 Then
        Valido = "123450"
    End If
    If Len(TxtHoraAl.Text) = 4 Or Len(TxtHoraAl.Text) = 0 Then
        Valido = "1234567890"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtHoraAl_LostFocus()
    TxtHoraAl.BackColor = &H80000005
    If Len(TxtHoraAl.Text) = 1 Then
        TxtHoraAl.Text = "0" & TxtHoraAl.Text & ":00"
    End If
    If Len(TxtHoraAl.Text) = 2 Then
        TxtHoraAl.Text = TxtHoraAl.Text & ":00"
    End If
    If Len(TxtHoraAl.Text) = 3 Then
        TxtHoraAl.Text = TxtHoraAl.Text & "00"
    End If
    If Len(TxtHoraAl.Text) = 4 Then
        TxtHoraAl.Text = TxtHoraAl.Text & "0"
    End If
End Sub
Private Sub TxtNomCliente_LostFocus()
    TxtNomCliente.BackColor = &H80000005
End Sub
Private Sub TxtNotaDomi_GotFocus()
    TxtNotaDomi.BackColor = &HFFE1E1
    TxtNotaDomi.SetFocus
    TxtNotaDomi.SelStart = 0
    TxtNotaDomi.SelLength = Len(TxtNotaDomi.Text)
End Sub
Private Sub TxtNotaDomi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.BtnGuardaDomi.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtNotaDomi_LostFocus()
    TxtNotaDomi.BackColor = &H80000005
End Sub
Private Sub txtProductoComanda_LostFocus()
    txtProductoComanda.BackColor = &H80000005
End Sub
Private Sub TxtTelefonoDomi_Change()
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
End Sub
Private Sub TxtTelefonoDomi_GotFocus()
    TxtTelefonoDomi.BackColor = &HFFE1E1
    TxtTelefonoDomi.SelStart = 0
    TxtTelefonoDomi.SelLength = Len(TxtTelefonoDomi.Text)
End Sub
Private Sub TxtTelefonoDomi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 47 And Len(TxtTelefonoDomi.Text) = 1 Then
        TxtTelefonoDomi.Text = "0" & TxtTelefonoDomi.Text
        TxtTelefonoDomi.SelStart = Len(TxtTelefonoDomi.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtTelefonoDomi.Text) = 3 Or Len(TxtTelefonoDomi.Text) = 6 Then
            TxtTelefonoDomi.Text = TxtTelefonoDomi.Text & "-"
            TxtTelefonoDomi.SelStart = Len(TxtTelefonoDomi.Text)
        End If
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        TxtNoArticulos.SetFocus
    End If
End Sub
Private Sub TxtNoArticulos_Change()
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
    If TxtNoArticulos.Text = "" Then
        TxtNoArticulos.Text = 0
    End If
End Sub
Private Sub TxtNoArticulos_GotFocus()
    TxtNoArticulos.BackColor = &HFFE1E1
    TxtNoArticulos.SelStart = 0
    TxtNoArticulos.SelLength = Len(TxtNoArticulos.Text)
End Sub
Private Sub TxtNoArticulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtHoraDe.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Finalizar()
    Text2.Text = ""
    TxtModelo.Text = ""
    TxtMarcaArt.Text = ""
    TxtDesPiez.Text = ""
    TxtTipoArt.Text = ""
    TxtComTec.Text = ""
    chk1.Value = 0
    chk2.Value = 0
    Check2.Value = 0
    Me.Command2.Enabled = False
    ListView3.ListItems.Clear
    ClvProd = ""
    DesProd = ""
    PreProd = ""
    CLVCLIEN = ""
    NomClien = ""
    DesClien = ""
    DelInd = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = "0.00"
    Text4.Text = "0.00"
    Text5.Text = "0.00"
    Text15.Text = "0.00"
    Text16.Text = "0.00"
    Text17.Text = "0.00"
    Text10.Text = ""
    IdBenta.Text = ""
    Combo2.Text = ""
    Text9.Text = ""
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    Text1.Enabled = True
    ListView1.Enabled = True
    Text1.SetFocus
    TxtNomCliente.Text = ""
    TxtDomiCleinte.Text = ""
    TxtHoraDe.Text = ""
    TxtHoraAl.Text = ""
    TxtNotaDomi.Text = ""
    TxtTelefonoDomi.Text = ""
    TxtNoArticulos.Text = ""
    CmbColonia.Text = ""
    LlenaCombo
End Sub
Private Sub cmdAceptarComanda_Click()
On Error GoTo ManejaError
    If Text12.Text <> "" And Text13.Text <> "" Then
        If Puede_Guardar Then
            Dim NoRe As Integer
            Dim Cont As Integer
            Dim nComanda As Integer
            Dim cTipo As String
            'Hora del sistema.
            sqlQuery = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_CLIENTE, ID_AGENTE, ID_SUCURSAL, NOMBRE, TELEFONO, SUCURSAL) VALUES (SYSDATETIME(), " & Me.txtId_Cliente.Text & ", " & VarMen.Text1(0).Text & ", " & VarMen.Text1(5).Text & ", '" & Text12.Text & "', '" & Text13.Text & "', '" & VarMen.Text4(0).Text & "')"
            cnn.Execute (sqlQuery)
            Me.lblEstado.Caption = "Enviando"
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            sqlQuery = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
            Set tRs = cnn.Execute(sqlQuery)
            nComanda = tRs.Fields("ID_COMANDA")
            Me.lblEstado.Caption = Me.lblEstado.Caption & " comanda " & nComanda
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            NoRe = Me.lvwNuevaComanda.ListItems.Count
            For Cont = 1 To NoRe
                If Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "T" Then
                    cTipo = "T" 'Toner
                ElseIf Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "I" Then
                    cTipo = "I" 'Tinta
                Else
                    cTipo = "X" 'Error
                End If
                sqlQuery = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO, CLASIFICACION) VALUES (" & nComanda & ", " & Cont & ", '" & Me.lvwNuevaComanda.ListItems.Item(Cont) & "', " & Me.lvwNuevaComanda.ListItems.Item(Cont).SubItems(2) & ", '" & cTipo & "','C')"
                cnn.Execute (sqlQuery)
                
                Me.lblEstado.Caption = Me.lblEstado.Caption & ", producto " & Cont & " de " & NoRe
                Me.lblEstado.ForeColor = vbBlack
                DoEvents
            Next Cont
            Imprimir_Ticket (nComanda)
            Imprimir_Ticket (nComanda)
            Borrar_Campos
        End If
        Text3.Text = "0.00"
        Text4.Text = "0.00"
        Text5.Text = "0.00"
        Finalizar
    Else
        MsgBox "FALTA INFORMACIÓN DEL CLIENTE!", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdAgregarComanda_Click()
    If Puede_Agregar_Comanda Then
        Set tLi = Me.lvwNuevaComanda.ListItems.Add(, , Me.lvwProductosComanda.SelectedItem)
        tLi.SubItems(1) = Me.lvwProductosComanda.SelectedItem.SubItems(1)
        tLi.SubItems(2) = Me.txtCantidadComanda.Text
        Me.lblEstado.Caption = ""
        Me.txtProductoComanda.SetFocus
    End If
End Sub
Private Sub cmdQuitarComanda_Click()
    If Me.lvwNuevaComanda.ListItems.Count <> 0 Then
        If Me.lvwNuevaComanda.SelectedItem.Selected Then
            Me.lvwNuevaComanda.ListItems.Remove (Me.lvwNuevaComanda.SelectedItem.Index)
        End If
    End If
End Sub
Private Sub lvwProductosComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.lvwProductosComanda.SelectedItem.Selected Then
            Me.txtProductoComanda.Text = Me.lvwProductosComanda.SelectedItem
            Me.txtCantidadComanda.SetFocus
        End If
    End If
End Sub
Private Sub txtCantidadComanda_GotFocus()
    txtCantidadComanda.BackColor = &HFFE1E1
    Me.txtCantidadComanda.SelStart = 0
    Me.txtCantidadComanda.SelLength = Len(Me.txtCantidadComanda.Text)
End Sub
Private Sub txtCantidadComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.cmdAgregarComanda.Value = True
    Else
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Function Puede_Buscar_Producto() As Boolean
    If Trim(Me.txtProductoComanda.Text) = "" Then
        Puede_Buscar_Producto = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    Puede_Buscar_Producto = True
End Function
Function Hay_Productos(cProducto As String) As Boolean
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    If Me.opnClaveComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    ElseIf Me.opnDescripcionComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE A.Descripcion LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    Else
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION = '" & cProducto & "' ORDER BY J.ID_REPARACION"
    End If
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not .EOF Then
            Hay_Productos = True
            Me.lblEstado.Caption = ""
            Me.lblEstado.ForeColor = vbBlue
            DoEvents
        Else
            Hay_Productos = False
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtProductoComanda.SetFocus
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub txtProductoComanda_GotFocus()
    txtProductoComanda.BackColor = &HFFE1E1
    Me.txtProductoComanda.SelStart = 0
    Me.txtProductoComanda.SelLength = Len(Me.txtProductoComanda.Text)
End Sub
Private Sub txtProductoComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lvwProductosComanda.SetFocus
        If Puede_Buscar_Producto Then
            Me.lblEstado.Caption = "Buscando"
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            Llenar_Lista_Productos Trim(Me.txtProductoComanda.Text)
        End If
    End If
End Sub
Function Llenar_Lista_Productos(cProducto As String)
On Error GoTo ManejaError
    If Me.opnClaveComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    ElseIf Me.opnDescripcionComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE A.Descripcion LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    Else
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION = '" & cProducto & "' ORDER BY J.ID_REPARACION"
    End If
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            Me.lblEstado.Caption = ""
            Me.lvwProductosComanda.ListItems.Clear
            Do While Not .EOF
                Set tLi = lvwProductosComanda.ListItems.Add(, , .Fields("ID_REPARACION"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                If Not IsNull(.Fields("GANANCIA")) Then tLi.SubItems(2) = .Fields("GANANCIA")
                If Not IsNull(.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                If Not IsNull(.Fields("PRECIO_COSTO")) And Not IsNull(.Fields("GANANCIA")) Then
                    tLi.SubItems(4) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "0.00")
                Else
                    MsgBox "BASE DE DATOS CORRUPTA", vbCritical, "ERROR GRAVE"
                End If
                .MoveNext
            Loop
        Else
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtProductoComanda.SetFocus
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Puede_Agregar_Comanda() As Boolean
    If Me.lvwProductosComanda.ListItems.Count = 0 Then
        Puede_Agregar_Comanda = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    If Trim(Val(Me.txtCantidadComanda.Text)) = 0 Then
        Puede_Agregar_Comanda = False
        Me.lblEstado.Caption = "Introsusca la cantidad"
        Me.lblEstado.ForeColor = vbRed
        Me.txtCantidadComanda.SetFocus
        Exit Function
    End If
    Puede_Agregar_Comanda = True
End Function
Function Puede_Guardar() As Boolean
    If Me.txtId_Cliente.Text = "" Then
        Puede_Guardar = False
        Me.lblEstado.Caption = "Introsusca el cliente"
        Me.lblEstado.ForeColor = vbRed
        Exit Function
    End If
    If Me.lvwNuevaComanda.ListItems.Count = 0 Then
        Puede_Guardar = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    Puede_Guardar = True
End Function
Function Borrar_Campos()
    Me.txtCantidadComanda.Text = "1"
    Me.txtProductoComanda.Text = ""
    Me.lvwProductosComanda.ListItems.Clear
    Me.lvwNuevaComanda.ListItems.Clear
    Me.lblEstado.Caption = ""
    txtId_Cliente.Text = ""
    Text12.Text = ""
    Text13.Text = ""
End Function
Function Imprimir_Ticket(cNoCom As Integer)
On Error GoTo ManejaError
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "FECHA : " & Now
    Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "TELEFONO SUCURSAL : " & VarMen.Text4(5).Text
    Printer.Print "No. DE COMANDA : " & cNoCom
    Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "CLIENTE : " & TxtNomCliente.Text
    Printer.Print "CONTACTO : " & Text12.Text
    Printer.Print "TELEFONO : " & Text13.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           RECARGA DE TINTA"
    Dim NRegistros As Integer
    NRegistros = Me.lvwNuevaComanda.ListItems.Count
    Dim Con As Integer
    Dim POSY As Integer
    POSY = 2600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For Con = 1 To NRegistros
        If Mid(Me.lvwNuevaComanda.ListItems.Item(Con), 3, 1) = "I" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(Con).SubItems(2)
        End If
    Next Con
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           RECARGA DE TONER"
    POSY = POSY + 600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For Con = 1 To NRegistros
        If Mid(Me.lvwNuevaComanda.ListItems.Item(Con), 3, 1) <> "I" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(Con).SubItems(2)
        End If
    Next Con
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print "SI SU PEDIDO NO SE RECOGE EN 30 DIAS LA"
    Printer.Print "   EMPRESA NO SE HACE RESPONSABLE"
    Printer.Print "                APLICA RESTRICCIONES"
    Printer.Print ""
    Printer.Print "Conserve su ticket"
    Printer.Print "El cobro se hará hasta la entrega del cartucho lleno"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub TxtTelefonoDomi_LostFocus()
    TxtTelefonoDomi.BackColor = &H80000005
End Sub
Private Sub TxtTipoArt_GotFocus()
    TxtTipoArt.BackColor = &HFFE1E1
End Sub
Private Sub TxtTipoArt_LostFocus()
    TxtTipoArt.BackColor = &H80000005
End Sub
Private Sub ExtraerAsistencia()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim PrecioTot As String
    Dim ClienteSi As String
    sBuscar = "SELECT ATENDIDO, ID_CLIENTE FROM ASISTENCIA_TECNICA WHERE ID_AS_TEC = " & Text10.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
        Set tRs1 = cnn.Execute(sBuscar)
        If tRs1.EOF And tRs1.BOF Then
            MsgBox "DEBE SELECCIONAR UN CLIENTE!", vbInformation, "SACC"
        Else
            ClienteSi = "S"
            CLVCLIEN = tRs.Fields("ID_CLIENTE")
            Text1.Text = tRs1.Fields("NOMBRE")
            If Not IsNull(tRs1.Fields("NOMBRE")) Then NomClien = tRs1.Fields("NOMBRE")
            If Not IsNull(tRs1.Fields("DESCUENTO")) Then DesClien = tRs1.Fields("DESCUENTO")
            If Not IsNull(tRs1.Fields("DIAS_CREDITO")) Then Combo2.Text = tRs1.Fields("DIAS_CREDITO")
            If Not IsNull(tRs1.Fields("LIMITE_CREDITO")) Then Text9.Text = tRs1.Fields("LIMITE_CREDITO")
        End If
        If tRs.Fields("ATENDIDO") = 3 Then
            sBuscar = "SELECT * FROM COBRO_ASISTENCIA_TECNICA WHERE ID_AS_TEC = " & Text10.Text
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    sBuscar = "SELECT DESCRIPCION, PRECIO_COSTO, GANANCIA, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                    Set tRs1 = cnn.Execute(sBuscar)
                    Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    If Not IsNull(tRs1.Fields("Descripcion")) Then tLi.SubItems(1) = tRs1.Fields("Descripcion")
                    If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                    tLi.SubItems(3) = Format(CDbl(tRs1.Fields("PRECIO_COSTO")) * CDbl((tRs1.Fields("GANANCIA") + 1)), "0.00")
                    tLi.SubItems(4) = "A" & Text10.Text
                    Text3.Text = Format((CDbl(Text3.Text) + CDbl(tRs1.Fields("PRECIO_COSTO")) * CDbl((tRs1.Fields("GANANCIA") + 1)) * CDbl(tRs.Fields("CANTIDAD"))), "0.00")
                    PrecioTot = Format((CDbl(tRs1.Fields("PRECIO_COSTO")) * CDbl((tRs1.Fields("GANANCIA") + 1))) * tRs.Fields("CANTIDAD"), "0.00")
                    If RFC = "XEXX010101000" Then
                        Text4.Text = "0.00"
                        Text17.Text = "0.00"
                        Text15.Text = "0.00"
                        Text16.Text = "0.00"
                        Text5.Text = Format(CDbl(Text3.Text), "0.00")
                    Else
                        Text4.Text = Format(CDbl(Text4.Text) + (CDbl(PrecioTot) * CDbl(tRs1.Fields("IVA"))), "0.00")
                        Text17.Text = Format(CDbl(Text17.Text) + (CDbl(PrecioTot) * CDbl(tRs1.Fields("P_RETENCION"))), "0.00")
                        Text15.Text = Format(CDbl(Text15.Text) + (CDbl(PrecioTot) * CDbl(tRs1.Fields("IMPUESTO1"))), "0.00")
                        Text16.Text = Format(CDbl(Text16.Text) + (CDbl(PrecioTot) * CDbl(tRs1.Fields("IMPUESTO2"))), "0.00")
                        Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) - CDbl(Text17.Text) + CDbl(Text15.Text) + CDbl(Text16.Text), "0.00")
                    End If
                    tLi.SubItems(5) = Format(CDbl(PrecioTot) * CDbl(tRs1.Fields("IVA")), "0.00")
                    tLi.SubItems(6) = Format(CDbl(PrecioTot) * CDbl(tRs1.Fields("P_RETENCION")), "0.00")
                    tLi.SubItems(7) = Format(CDbl(PrecioTot) * CDbl(tRs1.Fields("IMPUESTO1")), "0.00")
                    tLi.SubItems(8) = Format(CDbl(PrecioTot) * CDbl(tRs1.Fields("IMPUESTO2")), "0.00")
                    tRs.MoveNext
                    If ClienteSi = "S" Then
                        Command2.Enabled = True
                    End If
                Loop
                sBuscar = "UPDATE ASISTENCIA_TECNICA SET ATENDIDO = 1 WHERE ID_AS_TEC = " & Text10.Text
                Set tRs = cnn.Execute(sBuscar)
            Else
                MsgBox "LA ASISTENCIA HA SIDO FINALIZADA SIN COBRO PARA EL CLIENTE!", vbInformation, "SACC"
                sBuscar = "UPDATE ASISTENCIA_TECNICA SET ATENDIDO = 1 WHERE ID_AS_TEC = " & Text10.Text
                Set tRs = cnn.Execute(sBuscar)
            End If
        Else
            If tRs.Fields("ATENDIDO") = 1 Then
                MsgBox "LA ASISTENCIA YA FUE COBRADA!", vbInformation, "SACC"
            Else
                If tRs.Fields("ATENDIDO") = 0 Then
                    MsgBox "LA ASISTENCIA AUN NO HA SIDO FINALIZADA!", vbInformation, "SACC"
                End If
            End If
        End If
    End If
End Sub
Private Sub VerClienteaDetalle_Click()
    FrmVerClien.Show vbModal
End Sub
