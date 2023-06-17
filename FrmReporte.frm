VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporte 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Cuentas Por Cobrar"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmReporte.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9000
      TabIndex        =   57
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   56
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   54
      Top             =   5880
      Width           =   975
      Begin VB.Label Label21 
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
         TabIndex        =   55
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmReporte.frx":AF62
         MousePointer    =   99  'Custom
         Picture         =   "FrmReporte.frx":B26C
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   9120
      TabIndex        =   53
      Top             =   3720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9000
      TabIndex        =   46
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
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
      Height          =   195
      Left            =   9360
      Picture         =   "FrmReporte.frx":CDAE
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   28
      Top             =   7080
      Width           =   975
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmReporte.frx":F780
         MousePointer    =   99  'Custom
         Picture         =   "FrmReporte.frx":FA8A
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   21
      Top             =   8280
      Width           =   975
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmReporte.frx":1165C
         MousePointer    =   99  'Custom
         Picture         =   "FrmReporte.frx":11966
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   16536
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Reporte General"
      TabPicture(0)   =   "FrmReporte.frx":13A48
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Reporte Detallado"
      TabPicture(1)   =   "FrmReporte.frx":13A64
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame11"
      Tab(1).Control(1)=   "Check3"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(6)=   "ListView1"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "A Fecha"
      TabPicture(2)   =   "FrmReporte.frx":13A80
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView4"
      Tab(2).Control(1)=   "Check1"
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(3)=   "Frame14"
      Tab(2).Control(4)=   "Command3"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command3 
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
         Left            =   -67440
         Picture         =   "FrmReporte.frx":13A9C
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame Frame14 
         Caption         =   "Totales Generales Del Detallado"
         Height          =   855
         Left            =   -74880
         TabIndex        =   75
         Top             =   8280
         Width           =   8535
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   78
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   77
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   76
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   5760
            TabIndex        =   81
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Total de Abonos"
            Height          =   255
            Left            =   3000
            TabIndex        =   80
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Total de Compra"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Rango de Fechas"
         Height          =   975
         Left            =   -74760
         TabIndex        =   69
         Top             =   600
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   480
            TabIndex        =   70
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   375
            Left            =   2280
            TabIndex        =   71
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label19 
            Caption         =   "De :"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Al :"
            Height          =   255
            Left            =   1920
            TabIndex        =   72
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Solo Facturados"
         Height          =   195
         Left            =   -70800
         TabIndex        =   68
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Frame Frame11 
         Caption         =   "Morosos"
         Height          =   1575
         Left            =   -72720
         TabIndex        =   60
         Top             =   480
         Width           =   1215
         Begin VB.OptionButton Option7 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "45 Días"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "15 Días"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "5 Días"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "30 Días"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Solo Facturados"
         Height          =   195
         Left            =   -71400
         TabIndex        =   58
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Frame Frame9 
         Caption         =   "Totales Generales Del Detallado"
         Height          =   855
         Left            =   -74880
         TabIndex        =   49
         Top             =   8400
         Width           =   8535
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Total de Compra"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Total de Abonos"
            Height          =   255
            Left            =   3000
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   5760
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Reporte"
         Height          =   1575
         Left            =   -67800
         TabIndex        =   45
         Top             =   480
         Width           =   1455
         Begin VB.OptionButton Option12 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Abonos"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Pendientes"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Rango de Fechas"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   2055
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label15 
            Caption         =   "Al :"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "De :"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cliente"
         Height          =   1215
         Left            =   -71400
         TabIndex        =   43
         Top             =   480
         Width           =   3495
         Begin VB.CommandButton cmdBuscar 
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
            Left            =   2280
            Picture         =   "FrmReporte.frx":1646E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Totales Generales"
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   4500
         Width           =   8535
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   5760
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Total de Abonos"
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Total de Compra"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Totales Del Detalle"
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   8040
         Width           =   8535
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6720
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Total de Compra"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Total de Abonos"
            Height          =   255
            Left            =   3000
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   5880
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9600
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   240
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "CLIENTE"
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   420
         Width           =   5295
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
            Left            =   4080
            Picture         =   "FrmReporte.frx":18E40
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   5640
         TabIndex        =   23
         Top             =   420
         Width           =   3015
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Top             =   5520
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4471
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
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   1740
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4895
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   19
         Top             =   2160
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   10821
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
         Height          =   6495
         Left            =   -74880
         TabIndex        =   74
         Top             =   1680
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   11456
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
      Begin VB.Label Label3 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   -2760
         TabIndex        =   26
         Top             =   -360
         Width           =   735
      End
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   9000
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Pe"
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Total de Abonos"
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private cnn As ADODB.Connection
Dim Total As Double
Dim Abono As Double
Dim Total2 As Double
Dim Abono2 As Double
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim pendiente As Double
    ListView4.ListItems.Clear
    ListView3.ListItems.Clear
    ListView2.ListItems.Clear
    ListView1.ListItems.Clear
    Text11.Text = "0.00"
    Text12.Text = "0.00"
    Text13.Text = "0.00"
    Frame15.Visible = True
    If Option2.Value Then
        Total = 0
        Abono = 0
        pendiente = 0
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        With ListView1
            .View = lvwReport
            .GridLines = True
            .LabelEdit = lvwManual
            .HideSelection = False
            .HotTracking = False
            .HoverSelection = False
            .ColumnHeaders.Add , , "Venta", 500
            .ColumnHeaders.Add , , "Nombre", 1500
            .ColumnHeaders.Add , , "Folio", 1200
            .ColumnHeaders.Add , , "Fecha de Venta", 1200
            .ColumnHeaders.Add , , "Fecha de Vencimiento", 1200
            .ColumnHeaders.Add , , "Subtotal", 1500
            .ColumnHeaders.Add , , "IVA", 1500
            .ColumnHeaders.Add , , "Total", 1500
            .ColumnHeaders.Add , , "Abono", 1500
            .ColumnHeaders.Add , , "Pendiente", 1500
            .ColumnHeaders.Add , , "Fecha de Pago", 1500
            .ColumnHeaders.Add , , "Dias Vencidos", 1500
            .ColumnHeaders.Add , , "Estatus", 1500
            .ColumnHeaders.Add , , "Fecha Factura", 1500
        End With
        pendientes
    Else
        If Option3.Value Then
            ListView1.ListItems.Clear
            ListView1.ColumnHeaders.Clear
            With ListView1
                .View = lvwReport
                .GridLines = True
                .LabelEdit = lvwManual
                .HideSelection = False
                .HotTracking = False
                .HoverSelection = False
                .ColumnHeaders.Add , , "Id Cliente", 500
                .ColumnHeaders.Add , , "Nombre", 1500
                .ColumnHeaders.Add , , "Folio", 1200
                .ColumnHeaders.Add , , "Fecha", 1200
                .ColumnHeaders.Add , , "Fecha Vence", 1200
                .ColumnHeaders.Add , , "Abono", 1600
                .ColumnHeaders.Add , , "Efectivo", 800
                .ColumnHeaders.Add , , "No De Cheque", 1000
                .ColumnHeaders.Add , , "Referencia", 1000
            End With
            If Not IsNumeric(Text9.Text) Then
                sBuscar = "SELECT ID_CLIENTE, FOLIO, NOMBRE, FECHA, BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA, FECHA_VENCE FROM VSABONOS WHERE (NOMBRE LIKE '%" & Text9.Text & "%') AND (FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "') OR (FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "') AND (NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%') ORDER BY ID_CLIENTE, ID_VENTA"
            Else
                sBuscar = "SELECT ID_CLIENTE, FOLIO, NOMBRE, FECHA, BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA, FECHA_VENCE FROM VSABONOS WHERE (NOMBRE LIKE '%" & Text9.Text & "%') AND (FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "') OR (FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "') AND (NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%') OR (FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "') AND (ID_CLIENTE LIKE '%" & Text9.Text & "%') ORDER BY ID_CLIENTE, ID_VENTA"
            End If
            Set tRs = cnn.Execute(sBuscar)
            StrRep = sBuscar
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
                    If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
                    If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(4) = tRs.Fields("FECHA_VENCE")
                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then tLi.SubItems(5) = Format(tRs.Fields("CANT_ABONO"), "0.00")
                    If Not IsNull(tRs.Fields("EFECTIVO")) Then tLi.SubItems(6) = tRs.Fields("EFECTIVO")
                    If Not IsNull(tRs.Fields("NO_CHEQUE")) Then tLi.SubItems(7) = tRs.Fields("NO_CHEQUE")
                    If Not IsNull(tRs.Fields("REFERENCIA")) Then tLi.SubItems(8) = tRs.Fields("REFERENCIA")
                    Text12.Text = CDbl(Text12.Text) + CDbl(tRs.Fields("TOT_ABONOS"))
                    Text11.Text = CDbl(Text11.Text) + CDbl(tRs.Fields("TOTAL"))
                    Text13.Text = CDbl(Text11.Text) - CDbl(Text12.Text)
                    'Total2 = Format(CDbl(Total2) + CDbl(tRs.Fields("CANT_ABONO")), "###,###,##0.00")
                    Text12.Text = Total2
                    tRs.MoveNext
                Loop
            End If
        Else
            ListView1.ListItems.Clear
            ListView1.ColumnHeaders.Clear
            With ListView1
                .View = lvwReport
                .GridLines = True
                .LabelEdit = lvwManual
                .HideSelection = False
                .HotTracking = False
                .HoverSelection = False
                .ColumnHeaders.Add , , "Id Cliente", 500
                .ColumnHeaders.Add , , "Nombre", 1500
                .ColumnHeaders.Add , , "Folio", 1200
                .ColumnHeaders.Add , , "Fecha", 1200
                .ColumnHeaders.Add , , "Fecha Vence", 1200
                .ColumnHeaders.Add , , "Estatus", 1600
                .ColumnHeaders.Add , , "Subtotal", 800
                .ColumnHeaders.Add , , "IVA", 1000
                .ColumnHeaders.Add , , "Total", 1000
                .ColumnHeaders.Add , , "Total Abonos", 1000
            End With
            If Not IsNumeric(Text9.Text) Then
                sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, VSCXC3.FECHA_VENCE, VSCXC3.PAGADA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, IIF(VSCXC3.PAGADA = 'S',  SUM(VSCXC3.TOTAL) , ISNULL(SUM(ABONOS_CUENTA.CANT_ABONO), 0)) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA ON VSCXC3.ID_CUENTA = ABONOS_CUENTA.ID_CUENTA WHERE VSCXC3.NOMBRE LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR  VSCXC3.NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA ORDER BY VSCXC3.NOMBRE, VSCXC3.ID_VENTA"
            Else
                sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, VSCXC3.FECHA_VENCE, VSCXC3.PAGADA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, IIF(VSCXC3.PAGADA = 'S',  SUM(VSCXC3.TOTAL) , ISNULL(SUM(ABONOS_CUENTA.CANT_ABONO), 0)) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA ON VSCXC3.ID_CUENTA = ABONOS_CUENTA.ID_CUENTA WHERE VSCXC3.NOMBRE LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR  VSCXC3.NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR VSCXC3.ID_CLIENTE = '" & Text9.Text & "' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "'" & _
                "GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA ORDER BY VSCXC3.NOMBRE, VSCXC3.ID_VENTA"
            End If
            Set tRs = cnn.Execute(sBuscar)
            StrRep = sBuscar
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
                    If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
                    If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(4) = Format(tRs.Fields("FECHA_VENCE"), "dd/mm/yyyy")
                    If tRs.Fields("PAGADA") = "S" Then
                        tLi.SubItems(5) = "PAGADA"
                    Else
                        tLi.SubItems(5) = "PENDIENTE"
                    End If
                    If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(6) = Format(tRs.Fields("SUBTOTAL"), "0.00")
                    If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(7) = Format(tRs.Fields("IVA"), "0.00")
                    If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(8) = Format(tRs.Fields("TOTAL"), "0.00")
                    If Not IsNull(tRs.Fields("TOT_ABONOS")) Then tLi.SubItems(9) = Format(tRs.Fields("TOT_ABONOS"), "0.00")
                    'Total2 = Format(CDbl(Total2) + CDbl(tRs.Fields("TOT_ABONOS")), "###,###,##0.00")
                    Text12.Text = CDbl(Text12.Text) + CDbl(tRs.Fields("TOT_ABONOS"))
                    Text11.Text = CDbl(Text11.Text) + CDbl(tRs.Fields("TOTAL"))
                    Text13.Text = CDbl(Text11.Text) - CDbl(Text12.Text)
                    tRs.MoveNext
                Loop
            End If
        End If
    End If
    Text11.Text = Format(Text11.Text, "###,###,###,##0.00")
    Text12.Text = Format(Text12.Text, "###,###,###,##0.00")
    Text13.Text = Format(Text13.Text, "###,###,###,##0.00")
End Sub
Private Sub detalles()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    If Option2.Value Then
        If Not IsNumeric(Text1.Text) Then
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL FROM VSCXC3 WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' GROUP BY  NOMBRE, ID_CLIENTE, FOLIO, FECHA"
        Else
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL FROM VSCXC3 WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR ID_CLIENTE = '" & Text1.Text & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' GROUP BY  NOMBRE, ID_CLIENTE, FOLIO, FECHA"
        End If
        Set tRs = cnn.Execute(sBuscar)
        StrRep = sBuscar
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("FOLIO")
                tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
                tLi.SubItems(4) = tRs.Fields("TOTAL")
                sBuscar = "SELECT ID_CLIENTE,FOLIO,SUM(CANT_ABONO) AS CANT_ABONO FROM temporal_abonos WHERE FOLIO NOT IN ('CANCELADO') AND FOLIO='" & tRs.Fields("FOLIO") & "'   AND ID_CLIENTE='" & tRs.Fields("ID_CLIENTE") & "'  GROUP BY ID_CLIENTE,FOLIO"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    Do While Not (tRs1.EOF)
                        tLi.SubItems(5) = tRs1.Fields("CANT_ABONO")
                        tRs1.MoveNext
                    Loop
                    tLi.SubItems(6) = CDbl(tLi.SubItems(4)) - CDbl(tLi.SubItems(5))
                End If
                tRs.MoveNext
            Loop
        End If
    Else
        If Not IsNumeric(Text1.Text) Then
            sBuscar = "SELECT ID_CLIENTE, FOLIO, NOMBRE, FECHA,BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA FROM VSABONOS WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR FOLIO NOT IN ('CANCELADO') AND NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY NOMBRE"
        Else
            sBuscar = "SELECT ID_CLIENTE, FOLIO, NOMBRE, FECHA,BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA FROM VSABONOS WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR FOLIO NOT IN ('CANCELADO') AND NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' OR FOLIO NOT IN ('CANCELADO') AND ID_CLIENTE = '" & Text1.Text & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY NOMBRE"
        End If
        Set tRs = cnn.Execute(sBuscar)
        StrRep = sBuscar
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("FOLIO")
                tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
                tLi.SubItems(4) = tRs.Fields("FECHA_VENCE")
                tLi.SubItems(5) = tRs.Fields("CANT_ABONO")
                tLi.SubItems(6) = tRs.Fields("EFECTIVO")
                tLi.SubItems(7) = tRs.Fields("NO_CHEQUE")
                tLi.SubItems(8) = tRs.Fields("REFERENCIA")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    If Not IsNumeric(Text1.Text) Then
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, ID_DESCUENTO ,SUM(TOTAL) AS TOTAL FROM VSCXC4 WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND ID_DESCUENTO = '" & Combo1.Text & "' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND ID_DESCUENTO = '" & Combo1.Text & "' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' GROUP BY NOMBRE, ID_CLIENTE, ID_DESCUENTO"
    Else
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, ID_DESCUENTO ,SUM(TOTAL) AS TOTAL FROM VSCXC4 WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND ID_DESCUENTO = '" & Combo1.Text & "' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND ID_DESCUENTO = '" & Combo1.Text & "' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' OR ID_CLIENTE = '" & Text1.Text & "' AND ID_DESCUENTO = '" & Combo1.Text & "' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' GROUP BY NOMBRE, ID_CLIENTE, ID_DESCUENTO"
    End If
    StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(3) = tRs.Fields("id_descuento")
            tLi.SubItems(4) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sNombre As String
    Dim tot As Double
    Dim abo   As Double
    Dim Pend  As Double
    Text3.Text = "0.00"
    Text4.Text = "0.00"
    Text5.Text = "0.00"
    sNombre = Item
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    Frame15.Visible = True
    tot = 0
    abo = 0
    If Not IsNumeric(Text1.Text) Then
        sBuscar = "SELECT SUM(VENT.TOTAL) AS TOTAL, VENT.ID_CLIENTE, VENT.NOMBRE, ISNULL((SELECT SUM(CANT_ABONO) AS CANT_ABONO From temporal_abonos WHERE (FOLIO NOT IN ('CANCELADO')) AND (ID_CLIENTE = VENT.ID_CLIENTE) AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')), 0.00) AS CANT_ABONO FROM VENTAS AS VENT INNER JOIN CUENTA_VENTA ON VENT.ID_VENTA = CUENTA_VENTA.ID_VENTA INNER JOIN CUENTAS ON CUENTA_VENTA.ID_CUENTA = CUENTAS.ID_CUENTA INNER JOIN CLIENTE ON VENT.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (CUENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (CLIENTE.NOMBRE LIKE '%" & Text1.Text & "%') AND (VENT.FACTURADO IN (1, 2)) OR (CUENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (dbo.CLIENTE.NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%') AND (VENT.FACTURADO IN (1, 2)) GROUP BY VENT.NOMBRE, VENT.ID_CLIENTE ORDER BY VENT.NOMBRE"
    Else
        sBuscar = "SELECT SUM(VENT.TOTAL) AS TOTAL, VENT.ID_CLIENTE, VENT.NOMBRE, ISNULL((SELECT SUM(CANT_ABONO) AS CANT_ABONO From temporal_abonos WHERE (FOLIO NOT IN ('CANCELADO')) AND (ID_CLIENTE = VENT.ID_CLIENTE) AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')), 0.00) AS CANT_ABONO FROM VENTAS AS VENT INNER JOIN CUENTA_VENTA ON VENT.ID_VENTA = CUENTA_VENTA.ID_VENTA INNER JOIN CUENTAS ON CUENTA_VENTA.ID_CUENTA = CUENTAS.ID_CUENTA INNER JOIN CLIENTE ON VENT.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (CUENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (CLIENTE.NOMBRE LIKE '%" & Text1.Text & "%') AND (VENT.FACTURADO IN (1, 2)) OR (CUENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (dbo.CLIENTE.NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%') AND (VENT.FACTURADO IN (1, 2)) OR GROUP BY VENT.NOMBRE, VENT.ID_CLIENTE ORDER BY VENT.NOMBRE" & _
        " (CUENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (dbo.CLIENTE.ID_PRODUCTO LIKE '%" & Text1.Text & "%') AND (VENT.FACTURADO IN (1, 2)) GROUP BY VENT.NOMBRE, VENT.ID_CLIENTE ORDER BY VENT.NOMBRE"
    End If
    'sBuscar = "SELECT NOMBRE,ID_CLIENTE,SUM(TOTAL) AS TOTAL FROM VSCXC WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' GROUP BY  NOMBRE, ID_CLIENTE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            tot = Format(CDbl(tot) + CDbl(tRs.Fields("TOTAL")), "###,###,##0.00")
            tLi.SubItems(3) = Format(tRs.Fields("CANT_ABONO"), "###,###,##0.00")
            abo = Format(CDbl(abo) + CDbl(tRs.Fields("CANT_ABONO")), "###,###,##0.00")
            tLi.SubItems(4) = Format(CDbl(tLi.SubItems(2)) - CDbl(tLi.SubItems(3)), "###,###,##0.00")
            tRs.MoveNext
        Loop
    End If
    Text3.Text = tot
    Text4.Text = abo
    Text5.Text = Format(CDbl(Text3.Text) - CDbl(Text4.Text), "###,###,##0.00")
    Text1.SetFocus
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim pendiente As Double
    ListView4.ListItems.Clear
    ListView3.ListItems.Clear
    ListView2.ListItems.Clear
    ListView1.ListItems.Clear
    Text19.Text = "0.00"
    Text18.Text = "0.00"
    Text17.Text = "0.00"
    ListView4.ListItems.Clear
    If Check1.Value = 1 Then
        sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, ISNULL ((SELECT SUM(CANT_ABONO) AS TOTAL From ABONOS_CUENTA WHERE (VSCXC3.FACTURADO = 1) AND (ID_VENTA = VSCXC3.ID_VENTA) AND (FECHA <= '" & DTPicker6.Value & "')), 0) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA AS ABONOS_CUENTA_1 ON VSCXC3.ID_CUENTA = ABONOS_CUENTA_1.ID_CUENTA WHERE (VSCXC3.FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA, VSCXC3.FACTURADO ORDER BY VSCXC3.NOMBRE"
    Else
        sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, ISNULL ((SELECT SUM(CANT_ABONO) AS TOTAL From ABONOS_CUENTA WHERE (VSCXC3.FACTURADO IN (0,1)) AND (ID_VENTA = VSCXC3.ID_VENTA) AND (FECHA <= '" & DTPicker6.Value & "')), 0) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA AS ABONOS_CUENTA_1 ON VSCXC3.ID_CUENTA = ABONOS_CUENTA_1.ID_CUENTA WHERE (VSCXC3.FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA, VSCXC3.FACTURADO ORDER BY VSCXC3.NOMBRE"
    End If
    Set tRs = cnn.Execute(sBuscar)
    StrRep = sBuscar
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("ID_VENTA")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(5) = tRs.Fields("SUBTOTAL")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(6) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(7) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("TOT_ABONOS")) Then tLi.SubItems(8) = tRs.Fields("TOT_ABONOS")
            Text19.Text = Format(CDbl(Text19.Text) + CDbl(tRs.Fields("TOTAL")), "###,###,##0.00")
            Text18.Text = Format(CDbl(Text18.Text) + CDbl(tRs.Fields("TOT_ABONOS")), "###,###,##0.00")
            tRs.MoveNext
        Loop
    End If
    Text17.Text = CDbl(Text19.Text) - CDbl(Text18.Text)
    Text19.Text = Format(Text19.Text, "###,###,###,##0.00")
    Text18.Text = Format(Text18.Text, "###,###,###,##0.00")
    Text17.Text = Format(Text17.Text, "###,###,###,##0.00")
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    DTPicker3.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker4.Value = Format(Date, "dd/mm/yyyy")
    DTPicker5.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker6.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 500
        .ColumnHeaders.Add , , "Nombre", 1500
        .ColumnHeaders.Add , , "Folio", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Abono", 1500
        .ColumnHeaders.Add , , "Pendiente", 1500
        .ColumnHeaders.Add , , "Fecha Vencimiento", 1500
        .ColumnHeaders.Add , , "Dias vencidos", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 500
        .ColumnHeaders.Add , , "Nombre", 3000
        .ColumnHeaders.Add , , "Total de Compra", 1500
        .ColumnHeaders.Add , , "Total de Abonos", 1500
        .ColumnHeaders.Add , , "Pendiente de Pago", 1500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 500
        .ColumnHeaders.Add , , "Nombre", 3000
        .ColumnHeaders.Add , , "Folio", 1500
        .ColumnHeaders.Add , , "Venta", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "IVA", 1500
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Abonos", 1500
    End With
    sBuscar = "SELECT ID_DESCUENTO FROM DESCUENTOS ORDER BY ID_DESCUENTO"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("ID_DESCUENTO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub CUENTAS()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim Abonos As Double
    Dim Total1 As Double
    Dim Total2 As Double
    Dim total3 As Double
    Dim ConPag As Integer
    Dim sDias As Double
    Dim sIdCliente As String
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\reportecuentas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Cuentas Detallado)", "F2", 10, hCenter
    oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
    oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
    oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
    oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 110, 10, 30, 40, "Nota", "F2", 10, hCenter
    oDoc.WTextBox 110, 50, 50, 60, "Fecha Nota", "F2", 10, hCenter
    oDoc.WTextBox 110, 110, 30, 60, "Factura", "F2", 10, hCenter
    oDoc.WTextBox 110, 170, 50, 60, "Fecha Fact", "F2", 10, hCenter
    oDoc.WTextBox 110, 230, 40, 70, "Total-Fac", "F2", 10, hCenter
    oDoc.WTextBox 110, 300, 40, 70, "Abono", "F2", 10, hCenter
    oDoc.WTextBox 110, 370, 40, 70, "Pendiente", "F2", 10, hLeft
    oDoc.WTextBox 110, 440, 40, 60, "Estatus", "F2", 10, hLeft
    oDoc.WTextBox 110, 500, 40, 70, "Dias Vence", "F2", 10, hLeft
' Cuerpo del reporte
    deuda = 0
    deuda1 = 0
    deuda2 = 0
    totor = 0
    totpr = 0
    Conta = 0
    Total1 = 0
    Total2 = 0
    total3 = 0
    If Option2.Value Then
        If Check3.Value = 1 Then
            If Not IsNumeric(Text9.Text) Then
                sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' AND FACTURADO = 1 OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' AND FACTURADO = 1 GROUP BY NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, ID_VENTA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
            Else
                sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' AND FACTURADO = 1 OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' AND FACTURADO = 1 OR ID_CLIENTE = '" & Text9.Text & "' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' AND FACTURADO = 1 GROUP BY NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, ID_VENTA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
            End If
        Else
            If Option4.Value Then
                sDias = 5
            Else
                If Option5.Value Then
                    sDias = 15
                Else
                    If Option1.Value Then
                        sDias = 30
                    Else
                        If Option6.Value Then
                            sDias = 45
                        End If
                    End If
                End If
            End If
            If Not IsNumeric(Text9.Text) Then
                If Option7.Value Then
                    sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' GROUP BY ID_VENTA, NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
                Else
                    sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FECHA_VENCE >= '" & Date + sDias & "' OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FECHA_VENCE >= '" & Date + sDias & "' GROUP BY ID_VENTA, NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
                End If
            Else
                If Option7.Value Then
                    sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR ID_CLIENTE = '" & Text9.Text & "' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' GROUP BY ID_VENTA, NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
                Else
                    sBuscar = "SELECT ID_VENTA, NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FECHA_VENCE >= '" & Date + sDias & "' OR NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FECHA_VENCE >= '" & Date + sDias & "' OR ID_CLIENTE = '" & Text9.Text & "' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FECHA_VENCE >= '" & Date + sDias & "' GROUP BY ID_VENTA, NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE, ID_VENTA"
                End If
            End If
        End If
        Set tRs = cnn.Execute(sBuscar)
    End If
    Posi = 140
    sumdeuda = 0
    sumIndi = 0
    Pend = 0
    sumabonos = 0
    sumor = 0
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 110
    oDoc.WLineTo 580, 110
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 580, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If sNombre <> tRs.Fields("NOMBRE") Or sIdCliente <> tRs.Fields("ID_CLIENTE") Then
                If sumor > 0 Then
                    Posi = Posi + 13
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 230, Posi
                    oDoc.WLineTo 290, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 230, 40, 60, Format(sumor, "###,###,##0.00"), "F3", 10, hRight
                    sumor = 0
                    Posi = Posi + 15
                End If
                If sumabonos > 0 Then
                    Posi = Posi - 22
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 260, Posi
                    oDoc.WLineTo 310, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 250, 40, 60, Format(sumabonos, "###,###,##0.00"), "F3", 10, hRight
                    sumabonos = 0
                    Posi = Posi + 15
                End If
                If deuda > 0 Then
                    Posi = Posi - 22
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 360, Posi
                    oDoc.WLineTo 410, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 350, 40, 60, Format(deuda, "###,###,##0.00"), "F3", 10, hRight
                    deuda = 0
                    Posi = Posi + 15
                End If
                oDoc.WTextBox Posi, 20, 9, 500, tRs.Fields("ID_CLIENTE") & "  -  " & tRs.Fields("NOMBRE"), "F3", 10, hLeft
            End If
            Posi = Posi + 15
            If Not IsNull(tRs.Fields("ID_VENTA")) Then oDoc.WTextBox Posi, 10, 30, 40, tRs.Fields("ID_VENTA"), "F3", 10, hLeft
            If Not IsNull(tRs.Fields("FECHA")) Then oDoc.WTextBox Posi, 50, 50, 60, Format(tRs.Fields("FECHA"), "dd/mm/yyyy"), "F3", 10, hCenter
            If Not IsNull(tRs.Fields("FOLIO")) Then oDoc.WTextBox Posi, 110, 30, 60, tRs.Fields("FOLIO"), "F3", 10, hCenter
            If Not IsNull(tRs.Fields("FECHA")) Then oDoc.WTextBox Posi, 170, 50, 60, Format(tRs.Fields("FECHA"), "dd/mm/yyyy"), "F3", 10, hCenter
            If Not IsNull(tRs.Fields("TOTAL")) Then oDoc.WTextBox Posi, 230, 40, 70, Format(tRs.Fields("TOTAL"), "###,###,##0.00"), "F3", 10, hRight
            Total1 = Format(CDbl(Total1) + CDbl(tRs.Fields("TOTAL")), "###,###,##0.00")
            'If Option2.Value Then
            If Date - tRs.Fields("FECHA_VENCE") < 0 Then
                oDoc.WTextBox Posi, 500, 40, 70, "0", "F3", 10, hCenter
            Else
                If Not IsNull(tRs.Fields("FECHA_VENCE")) Then oDoc.WTextBox Posi, 500, 40, 70, Date - tRs.Fields("FECHA_VENCE"), "F3", 10, hCenter
            End If
            If tRs.Fields("PAGADA") = "S" Then
                oDoc.WTextBox Posi, 440, 40, 60, "PAGADA", "F3", 10, hLeft
            End If
            If tRs.Fields("pagada") = "N" Then
                oDoc.WTextBox Posi, 440, 40, 60, "PENDIENTE", "F3", 10, hLeft
            End If
            PosVer = Posi
            sumor = Format(CDbl(sumor) + CDbl(tRs.Fields("TOTAL")), "0.00")
            sumpr = Format(CDbl(sumpr) + CDbl(tRs.Fields("TOTAL")), "0.00")
            If Option3.Value Then
                oDoc.WTextBox Posi, 300, 40, 70, Format(tRs.Fields("CANT_ABONO"), "###,###,##0.00"), "F3", 9, hRight
                Pend = Format(CDbl(tRs.Fields("TOTAL")) - CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                sumabonos = Format(CDbl(sumabonos) + CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                totpr = Format(CDbl(totpr) + CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                Abonos = Abonos + CDbl(tRs.Fields("CANT_ABONO"))
            Else
                sBuscar = "SELECT ID_CLIENTE, FOLIO, SUM(CANT_ABONO) AS CANT_ABONO FROM temporal_abonos WHERE FOLIO NOT IN ('CANCELADO') AND ID_VENTA = '" & tRs.Fields("ID_VENTA") & "' AND ID_CLIENTE = '" & tRs.Fields("ID_CLIENTE") & "'  GROUP BY ID_CLIENTE, FOLIO"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    Do While Not (tRs1.EOF)
                        oDoc.WTextBox Posi, 300, 40, 70, Format(tRs1.Fields("CANT_ABONO"), "###,###,##0.00"), "F3", 9, hRight
                        Pend = Format(CDbl(tRs.Fields("TOTAL")) - CDbl(tRs1.Fields("CANT_ABONO")), "###,###,##0.00")
                        tRs1.MoveNext
                    Loop
                Else
                    oDoc.WTextBox Posi, 300, 40, 70, "0.00", "F3", 9, hRight
                    Pend = Format(tRs.Fields("TOTAL"), "0.00")
                End If
            End If
            deuda = Format(CDbl(deuda) + CDbl(Pend), "0.00")
            deuda1 = Format(CDbl(deuda1) + CDbl(Pend), "0.00")
            If Pend > 0 Then
                oDoc.WTextBox Posi, 370, 40, 70, Format(Pend, "###,###,##0.00"), "F3", 10, hRight
            Else
                oDoc.WTextBox Posi, 370, 40, 70, "0.00", "F3", 10, hRight
            End If
            Pend = 0
            sNombre = tRs.Fields("NOMBRE")
            sIdCliente = tRs.Fields("ID_CLIENTE")
            tRs.MoveNext
            If Posi >= 700 Then
                oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 140
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Cuentas Detallado)", "F2", 10, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
                oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
                oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 110, 10, 30, 40, "Nota", "F2", 10, hCenter
                oDoc.WTextBox 110, 50, 50, 60, "Fecha Nota", "F2", 10, hCenter
                oDoc.WTextBox 110, 110, 30, 60, "Factura", "F2", 10, hCenter
                oDoc.WTextBox 110, 170, 50, 60, "Fecha Fact", "F2", 10, hCenter
                oDoc.WTextBox 110, 230, 40, 60, "Total-Fac", "F2", 10, hCenter
                oDoc.WTextBox 110, 290, 40, 60, "Abono", "F2", 10, hCenter
                oDoc.WTextBox 110, 350, 40, 60, "Pendiente", "F2", 10, hLeft
                oDoc.WTextBox 110, 410, 40, 60, "Status", "F2", 10, hLeft
                oDoc.WTextBox 110, 470, 40, 70, "Dias Vence", "F2", 10, hLeft
                ' Cuerpo del reporte
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 110
                oDoc.WLineTo 580, 110
                oDoc.LineStroke
                oDoc.MoveTo 10, 125
                oDoc.WLineTo 580, 125
                oDoc.LineStroke
            End If
        Loop
        'pendientes
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 350, Posi
        oDoc.WLineTo 400, Posi
        oDoc.LineStroke
        Posi = Posi + 7
        oDoc.WTextBox Posi, 370, 40, 70, Format(deuda, "###,###,##0.00"), "F3", 10, hRight
        deuda = 0
        '////
        Posi = Posi + 15
        Posi = Posi + 30
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Compras ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(Total1, "###,###,##0.00"), "F2", 9, hRight 'Format(Text11.Text, "###,###,##0.00"), "F2", 9, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Abonos ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(Abonos, "###,###,##0.00"), "F2", 9, hRight 'Format(Text12.Text, "###,###,##0.00"), "F2", 9, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Deudas ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(Total1 - Abonos, "###,###,##0.00"), "F2", 9, hRight ' Format(Text13.Text, "###,###,##0.00"), "F2", 9, hRight
        Cont = Cont + 1
        oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Abonos()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim ConPag As Integer
    ConPag = 1
    If Not oDoc.PDFCreate(App.Path & "\RepCuentas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Abonos Detallado)", "F2", 10, hCenter
    oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
    oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
    oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
    oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 100, 20, 30, 40, "Folio", "F2", 10, hCenter
    oDoc.WTextBox 100, 60, 30, 80, "Fecha", "F2", 10, hCenter
    oDoc.WTextBox 100, 100, 50, 160, "Banco", "F2", 10, hCenter
    oDoc.WTextBox 100, 250, 50, 160, "Efectivo", "F2", 10, hLeft
    oDoc.WTextBox 100, 300, 50, 160, "Num-Cheque", "F2", 10, hLeft
    oDoc.WTextBox 100, 400, 50, 160, "Referencia", "F2", 10, hLeft
    oDoc.WTextBox 100, 480, 40, 200, "Abono", "F2", 10, hLeft
' Cuerpo del reporte
    totor = 0
    totpr = 0
    Conta = 0
    If Not IsNull(Text9.Text) Then
        sBuscar = "SELECT FOLIO, NOMBRE, FECHA, BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA FROM VSABONOS WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text9.Text & "%' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR FOLIO NOT IN ('CANCELADO') AND NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' ORDER BY NOMBRE"
    Else
        sBuscar = "SELECT FOLIO, NOMBRE, FECHA, BANCO, CANT_ABONO, EFECTIVO, NO_CHEQUE, REFERENCIA FROM VSABONOS WHERE FOLIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text9.Text & "%' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR FOLIO NOT IN ('CANCELADO') AND NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' OR FOLIO NOT IN ('CANCELADO') AND ID_CLIENTE = '" & Text9.Text & "' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " 'ORDER BY NOMBRE"
    End If
    Set tRs = cnn.Execute(sBuscar)
    Posi = 140
    sumdeuda = 0
    sumIndi = 0
    sumabonos = 0
    sumIndiabonos = 0
    sumor = 0
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 100
    oDoc.WLineTo 580, 100
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 580, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If sNombre <> tRs.Fields("NOMBRE") Then
                If sumor > 0 Then
                    Posi = Posi + 13
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 480, Posi
                    oDoc.WLineTo 530, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 400, 40, 140, Format(sumor, "###,###,###,##0.00"), "F3", 10, hRight
                    sumor = 0
                    Posi = Posi + 15
                End If
                oDoc.WTextBox Posi, 20, 9, 500, tRs.Fields("NOMBRE"), "F2", 9, hLeft
            End If
            Posi = Posi + 15
            oDoc.WTextBox Posi, 20, 30, 40, tRs.Fields("FOLIO"), "F3", 10, hLeft
            oDoc.WTextBox Posi, 60, 30, 80, Format(tRs.Fields("FECHA"), "dd/mm/yyyy"), "F3", 10, hCenter
            oDoc.WTextBox Posi, 100, 50, 160, tRs.Fields("BANCO"), "F3", 10, hCenter
            oDoc.WTextBox Posi, 250, 50, 160, tRs.Fields("EFECTIVO"), "F3", 10, hLeft
            oDoc.WTextBox Posi, 300, 50, 160, tRs.Fields("NO_CHEQUE"), "F3", 10, hLeft
            oDoc.WTextBox Posi, 400, 50, 160, tRs.Fields("REFERENCIA"), "F3", 10, hLeft
            oDoc.WTextBox Posi, 400, 40, 140, Format(tRs.Fields("CANT_ABONO"), "###,###,###,##0.00"), "F3", 10, hRight
            PosVer = Posi
            sumor = CDbl(sumor) + CDbl(tRs.Fields("CANT_ABONO"))
            sumpr = CDbl(sumpr) + CDbl(tRs.Fields("CANT_ABONO"))
            sNombre = tRs.Fields("NOMBRE")
            tRs.MoveNext
            If Posi >= 720 Then
                oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 140
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Abonos Detallado)", "F2", 10, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
                oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
                oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 100, 20, 30, 40, "Folio", "F2", 10, hCenter
                oDoc.WTextBox 100, 60, 30, 80, "Fecha", "F2", 10, hCenter
                oDoc.WTextBox 100, 100, 50, 160, "Banco", "F2", 10, hCenter
                oDoc.WTextBox 100, 250, 50, 160, "Efectivo", "F2", 10, hLeft
                oDoc.WTextBox 100, 300, 50, 160, "Num-Cheque", "F2", 10, hLeft
                oDoc.WTextBox 100, 400, 50, 160, "Referencia", "F2", 10, hLeft
                oDoc.WTextBox 100, 480, 40, 200, "Abono", "F2", 10, hLeft
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 100
                oDoc.WLineTo 580, 100
                oDoc.LineStroke
                oDoc.MoveTo 10, 125
                oDoc.WLineTo 580, 125
                oDoc.LineStroke
            End If
        Loop
        If sumor > 0 Then
            Posi = Posi + 13
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 480, Posi
            oDoc.WLineTo 530, Posi
            oDoc.LineStroke
            Posi = Posi + 7
            oDoc.WTextBox Posi, 400, 40, 140, Format(sumor, "###,###,###,##0.00"), "F3", 10, hRight
            sumor = 0
            Posi = Posi + 15
        End If
        Posi = Posi + 15
        oDoc.WTextBox Posi, 300, 10, 160, "Total", "F2", 12, hLeft
        oDoc.WTextBox Posi, 320, 10, 220, Format(sumpr, "$###,###,###,##0.00"), "F2", 12, hRight
        oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
        Cont = Cont + 1
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Image1_Click()
    If SSTab1.Tab = 0 Then
        If MsgBox("¿DESEA LA INFORMACION GLOBAL (SI), DETALLADA (N0)?  ", vbYesNo + vbCritical + vbDefaultButton1) = vbYes Then
            RepGeneral
        Else
            CXCDET
        End If
    End If
    If SSTab1.Tab = 1 Then
        If Option2.Value Then
            CUENTAS
        Else
            If Option3.Value Then
                Abonos
            Else
                CuentasAbonos
            End If
        End If
    End If
End Sub
Private Sub Image4_Click()
On Error GoTo ManejaError
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    Ruta = Me.CommonDialog1.FileName
    If Ruta <> "" Then
        If ListView1.ListItems.Count > 0 And SSTab1.Tab = 1 Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & " " & Chr(9)
            Next
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.Count
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & " " & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
        End If
        If ListView3.ListItems.Count > 0 And SSTab1.Tab = 0 Then
            NumColum = ListView3.ColumnHeaders.Count
            For Con = 1 To ListView3.ColumnHeaders.Count
                StrCopi = StrCopi & ListView3.ColumnHeaders(Con).Text & " " & Chr(9)
            Next
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView3.ListItems.Count
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView3.ListItems.Count
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & " " & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & " " & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
        Else
            If ListView2.ListItems.Count > 0 And SSTab1.Tab = 1 Then
                NumColum = ListView2.ColumnHeaders.Count
                For Con = 1 To ListView2.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & " " & Chr(9)
                Next
                ProgressBar1.Value = 0
                ProgressBar1.Visible = True
                ProgressBar1.Min = 0
                ProgressBar1.Max = ListView2.ListItems.Count
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView2.ListItems.Count
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & " " & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                    ProgressBar1.Value = Con
                Next
            End If
        End If
        If ListView4.ListItems.Count > 0 And SSTab2.Tab = 1 Then
                NumColum = ListView4.ColumnHeaders.Count
                For Con = 1 To ListView4.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView4.ColumnHeaders(Con).Text & " " & Chr(9)
                Next
                ProgressBar1.Value = 0
                ProgressBar1.Visible = True
                ProgressBar1.Min = 0
                ProgressBar1.Max = ListView4.ListItems.Count
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView4.ListItems.Count
                    StrCopi = StrCopi & ListView4.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & " " & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                    ProgressBar1.Value = Con
                Next
            End If
        'archivo TXT
        Dim foo As Integer
        foo = FreeFile
        Open Ruta For Output As #foo
        Print #foo, StrCopi ' Text16.Text
        Close #foo
    End If
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
    ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sNombre As String
    Dim tot As Double
    Dim abo   As Double
    Dim Pend  As Double
    Text6.Text = "0.00"
    Text7.Text = "0.00"
    Text8.Text = "0.00"
    Text10.Text = Item.SubItems(1)
    ListView2.ListItems.Clear
    sNombre = Item
    sBuscar = "SELECT NOMBRE, ID_CLIENTE, FOLIO, MIN(FECHA) as FECHA, SUM(TOTAL) AS TOTAL, MIN(FECHA_VENCE) AS FECHA_VENCE, PAGADA FROM VSCXC3 WHERE FOLIO NOT IN ('CANCELADO') AND ID_CLIENTE = '" & sNombre & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' GROUP BY  NOMBRE, ID_CLIENTE, FOLIO, PAGADA"
    Set tRs = cnn.Execute(sBuscar)
    tot = 0
    abo = 0
    Pend = 0
    StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("FOLIO")
            tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
            tLi.SubItems(4) = tRs.Fields("TOTAL")
            tot = CDbl(tot) + CDbl(tRs.Fields("TOTAL"))
            If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(7) = tRs.Fields("FECHA_VENCE")
            If tRs.Fields("PAGADA") = "S" Then
                tLi.SubItems(8) = 0
            Else
                If Not IsNull(tRs.Fields("FECHA_VENCE")) Then
                    If tRs.Fields("FECHA_VENCE") <= Date Then
                        tLi.SubItems(8) = "Vencido por " & Date - tRs.Fields("FECHA_VENCE") & " días"
                    Else
                        tLi.SubItems(8) = "Vence en " & tRs.Fields("FECHA_VENCE") - Date & " días"
                    End If
                End If
                If tLi.SubItems(8) < 0 Then
                    tLi.SubItems(8) = 0
                End If
            End If
            If Format(Date, "dd/mm/yyyy") > tRs.Fields("FECHA_VENCE") And tRs.Fields("PAGADA") <> "S" Then
                tLi.ForeColor = &HFF&
                tLi.ListSubItems.Item(1).ForeColor = &HFF&
                tLi.ListSubItems.Item(2).ForeColor = &HFF&
                tLi.ListSubItems.Item(3).ForeColor = &HFF&
                tLi.ListSubItems.Item(4).ForeColor = &HFF&
                tLi.ListSubItems.Item(5).ForeColor = &HFF&
                tLi.ListSubItems.Item(6).ForeColor = &HFF&
                tLi.ListSubItems.Item(7).ForeColor = &HFF&
            End If
            sBuscar = "SELECT ID_CLIENTE, FOLIO, SUM(CANT_ABONO) AS CANT_ABONO FROM temporal_abonos WHERE FOLIO NOT IN ('CANCELADO') AND FOLIO='" & tRs.Fields("FOLIO") & "'   AND ID_CLIENTE='" & tRs.Fields("ID_CLIENTE") & "'  GROUP BY ID_CLIENTE,FOLIO"
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                Do While Not (tRs1.EOF)
                    tLi.SubItems(5) = tRs1.Fields("CANT_ABONO")
                    tLi.SubItems(6) = CDbl(tRs.Fields("TOTAL")) - CDbl(tRs1.Fields("CANT_ABONO"))
                    abo = Format(CDbl(abo) + CDbl(tRs1.Fields("CANT_ABONO")), "###,###,##0.00")
                    tRs1.MoveNext
                Loop
                tLi.SubItems(6) = CDbl(tLi.SubItems(4)) - CDbl(tLi.SubItems(5))
            End If
            tRs.MoveNext
        Loop
    End If
    Text6.Text = tot
    Text7.Text = abo
    Text8.Text = Format(CDbl(tot) - CDbl(abo), "###,###,##0.00")
    Text1.SetFocus
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    ListView3.ListItems.Clear
    If KeyAscii = 13 Then
        Me.Command2.Value = True
    End If
End Sub
Private Sub RepGeneral()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Dim tRs2  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim AcumDeudas As String
    Dim NoRe As Integer
    Dim ConPag As Integer
    ConPag = 1
    AcumDeudas = "0.00"
    NoRe = Me.ListView3.ListItems.Count
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\CXCGEN.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Cuentas)", "F2", 10, hCenter
    If Option2.Value Then
        oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
        oDoc.WTextBox 70, 450, 20, 250, DTPicker1.Value, "F2", 8, hLeft
        oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
        oDoc.WTextBox 70, 530, 20, 250, DTPicker2.Value, "F2", 8, hLeft
    End If
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
    oDoc.WTextBox 100, 20, 30, 40, "ID : ", "F2", 9, hLeft
    oDoc.WTextBox 100, 60, 30, 80, "NOMBRE :", "F2", 9, hLeft
    oDoc.WTextBox 100, 300, 40, 200, "TOTAL DE COMPRAS", "F2", 9, hLeft
    oDoc.WTextBox 100, 400, 40, 300, "TOTAL DE ABONOS", "F2", 9, hLeft
    oDoc.WTextBox 100, 500, 40, 300, "TOTAL DE DEUDAS", "F2", 9, hLeft
' Encabezado de pagina
    Posi = 110
' Cuerpo del reporte
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 100
    oDoc.WLineTo 580, 100
    oDoc.LineStroke
    For Cont = 1 To NoRe
        'If CDbl(ListView3.ListItems.Item(Cont).SubItems(4)) >= 0.8 Then
            Posi = Posi + 10
            oDoc.WTextBox Posi, 20, 30, 40, ListView3.ListItems.Item(Cont), "F3", 9, hLeft
            oDoc.WTextBox Posi, 60, 40, 300, ListView3.ListItems.Item(Cont).SubItems(1), "F3", 9, hLeft
            Posi = Posi + 10
            oDoc.WTextBox Posi, 300, 30, 100, ListView3.ListItems.Item(Cont).SubItems(2), "F3", 9, hLeft
            oDoc.WTextBox Posi, 420, 30, 100, ListView3.ListItems.Item(Cont).SubItems(3), "F3", 9, hLeft
            oDoc.WTextBox Posi, 500, 30, 100, ListView3.ListItems.Item(Cont).SubItems(4), "F3", 9, hLeft
            Posi = Posi + 12
            ''''linea  primea
            oDoc.SetLineFormat 0.2, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, Posi
            oDoc.WLineTo 580, Posi
            oDoc.LineStroke
            If Posi >= 760 Then
                oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 110
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Cuentas)", "F2", 10, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                oDoc.WTextBox 100, 20, 30, 40, "ID : ", "F2", 9, hLeft
                oDoc.WTextBox 100, 60, 30, 80, "NOMBRE :", "F2", 9, hLeft
                oDoc.WTextBox 100, 300, 40, 200, "TOTAL DE COMPRAS", "F2", 9, hLeft
                oDoc.WTextBox 100, 400, 40, 300, "TOTAL DE ABONOS", "F2", 9, hLeft
                oDoc.WTextBox 100, 500, 40, 300, "TOTAL DE DEUDAS", "F2", 9, hLeft
                ' Encabezado de pagina
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 100
                oDoc.WLineTo 580, 100
                oDoc.LineStroke
            End If
        'End If
    Next
    Posi = Posi + 30
    oDoc.WTextBox Posi, 20, 40, 80, "Total De Compras ", "F2", 9, hLeft
    oDoc.WTextBox Posi, 110, 40, 80, Text3.Text, "F2", 9, hLeft
    oDoc.WTextBox Posi, 200, 40, 80, "Total De Abonos ", "F2", 9, hLeft
    oDoc.WTextBox Posi, 300, 40, 100, Text4.Text, "F2", 9, hLeft
    oDoc.WTextBox Posi, 400, 40, 80, "Total De Deudas ", "F2", 9, hLeft
    oDoc.WTextBox Posi, 480, 40, 100, Text5.Text, "F2", 9, hLeft
    oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
    Cont = Cont + 1
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub CXCDET()
    If ListView2.ListItems.Count > 0 Then
        Dim oDoc  As cPDF
        Dim dblX  As Double
        Dim dblY  As Double
        Dim Angle As Double
        Dim Cont As Integer
        Dim Posi As Integer
        Dim loca As Integer
        Dim sBuscar As String
        Dim tRs1 As ADODB.Recordset
        Dim PosVer As Integer
        Dim Posabo As Integer
        Dim tRs  As ADODB.Recordset
        Dim tRs2  As ADODB.Recordset
        Set oDoc = New cPDF
        Dim sumdeuda As Double
        Dim sumIndi As Double
        Dim sNombre As String
        Dim sumabonos As Double
        Dim sumIndiabonos As Double
        Dim AcumDeudas As String
        Dim NoRe As Integer
        Dim ConPag As Integer
        ConPag = 1
        AcumDeudas = "0.00"
        Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        NoRe = Me.ListView2.ListItems.Count
        If Not oDoc.PDFCreate(App.Path & "\CXCGEN.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        oDoc.LoadImage Image2, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 80, 40, 43, 161, "Logo"
        oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
        oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
        oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
        oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (FACTURAS)", "F2", 10, hCenter
        oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
        If Option2.Value Then
            oDoc.WTextBox 70, 450, 20, 250, DTPicker1.Value, "F2", 8, hLeft
            oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
            oDoc.WTextBox 70, 530, 20, 250, DTPicker2.Value, "F2", 8, hLeft
        End If
        oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
        oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
        ' Encabezado de pagina
        Posi = 110
        ' Cuerpo del reporte
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 100
        oDoc.WLineTo 580, 100
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 30, 300, "Cliente :", "F3", 9, hLeft
        Posi = Posi + 15
        oDoc.WTextBox Posi, 20, 30, 300, Text10.Text, "F3", 9, hLeft
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 30, 40, "FACTURA", "F3", 9, hLeft
        oDoc.WTextBox Posi, 80, 40, 100, "FECHA", "F3", 9, hLeft
        oDoc.WTextBox Posi, 150, 40, 100, "TOTAL", "F3", 9, hLeft
        oDoc.WTextBox Posi, 200, 30, 100, "ABONO", "F3", 9, hLeft
        oDoc.WTextBox Posi, 260, 30, 100, "PENDIENTE", "F3", 9, hLeft
        oDoc.WTextBox Posi, 320, 30, 200, "VENCIMIENTO", "F3", 9, hLeft
        oDoc.WTextBox Posi, 420, 30, 100, "DIAS VENCIDOS", "F3", 9, hLeft
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        For Cont = 1 To ListView3.ListItems.Count
            If CDbl(ListView2.ListItems.Item(Cont).SubItems(4)) >= 0.8 Then
                Posi = Posi + 10
                oDoc.WTextBox Posi, 20, 30, 40, ListView2.ListItems(Cont).SubItems(2), "F3", 9, hLeft
                oDoc.WTextBox Posi, 80, 40, 300, ListView2.ListItems.Item(Cont).SubItems(3), "F3", 9, hLeft
                oDoc.WTextBox Posi, 150, 30, 100, ListView2.ListItems.Item(Cont).SubItems(4), "F3", 9, hLeft
                oDoc.WTextBox Posi, 200, 30, 100, ListView2.ListItems.Item(Cont).SubItems(5), "F3", 9, hLeft
                oDoc.WTextBox Posi, 260, 30, 100, ListView2.ListItems.Item(Cont).SubItems(6), "F3", 9, hLeft
                oDoc.WTextBox Posi, 330, 30, 100, ListView2.ListItems.Item(Cont).SubItems(7), "F3", 9, hLeft
                oDoc.WTextBox Posi, 430, 30, 100, ListView2.ListItems.Item(Cont).SubItems(8), "F3", 9, hLeft
                If Posi >= 760 Then
                    oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    ' Encabezado del reporte
                    Posi = 110
                    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Cuentas)", "F2", 10, hCenter
                    oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                    oDoc.WTextBox 100, 20, 30, 40, "ID : ", "F2", 9, hLeft
                    oDoc.WTextBox 100, 60, 30, 80, "NOMBRE :", "F2", 9, hLeft
                    oDoc.WTextBox 100, 300, 40, 200, "RFC", "F2", 9, hLeft
                    oDoc.WTextBox 100, 400, 40, 300, "TOTAL DE COMPRAS", "F2", 9, hLeft
                    oDoc.WTextBox 100, 500, 40, 300, "TOTAL DE DEUDAS", "F2", 9, hLeft
                    ' Encabezado de pagina
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, 100
                    oDoc.WLineTo 580, 100
                    oDoc.LineStroke
                End If
            End If
        Next
        Posi = Posi + 30
        Posi = Posi + 30
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Compras ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 80, Text6.Text, "F2", 9, hLeft
        oDoc.WTextBox Posi, 200, 40, 80, "Total De Abonos ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 300, 40, 100, Text7.Text, "F2", 9, hLeft
        oDoc.WTextBox Posi, 400, 40, 80, "Total De Deudas ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 480, 40, 100, Text8.Text, "F2", 9, hLeft
        oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
        Cont = Cont + 1
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontraron resultados", vbExclamation, "SACC"
    End If
End Sub
Private Sub pendientes()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim sDias As Double
    If Option4.Value Then
        sDias = 5
    Else
        If Option5.Value Then
            sDias = 15
        Else
            If Option1.Value Then
                sDias = 30
            Else
                If Option6.Value Then
                    sDias = 45
                End If
            End If
        End If
    End If
    If Not IsNumeric(Text9.Text) Then
        If Check3.Value = 1 Then
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, SUM(SUBTOTAL) AS SUBTOTAL, SUM(IVA) AS IVA, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FACTURADO = 1"
        Else
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, SUM(SUBTOTAL) AS SUBTOTAL, SUM(IVA) AS IVA, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FACTURADO IN (1, 0)"
        End If
    Else
        If Check3.Value = 1 Then
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, SUM(SUBTOTAL) AS SUBTOTAL, SUM(IVA) AS IVA, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FACTURADO = 1"
        Else
            sBuscar = "SELECT NOMBRE, ID_CLIENTE, FECHA_VENCE, FOLIO, FECHA, SUM(TOTAL) AS TOTAL, SUM(SUBTOTAL) AS SUBTOTAL, SUM(IVA) AS IVA, PAGADA, FECHA_FACTURA FROM VSCXC3 WHERE NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND PAGADA='N' AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' AND FACTURADO IN (1, 0)"
        End If
    End If
    If Option7.Value Then
        sBuscar = sBuscar & " GROUP BY NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE"
    Else
        sBuscar = sBuscar & " AND FECHA_VENCE >= '" & Date + sDias & "' GROUP BY NOMBRE, ID_CLIENTE, FOLIO, FECHA, FECHA_VENCE, PAGADA, FECHA_FACTURA ORDER BY ID_CLIENTE"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = Format(tRs.Fields("FECHA"), "dd/mm/yyyy")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(5) = Format(tRs.Fields("SUBTOTAL"), "0.00")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(6) = Format(tRs.Fields("IVA"), "0.00")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(7) = Format(tRs.Fields("TOTAL"), "0.00")
            Total = Format(CDbl(Total) + CDbl(tRs.Fields("TOTAL")), "0.00")
            If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(4) = tRs.Fields("FECHA_VENCE")
            If tRs.Fields("PAGADA") = "S" Then
                tLi.SubItems(8) = 0
            Else
                If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(11) = Date - tRs.Fields("FECHA_VENCE")
                If IsNumeric(tLi.SubItems(11)) Then
                    If tLi.SubItems(11) < 0 Then
                        tLi.SubItems(11) = 0
                    End If
                Else
                    tLi.SubItems(11) = 0
                End If
            End If
            If tRs.Fields("pagada") = "S" Then
                tLi.SubItems(12) = "PAGADA"
            End If
            If tRs.Fields("pagada") = "N" Then
                tLi.SubItems(12) = "PENDIENTE"
            End If
            If Not IsNull(tRs.Fields("FECHA_FACTURA")) Then tLi.SubItems(13) = Format(tRs.Fields("FECHA_FACTURA"), "dd/MM/yyyy")
            If Format(Date, "dd/mm/yyyy") > tRs.Fields("FECHA_VENCE") And tRs.Fields("PAGADA") <> "S" Then
                tLi.ForeColor = &HFF&
                tLi.ListSubItems.Item(1).ForeColor = &HFF&
                tLi.ListSubItems.Item(2).ForeColor = &HFF&
                tLi.ListSubItems.Item(3).ForeColor = &HFF&
                tLi.ListSubItems.Item(4).ForeColor = &HFF&
                tLi.ListSubItems.Item(5).ForeColor = &HFF&
                tLi.ListSubItems.Item(6).ForeColor = &HFF&
                tLi.ListSubItems.Item(7).ForeColor = &HFF&
                tLi.ListSubItems.Item(8).ForeColor = &HFF&
                tLi.ListSubItems.Item(9).ForeColor = &HFF&
                tLi.ListSubItems.Item(10).ForeColor = &HFF&
                tLi.ListSubItems.Item(11).ForeColor = &HFF&
                tLi.ListSubItems.Item(12).ForeColor = &HFF&
                'tLi.ListSubItems.Item(13).ForeColor = &HFF&
            End If
            sBuscar = "SELECT ID_CLIENTE, FOLIO, SUM(CANT_ABONO) AS CANT_ABONO FROM temporal_abonos WHERE FOLIO NOT IN ('CANCELADO') AND FOLIO = '" & tRs.Fields("FOLIO") & "' AND ID_CLIENTE = " & tRs.Fields("ID_CLIENTE") & " GROUP BY ID_CLIENTE, FOLIO"
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                Do While Not (tRs1.EOF)
                    tLi.SubItems(8) = tRs1.Fields("CANT_ABONO")
                    Abono = Format(CDbl(Abono) + CDbl(tRs1.Fields("CANT_ABONO")), "###,###,##0.00")
                    tRs1.MoveNext
                Loop
                tLi.SubItems(9) = CDbl(tLi.SubItems(7)) - CDbl(tLi.SubItems(8))
            Else
                tLi.SubItems(8) = "0.00"
                tLi.SubItems(9) = CDbl(tLi.SubItems(7))
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub CuentasAbonos()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim Abonos As Double
    Dim Total1 As Double
    Dim Total2 As Double
    Dim total3 As Double
    Dim ConPag As Integer
    Dim sDias As Double
    Dim sIdCliente As String
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\reportecuentas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Deudas/Abonos)", "F2", 10, hCenter
    oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
    oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
    oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
    oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 110, 40, 30, 40, "Nota", "F2", 10, hCenter
    oDoc.WTextBox 110, 80, 50, 60, "Factura", "F2", 10, hCenter
    oDoc.WTextBox 110, 140, 30, 60, "Fecha Vence", "F2", 10, hCenter
    oDoc.WTextBox 110, 200, 50, 60, "Estatus", "F2", 10, hCenter
    oDoc.WTextBox 110, 260, 40, 70, "Subtotal", "F2", 10, hCenter
    oDoc.WTextBox 110, 330, 40, 70, "IVA", "F2", 10, hCenter
    oDoc.WTextBox 110, 400, 40, 70, "Total", "F2", 10, hLeft
    oDoc.WTextBox 110, 470, 40, 70, "Total Abonos", "F2", 10, hLeft
    'oDoc.WTextBox 110, 500, 40, 70, "Dias Vence", "F2", 10, hLeft
' Cuerpo del reporte
    deuda = 0
    deuda1 = 0
    deuda2 = 0
    totor = 0
    totpr = 0
    Conta = 0
    Total1 = 0
    Total2 = 0
    total3 = 0
    If Not IsNumeric(Text9.Text) Then
        sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, VSCXC3.FECHA_VENCE, VSCXC3.PAGADA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, IIF(VSCXC3.PAGADA = 'S',  SUM(VSCXC3.TOTAL) , ISNULL(SUM(ABONOS_CUENTA.CANT_ABONO), 0)) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA ON VSCXC3.ID_CUENTA = ABONOS_CUENTA.ID_CUENTA WHERE VSCXC3.NOMBRE LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA ORDER BY VSCXC3.NOMBRE, VSCXC3.ID_VENTA"
    Else
        sBuscar = "SELECT VSCXC3.ID_CLIENTE, VSCXC3.NOMBRE, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.FOLIO, VSCXC3.ID_VENTA, VSCXC3.FECHA, VSCXC3.FECHA_VENCE, VSCXC3.PAGADA, SUM(VSCXC3.SUBTOTAL) AS SUBTOTAL, SUM(VSCXC3.IVA) AS IVA, SUM(VSCXC3.TOTAL) AS TOTAL, IIF(VSCXC3.PAGADA = 'S',  SUM(VSCXC3.TOTAL) , ISNULL(SUM(ABONOS_CUENTA.CANT_ABONO), 0)) AS TOT_ABONOS FROM VSCXC3 LEFT OUTER JOIN ABONOS_CUENTA ON VSCXC3.ID_CUENTA = ABONOS_CUENTA.ID_CUENTA WHERE VSCXC3.NOMBRE_COMERCIAL LIKE '%" & Text9.Text & "%' AND VSCXC3.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " ' GROUP BY VSCXC3.NOMBRE, VSCXC3.ID_CLIENTE, VSCXC3.FECHA_VENCE, VSCXC3.FOLIO, VSCXC3.FECHA, VSCXC3.PAGADA, VSCXC3.FECHA_FACTURA, VSCXC3.NOMBRE_COMERCIAL, VSCXC3.ID_VENTA ORDER BY VSCXC3.NOMBRE, VSCXC3.ID_VENTA"
    End If
    Set tRs = cnn.Execute(sBuscar)
    Posi = 140
    sumdeuda = 0
    sumIndi = 0
    Pend = 0
    sumabonos = 0
    sumor = 0
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 110
    oDoc.WLineTo 580, 110
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 580, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If sNombre <> tRs.Fields("NOMBRE") Or sIdCliente <> tRs.Fields("ID_CLIENTE") Then
                If sumor > 0 Then
                    Posi = Posi + 13
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 380, Posi
                    oDoc.WLineTo 440, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 400, 40, 70, Format(sumor, "###,###,##0.00"), "F3", 10, hRight
                    sumor = 0
                    Posi = Posi + 15
                'End If
                'If Abonos > 0 Then
                    Posi = Posi - 22
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 440, Posi
                    oDoc.WLineTo 500, Posi
                    oDoc.LineStroke
                    Posi = Posi + 7
                    oDoc.WTextBox Posi, 470, 40, 70, Format(Abonos, "###,###,##0.00"), "F3", 10, hRight
                    sumabonos = sumabonos + Abonos
                    Posi = Posi + 15
                    Abonos = 0
                End If
                'If deuda > 0 Then
                '    Posi = Posi - 22
                '    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                '    oDoc.MoveTo 360, Posi
                '    oDoc.WLineTo 410, Posi
                '    oDoc.LineStroke
                '    Posi = Posi + 7
                '    oDoc.WTextBox Posi, 350, 40, 60, Format(deuda, "###,###,##0.00"), "F3", 10, hRight
                '    deuda = 0
                '    Posi = Posi + 15
                'End If
                oDoc.WTextBox Posi, 20, 9, 500, tRs.Fields("ID_CLIENTE") & "  -  " & tRs.Fields("NOMBRE"), "F3", 10, hLeft
            End If
            Posi = Posi + 15
            If Not IsNull(tRs.Fields("ID_VENTA")) Then oDoc.WTextBox Posi, 40, 30, 40, tRs.Fields("ID_VENTA"), "F3", 10, hLeft
            If Not IsNull(tRs.Fields("FOLIO")) Then oDoc.WTextBox Posi, 80, 50, 60, Format(tRs.Fields("FOLIO"), "dd/mm/yyyy"), "F3", 10, hCenter
            If Not IsNull(tRs.Fields("FECHA_VENCE")) Then oDoc.WTextBox Posi, 140, 30, 60, Format(tRs.Fields("FECHA_VENCE"), "dd/mm/yyyy"), "F3", 10, hCenter
            If tRs.Fields("PAGADA") = "S" Then
                oDoc.WTextBox Posi, 200, 50, 60, "PAGADA", "F3", 10, hCenter
            Else
                oDoc.WTextBox Posi, 200, 50, 60, "PENDIENTE", "F3", 10, hCenter
            End If
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then oDoc.WTextBox Posi, 260, 40, 70, Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00"), "F3", 10, hRight
            If Not IsNull(tRs.Fields("IVA")) Then oDoc.WTextBox Posi, 330, 40, 70, Format(tRs.Fields("IVA"), "###,###,##0.00"), "F3", 10, hRight
            If Not IsNull(tRs.Fields("TOTAL")) Then oDoc.WTextBox Posi, 400, 40, 70, Format(tRs.Fields("TOTAL"), "###,###,##0.00"), "F3", 10, hRight
            If Not IsNull(tRs.Fields("TOT_ABONOS")) Then oDoc.WTextBox Posi, 470, 40, 70, Format(tRs.Fields("TOT_ABONOS"), "###,###,##0.00"), "F3", 10, hRight
            Total1 = Format(CDbl(Total1) + CDbl(tRs.Fields("TOTAL")), "###,###,##0.00")
            Abonos = Format(CDbl(Abonos) + CDbl(tRs.Fields("TOT_ABONOS")), "###,###,##0.00")
            PosVer = Posi
            sumor = Format(CDbl(sumor) + CDbl(tRs.Fields("TOTAL")), "0.00")
            sumpr = Format(CDbl(sumpr) + CDbl(tRs.Fields("TOTAL")), "0.00")
            Pend = 0
            sNombre = tRs.Fields("NOMBRE")
            sIdCliente = tRs.Fields("ID_CLIENTE")
            tRs.MoveNext
            If Posi >= 700 Then
                oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 140
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Cobrar (Deudas/Abonos)", "F2", 10, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                oDoc.WTextBox 70, 450, 20, 250, DTPicker3.Value, "F2", 8, hLeft
                oDoc.WTextBox 70, 500, 20, 20, "Al", "F2", 8, hLeft
                oDoc.WTextBox 70, 530, 20, 250, DTPicker4.Value, "F2", 8, hLeft
                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 110, 40, 30, 40, "Nota", "F2", 10, hCenter
                oDoc.WTextBox 110, 80, 50, 60, "Factura", "F2", 10, hCenter
                oDoc.WTextBox 110, 140, 30, 60, "Fecha Vence", "F2", 10, hCenter
                oDoc.WTextBox 110, 200, 50, 60, "Estatus", "F2", 10, hCenter
                oDoc.WTextBox 110, 260, 40, 60, "Subtotal", "F2", 10, hCenter
                oDoc.WTextBox 110, 320, 40, 60, "IVA", "F2", 10, hCenter
                oDoc.WTextBox 110, 380, 40, 60, "Total", "F2", 10, hLeft
                oDoc.WTextBox 110, 440, 40, 70, "Total Abonos", "F2", 10, hLeft
                'oDoc.WTextBox 110, 500, 40, 70, "Dias Vence", "F2", 10, hLeft
                ' Cuerpo del reporte
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 110
                oDoc.WLineTo 580, 110
                oDoc.LineStroke
                oDoc.MoveTo 10, 125
                oDoc.WLineTo 580, 125
                oDoc.LineStroke
            End If
        Loop
        'pendientes
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 350, Posi
        oDoc.WLineTo 400, Posi
        oDoc.LineStroke
        Posi = Posi + 7
        oDoc.WTextBox Posi, 350, 40, 60, Format(deuda, "###,###,##0.00"), "F3", 10, hRight
        deuda = 0
        '////
        Posi = Posi + 15
        Posi = Posi + 30
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Compras ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(Total1, "###,###,##0.00"), "F2", 9, hRight 'Format(Text11.Text, "###,###,##0.00"), "F2", 9, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Abonos ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(sumabonos, "###,###,##0.00"), "F2", 9, hRight 'Format(Text12.Text, "###,###,##0.00"), "F2", 9, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 40, 80, "Total De Deudas ", "F2", 9, hLeft
        oDoc.WTextBox Posi, 110, 40, 100, Format(Total1 - sumabonos, "###,###,##0.00"), "F2", 9, hRight ' Format(Text13.Text, "###,###,##0.00"), "F2", 9, hRight
        Cont = Cont + 1
        oDoc.WTextBox 800, 500, 20, 175, ConPag, "F2", 7, hLeft
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
