VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frmrepcomprass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ordenes De Compras  Generadas (Proveedor-Producto)"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   109
      Text            =   "Frmrepcomprass.frx":0000
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   84
      Top             =   3240
      Width           =   975
      Begin VB.Label Label28 
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
         TabIndex        =   85
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "Frmrepcomprass.frx":0007
         MousePointer    =   99  'Custom
         Picture         =   "Frmrepcomprass.frx":0311
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtExcel 
      Height          =   375
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   71
      Text            =   "Frmrepcomprass.frx":08A0
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   9840
      TabIndex        =   63
      Text            =   "Text7"
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9840
      TabIndex        =   60
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Frmrepcomprass.frx":08A6
         MousePointer    =   99  'Custom
         Picture         =   "Frmrepcomprass.frx":0BB0
         Top             =   240
         Width           =   720
      End
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
         TabIndex        =   61
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   14
      Top             =   4440
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "Frmrepcomprass.frx":26F2
         MousePointer    =   99  'Custom
         Picture         =   "Frmrepcomprass.frx":29FC
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   12
      Top             =   5640
      Width           =   975
      Begin VB.Label Label10 
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Frmrepcomprass.frx":45CE
         MousePointer    =   99  'Custom
         Picture         =   "Frmrepcomprass.frx":48D8
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   10
      Top             =   6840
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Frmrepcomprass.frx":641A
         MousePointer    =   99  'Custom
         Picture         =   "Frmrepcomprass.frx":6724
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Ordenes de Compras"
      TabPicture(0)   =   "Frmrepcomprass.frx":8806
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Entradas"
      TabPicture(1)   =   "Frmrepcomprass.frx":8822
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Frame7"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Ordenes Rapidas"
      TabPicture(2)   =   "Frmrepcomprass.frx":883E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "Label24"
      Tab(2).Control(3)=   "ListView3"
      Tab(2).Control(4)=   "ListView4"
      Tab(2).Control(5)=   "ListView5"
      Tab(2).Control(6)=   "Frame8"
      Tab(2).Control(7)=   "Frame9"
      Tab(2).Control(8)=   "Frame12"
      Tab(2).Control(9)=   "Text11"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Abonos a Ordenes"
      TabPicture(3)   =   "Frmrepcomprass.frx":885A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView6"
      Tab(3).Control(1)=   "ListView7"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).Control(3)=   "Command3"
      Tab(3).Control(4)=   "Frame13"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Abonos a Gastos"
      TabPicture(4)   =   "Frmrepcomprass.frx":8876
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView11"
      Tab(4).Control(1)=   "Frame17"
      Tab(4).Control(2)=   "Command5"
      Tab(4).Control(3)=   "Frame18"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Contabilidad"
      TabPicture(5)   =   "Frmrepcomprass.frx":8892
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame20"
      Tab(5).Control(1)=   "Frame19"
      Tab(5).Control(2)=   "ListView14"
      Tab(5).Control(3)=   "ListView15"
      Tab(5).ControlCount=   4
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -71880
         TabIndex        =   108
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame20 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   98
         Top             =   480
         Width           =   6975
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
            Left            =   5640
            Picture         =   "Frmrepcomprass.frx":88AE
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1440
            TabIndex        =   101
            Top             =   480
            Width           =   4095
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Orden"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Fecha"
         Height          =   1215
         Left            =   -67800
         TabIndex        =   93
         Top             =   360
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker11 
            Height          =   375
            Left            =   720
            TabIndex        =   94
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker12 
            Height          =   375
            Left            =   720
            TabIndex        =   95
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label16 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Proveedor"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   90
         Top             =   480
         Width           =   5655
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   5415
         End
         Begin MSComctlLib.ListView ListView12 
            Height          =   975
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1720
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
      Begin VB.Frame Frame13 
         Caption         =   "Proveedor"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   87
         Top             =   480
         Width           =   5655
         Begin MSComctlLib.ListView ListView9 
            Height          =   975
            Left            =   120
            TabIndex        =   89
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1720
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
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   5415
         End
      End
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
         Left            =   -66720
         Picture         =   "Frmrepcomprass.frx":B280
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nombre"
         Height          =   2775
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   4935
         Begin MSComctlLib.ListView ListView8 
            Height          =   1695
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2990
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
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
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
            Left            =   3720
            Picture         =   "Frmrepcomprass.frx":DC52
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Nota:  Se desglosa las ordenes Nacionales,Internacionales"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1575
         Left            =   -71760
         TabIndex        =   80
         Top             =   600
         Width           =   1815
         Begin VB.OptionButton Option9 
            Caption         =   "Producto"
            Height          =   375
            Left            =   360
            TabIndex        =   82
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   360
            TabIndex        =   81
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
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
         Left            =   -66720
         Picture         =   "Frmrepcomprass.frx":10624
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame Frame17 
         Caption         =   "Fecha de pago"
         Height          =   1695
         Left            =   -69120
         TabIndex        =   74
         Top             =   480
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker10 
            Height          =   375
            Left            =   720
            TabIndex        =   78
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   40080
         End
         Begin MSComCtl2.DTPicker DTPicker9 
            Height          =   375
            Left            =   720
            TabIndex        =   75
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   40080
         End
         Begin VB.Label Label23 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   600
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   73
         Top             =   2280
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9551
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
      Begin VB.Frame Frame5 
         Caption         =   "Rango del Reporte"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   64
         Top             =   600
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   720
            TabIndex        =   65
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   720
            TabIndex        =   66
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label6 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "Al :"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   375
         End
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1440
         TabIndex        =   58
         Top             =   7560
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   57
         Top             =   7560
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Frame Frame12 
         Caption         =   "Filtro"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   48
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton Option1 
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
         Begin VB.CheckBox Check6 
            Caption         =   "Sin Filtro"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Fecha de pago"
         Height          =   1695
         Left            =   -69120
         TabIndex        =   39
         Top             =   480
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   375
            Left            =   720
            TabIndex        =   40
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   375
            Left            =   720
            TabIndex        =   41
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label18 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   495
         Left            =   -66480
         TabIndex        =   38
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   37
         Top             =   2280
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9551
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
      Begin VB.Frame Frame9 
         Caption         =   "Nombre"
         Height          =   1095
         Left            =   -73320
         TabIndex        =   30
         Top             =   1080
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
            Picture         =   "Frmrepcomprass.frx":12FF6
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label15 
            Caption         =   "Producto"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Rango de Fecha"
         Height          =   1575
         Left            =   -67920
         TabIndex        =   25
         Top             =   600
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   720
            TabIndex        =   26
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
            Left            =   720
            TabIndex        =   27
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   39576
         End
         Begin VB.Label Label9 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   480
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   255
         Left            =   -66360
         TabIndex        =   24
         Top             =   7800
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   23
         Top             =   5160
         Width           =   9375
         _ExtentX        =   16536
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   22
         Top             =   2520
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3836
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
      Begin VB.Frame Frame6 
         Caption         =   "Nombre"
         Height          =   2535
         Left            =   -69840
         TabIndex        =   18
         Top             =   600
         Width           =   4215
         Begin MSComctlLib.ListView ListView10 
            Height          =   1215
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2143
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
            Left            =   3000
            Picture         =   "Frmrepcomprass.frx":159C8
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Inicio"
         Height          =   195
         Left            =   9240
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   75
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   16
         Top             =   3240
         Width           =   9015
         _ExtentX        =   15901
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
      Begin VB.Frame Frame1 
         Caption         =   "Rango del Reporte"
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4455
         Begin VB.OptionButton Option13 
            Caption         =   "del Pago"
            Height          =   255
            Left            =   2160
            TabIndex        =   106
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option12 
            Caption         =   "de la Orden"
            Height          =   195
            Left            =   600
            TabIndex        =   105
            Top             =   720
            Value           =   -1  'True
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   72
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50003969
            CurrentDate     =   40077
         End
         Begin VB.Frame Frame15 
            Height          =   1815
            Left            =   1320
            TabIndex        =   52
            Top             =   960
            Width           =   3135
            Begin VB.OptionButton Option14 
               Caption         =   "Todos las activas"
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   1560
               Width           =   2535
            End
            Begin VB.OptionButton Option16 
               Caption         =   "Todos"
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   1320
               Width           =   2535
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Con Pago/Sin Entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   1080
               Width           =   2415
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Con Pago/Con Entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   840
               Width           =   2775
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Sin Pago/Llegada Parcial"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   600
               Width           =   2895
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Sin Entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Sin Pago/Con Entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   2655
            End
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2400
            TabIndex        =   6
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
            Left            =   2040
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4455
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7858
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
      Begin MSComctlLib.ListView ListView14 
         Height          =   495
         Left            =   -66120
         TabIndex        =   103
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView15 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   104
         Top             =   1680
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9763
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
      Begin VB.Label Label24 
         Caption         =   "Numero de Orden"
         Height          =   255
         Left            =   -73320
         TabIndex        =   107
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Nota: Las  entradas aparece de acuerdo a la orden que se le dio entrada."
         Height          =   495
         Left            =   -74640
         TabIndex        =   70
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Label14 
         Caption         =   "Desglose de Productos"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Lista de Proveedores"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   2280
         Width           =   2655
      End
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   9840
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Frmrepcomprass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim StrRep5 As String
Dim StrRep6 As String
Dim org As Integer
Dim sIdProv As String
Dim IdProveedor As String
Dim IdProveedor1 As String
Dim IdProveedor2 As String
Private Sub Check1_Click()
    cmdBuscar.Enabled = True
    Check4.Value = 0
    Check6.Value = 0
End Sub
Private Sub Check4_Click()
    cmdBuscar.Enabled = True
    Check1.Value = 0
    Check6.Value = 0
End Sub
Private Sub Check3_Click()
    cmdBuscar.Enabled = True
    Check2.Value = 0
    Check1.Value = 0
    Check4.Value = 0
End Sub
Private Sub Check6_Click()
    Check4.Value = 0
    Check1.Value = 0
End Sub
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs1 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim sCadena As String
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView1.ListItems.Clear
    If Check1.Value = 0 And Check4.Value = 0 Then
        If Check6.Value = 1 Then
            ListView1.ColumnHeaders.Clear
            With ListView1
                .View = lvwReport
                .GridLines = True
                .LabelEdit = lvwManual
                .HideSelection = False
                .HotTracking = False
                .HoverSelection = False
                .ColumnHeaders.Add , , "O.C", 1000
                .ColumnHeaders.Add , , "Nombre", 2500
                .ColumnHeaders.Add , , "Fecha", 1200
                .ColumnHeaders.Add , , "Total", 1000
                .ColumnHeaders.Add , , "Idprove", 0
                .ColumnHeaders.Add , , "Tipo", 800
                .ColumnHeaders.Add , , "Estado", 3000
                .ColumnHeaders.Add , , "Producto", 1000
                .ColumnHeaders.Add , , "Cantidad", 1000
                .ColumnHeaders.Add , , "Precio", 1000
                .ColumnHeaders.Add , , "Surtido", 1000
                .ColumnHeaders.Add , , "Entrada", 1000
                .ColumnHeaders.Add , , "Dio Etrada", 1000
                .ColumnHeaders.Add , , "IVA", 1000
                .ColumnHeaders.Add , , "Flete", 1000
                .ColumnHeaders.Add , , "Otros", 1000
                .ColumnHeaders.Add , , "Cheque", 1000
                .ColumnHeaders.Add , , "Fecha Cheque", 1000
                .ColumnHeaders.Add , , "Factura", 1000
            End With
'cv se va sin proveedor y marca error
            sBuscar = "SELECT * FROM VsOrdenes WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
            If IdProveedor <> "" Then
                sBuscar = sBuscar & " AND ID_PROVEEDOR = " & IdProveedor
            End If
            sBuscar = sBuscar & " ORDER BY NUM_ORDEN, FECHA"
            Set tRs = cnn.Execute(sBuscar)
            StrRep = sBuscar
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
                    'If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = Format((tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS")) - tRs.Fields("DISCOUNT"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then tLi.SubItems(4) = tRs.Fields("ID_PROVEEDOR")
                    If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(5) = tRs.Fields("TIPO")
                    orde = tRs.Fields("NUM_ORDEN")
                    If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(7) = tRs.Fields("ID_PRODUCTO")
                    If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(8) = tRs.Fields("CANTIDAD")
                    If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(9) = Format(tRs.Fields("PRECIO"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(10) = tRs.Fields("SURTIDO")
                    If Not IsNull(tRs.Fields("ID_ENTRADA")) Then tLi.SubItems(11) = tRs.Fields("ID_ENTRADA")
                    If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(12) = tRs.Fields("USUARIO")
                    If tRs.Fields("CONFIRMADA") = "N" Then
                        tLi.SubItems(6) = "PRE-ORDEN"
                    End If
                    If tRs.Fields("CONFIRMADA") = "P" Then
                        tLi.SubItems(6) = "PENDIENTE DE AUTORIZAR"
                    End If
                    If tRs.Fields("CONFIRMADA") = "S" Then
                        tLi.SubItems(6) = "PENDIENTE DE IMPRIMIR"
                    End If
                    sBuscar = "SELECT * FROM vsordpende WHERE NUM_ORDEN= '" & tRs.Fields("NUM_ORDEN") & "' AND  TIPO= '" & tRs.Fields("TIPO") & "' AND  ID_PRODUCTO= '" & tRs.Fields("ID_PRODUCTO") & "'"
                    Set tRs3 = cnn.Execute(sBuscar)
                    Dim catpe  As Double
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        catpe = CDbl(tRs3.Fields("CANTIDAD")) - CDbl(tRs3.Fields("SURTIDO"))
                        If tRs.Fields("CONFIRMADA") = "X" And tRs3.Fields("SUR") = 0 Then
                            tLi.SubItems(6) = "SIN PAGO/SIN ENTRADA"
                        End If
                        If tRs.Fields("CONFIRMADA") = "X" And catpe < tRs3.Fields("CAN") And tRs3.Fields("SUR") <> 0 Then
                            tLi.SubItems(6) = "SIN PAGO/ENTRADA PARCIAL"
                        End If
                        If tRs.Fields("CONFIRMADA") = "Y" And tRs3.Fields("SURTIDO") = 0 Then
                            tLi.SubItems(6) = "CON PAGO/SIN ENTRADA"
                        End If
                        If tRs.Fields("CONFIRMADA") = "Y" And catpe < tRs3.Fields("CANTIDAD") And tRs3.Fields("SURTIDO") <> 0 Then
                            tLi.SubItems(6) = "CON PAGO/ENTRADA PARCIAL"
                        End If
                    Else
                        tLi.SubItems(6) = "NO SE ENCONTRO DETALLE"
                    End If
                    If tRs.Fields("CONFIRMADA") = "X" And catpe = 0 Then
                        tLi.SubItems(6) = "SIN PAGO/CON ENTRADA"
                    End If
                    
                    If tRs.Fields("CONFIRMADA") = "Y" And catpe = 0 Then
                        tLi.SubItems(6) = "CON PAGO/CON ENTRADA"
                    End If
                    If tRs.Fields("CONFIRMADA") = "E" Then
                        tLi.SubItems(6) = "CANCELADA"
                    End If
                    If Not IsNull(tRs.Fields("TAX")) Then tLi.SubItems(13) = Format(tRs.Fields("TAX"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("FREIGHT")) Then tLi.SubItems(14) = Format(tRs.Fields("FREIGHT"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("OTROS_CARGOS")) Then tLi.SubItems(15) = Format(tRs.Fields("OTROS_CARGOS"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("FACT_PROVE")) Then tLi.SubItems(16) = tRs.Fields("FACT_PROVE")
                    catpe = 0
                    tRs.MoveNext
                Loop
            End If
        End If
    Else
        If Option12.Value Then
            If Check1.Value = 1 Then
                If Option4.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURTIDO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, PROVEEDOR.ID_PROVEEDOR, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.CONFIRMADA = 'X' AND ORDEN_COMPRA_DETALLE.SURTIDO = 0 AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, PROVEEDOR.ID_PROVEEDOR, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option3.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURTIDO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, PROVEEDOR.ID_PROVEEDOR, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA_DETALLE.SURTIDO = ORDEN_COMPRA_DETALLE.CANTIDAD) AND PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.CONFIRMADA = 'X' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, PROVEEDOR.ID_PROVEEDOR, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option5.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURITIDO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA_DETALLE.SURTIDO < ORDEN_COMPRA_DETALLE.CANTIDAD) AND (ORDEN_COMPRA_DETALLE.SURTIDO <> 0) AND PROVEEDORNOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.CONFIRMADA = 'X' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option6.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURITIDO, ORDEN_COMPRA.TOTAL, dbo.ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA_DETALLE.SURTIDO < ORDEN_COMPRA_DETALLE.CANTIDAD) AND (ORDEN_COMPRA_DETALLE.SURTIDO <> 0) AND PROVEEDORNOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.CONFIRMADA = 'Y' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option7.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURITIDO, ORDEN_COMPRA.TOTAL, dbo.ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.CONFIRMADA = 'Y' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option16.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURITIDO, ORDEN_COMPRA.TOTAL, dbo.ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                If Option14.Value = True Then
                    sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND ORDEN_COMPRA.CONFIRMADA NOT IN ('E', 'D')" & _
                    "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
                End If
                StrRep = sBuscar
            End If
            If Check4.Value = 1 Then
                sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, SUM(ORDEN_COMPRA_DETALLE.SURTIDO) AS SURITIDO, ORDEN_COMPRA.TOTAL, dbo.ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE ORDEN_COMPRA_DETALLE.ID_PRODUCTO LIKE '%" & Text1.Text & "%'AND ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' " & _
                "GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.FACT_PROVE"
            End If
            Text10.Text = sBuscar
        Else
            If Check1.Value = 1 Then
            'agregas subconsulta que busque
                If Option4.Value = True Then
                    sBuscar = "SELECT * FROM vsordencom WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'X' AND SURTIDO = 0 AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY NUM_ORDEN, FECHA"
                End If
                If Option3.Value = True Then
                    sBuscar = "SELECT * FROM vsordencom1 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'X' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY NUM_ORDEN, FECHA"
                End If
                If Option5.Value = True Then
                    sBuscar = "SELECT * FROM vsordencom2 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'X' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY FNUM_ORDEN, ECHA"
                End If
                If Option6.Value = True Then
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'NACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = "SELECT * FROM vsordencom2 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'N'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INTERNACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom2 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'I'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INDIRECTA'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom2 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'X'"
                    sCadena = ""
                End If
                If Option7.Value = True Then
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'NACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = "SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'N'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INTERNACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'I'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INDIRECTA'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'X'"
                    sCadena = ""
                End If
                If Option16.Value = True Or Option14.Value = True Then
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'NACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = "SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'N'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INTERNACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'I'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INDIRECTA'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom3 WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'X'"
                    sCadena = ""
                End If
                StrRep = sBuscar
            End If
            If Check4.Value = 1 Then
                                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'NACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = "SELECT * FROM vsordencom WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'N'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INTERNACIONAL'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'I'"
                    sCadena = ""
                    sBuscar = "SELECT NUM_ORDEN FROM CHEQUES WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND TIPO_ORDEN = 'INDIRECTA'"
                    Set tRs = cnn.Execute(sBuscar)
                    Do While Not tRs.EOF
                        sCadena = sCadena + tRs.Fields(num_orden)
                        tRs.MoveNext
                    Loop
                    sBuscar = sBuscar + "UNION SELECT * FROM vsordencom WHERE  NOMBRE LIKE '%" & Text1.Text & "%' AND CONFIRMADA = 'Y' AND NUM_ORDEN IN (" & Mid(sCadena, 1, Len(sCadena) - 2) & ") AND TIPO = 'X'"
                    sCadena = ""
            End If
        End If
        Set tRs = cnn.Execute(sBuscar)
        StrRep2 = sBuscar
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
                'If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = Format((tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("DISCOUNT")), "###,###,###,##0.00")
                If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then tLi.SubItems(4) = tRs.Fields("ID_PROVEEDOR")
                If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(5) = tRs.Fields("TIPO")
                If Not IsNull(tRs.Fields("NUM_ORDEN")) Then orde = tRs.Fields("NUM_ORDEN")
                If Check1 = 0 And Check4 = 0 Then
                    If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(7) = tRs.Fields("ID_PRODUCTO")
                    If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(8) = tRs.Fields("CANTIDAD")
                    If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(9) = Format(tRs.Fields("PRECIO"), "###,###,###,##0.00")
                    If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(10) = tRs.Fields("SURTIDO")
                End If
                If tRs.Fields("CONFIRMADA") = "N" Then
                    tLi.SubItems(6) = "PRE-ORDEN"
                End If
                If tRs.Fields("CONFIRMADA") = "P" Then
                    tLi.SubItems(6) = "PENDIENTE DE AUTORIZAR"
                End If
                If tRs.Fields("CONFIRMADA") = "S" Then
                    tLi.SubItems(6) = "PENDIENTE DE IMPRIMIR"
                End If
                sBuscar = "SELECT SUM(ISNULL(CANTIDAD, 0)) AS CAN, SUM(ISNULL(SURTIDO, 0)) AS SUR FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA= " & tRs.Fields("ID_ORDEN_COMPRA")
                Set tRs3 = cnn.Execute(sBuscar)
                If Not (tRs3.EOF And tRs3.BOF) Then
                    If tRs3.Fields("CAN") <> "" Then
                        catpe = CDbl(tRs3.Fields("CAN")) - CDbl(tRs3.Fields("SUR"))
                    End If
                End If
                If tRs.Fields("CONFIRMADA") = "X" And tRs3.Fields("SUR") = 0 Then
                    tLi.SubItems(6) = "SIN PAGO/SIN ENTRADA"
                End If
                If tRs.Fields("CONFIRMADA") = "X" And catpe = 0 Then
                    tLi.SubItems(6) = "SIN PAGO/CON ENTRADA"
                End If
                If tRs.Fields("CONFIRMADA") = "X" And catpe < tRs3.Fields("CAN") And tRs3.Fields("SUR") <> 0 Then
                    tLi.SubItems(6) = "SIN PAGO/ENTRADA PARCIAL"
                End If
                If tRs.Fields("CONFIRMADA") = "Y" And catpe < tRs3.Fields("CAN") And tRs3.Fields("SUR") <> 0 Then
                    tLi.SubItems(6) = "CON PAGO/ENTRADA PARCIAL"
                End If
                If tRs.Fields("CONFIRMADA") = "Y" And catpe = 0 Then
                    tLi.SubItems(6) = "CON PAGO/CON ENTRADA"
                End If
                If tRs.Fields("CONFIRMADA") = "Y" And tRs3.Fields("SUR") = 0 Then
                    tLi.SubItems(6) = "CON PAGO/SIN ENTRADA"
                End If
                If tRs.Fields("CONFIRMADA") = "E" Then
                    tLi.SubItems(6) = "CANCELADA"
                End If
                If Not IsNull(tRs.Fields("TAX")) Then tLi.SubItems(11) = Format(tRs.Fields("TAX"), "###,###,###,##0.00")
                If Not IsNull(tRs.Fields("FREIGHT")) Then tLi.SubItems(12) = Format(tRs.Fields("FREIGHT"), "###,###,###,##0.00")
                If Not IsNull(tRs.Fields("OTROS_CARGOS")) Then tLi.SubItems(13) = Format(tRs.Fields("OTROS_CARGOS"), "###,###,###,##0.00")
                If Not IsNull(tRs.Fields("FACT_PROVE")) Then tLi.SubItems(14) = tRs.Fields("FACT_PROVE")
                catpe = 0
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
    Frame11.Visible = True
    ListView2.ListItems.Clear
    If Option8.Value = False And Option9.Value = False Then
        MsgBox "SELECCIONE UNA FORMA DE BUSQUEDA ", vbInformation, "SACC"
    Else
        If Option8.Value = True Then
            sBuscar = "SELECT * FROM vsentrconta WHERE  NOMBRE LIKE '%" & Text2.Text & "%'AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' ORDER BY FECHA"
        End If
        If Option9.Value = True Then
            sBuscar = "SELECT * FROM vsentrconta WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'AND FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & "' ORDER BY FECHA"
        End If
          StrRep6 = sBuscar
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("FECHA")
                tLi.SubItems(3) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
                tLi.SubItems(4) = tRs.Fields("NUM_ORDEN")
                If Not IsNull(tRs.Fields("ID_ENTRADA")) Then tLi.SubItems(5) = tRs.Fields("ID_ENTRADA")
                tLi.SubItems(6) = tRs.Fields("ID_PRODUCTO")
                tLi.SubItems(7) = tRs.Fields("CANTIDAD")
                tLi.SubItems(8) = Format(tRs.Fields("PRECIO"), "###,###,###,##0.00")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView3.ListItems.Clear
    If Option1.Value = True Then
        sBuscar = "SELECT ID_ORDEN_RAPIDA, NOMBRE, ID_PROVEEDOR, FECHA, IVARETENIDO, ISR2, IVADIEZ, RETENCION, TOTAL, ESTADO FROM VsOrdenRapida WHERE  NOMBRE LIKE '%" & Text3.Text & "%'AND FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "' GROUP  BY NOMBRE, ID_PROVEEDOR, ID_ORDEN_RAPIDA, FECHA, IVARETENIDO, ISR2, IVADIEZ, RETENCION, TOTAL, ESTADO ORDER BY ID_ORDEN_RAPIDA"
        If Text11.Text <> "" Then
            sBuscar = "SELECT ID_ORDEN_RAPIDA, NOMBRE, ID_PROVEEDOR, FECHA, IVARETENIDO, ISR2, IVADIEZ, RETENCION, TOTAL, ESTADO FROM VsOrdenRapida WHERE (ID_ORDEN_RAPIDA = " & Text11.Text & ") GROUP BY NOMBRE, ID_PROVEEDOR, ID_ORDEN_RAPIDA, FECHA, IVARETENIDO, ISR2, IVADIEZ, RETENCION, TOTAL, ESTADO"
        End If
        Set tRs = cnn.Execute(sBuscar)
        StrRep6 = sBuscar
        org = 1
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("IVARETENIDO")) Then
                    tLi.SubItems(3) = tRs.Fields("IVARETENIDO")
                Else
                    tLi.SubItems(3) = "0"
                End If
                If Not IsNull(tRs.Fields("ISR2")) Then
                    tLi.SubItems(4) = tRs.Fields("ISR2")
                Else
                    tLi.SubItems(4) = "0"
                End If
                If Not IsNull(tRs.Fields("IVADIEZ")) Then
                    tLi.SubItems(5) = tRs.Fields("IVADIEZ")
                Else
                    tLi.SubItems(5) = "0"
                End If
                If Not IsNull(tRs.Fields("RETENCION")) Then
                    tLi.SubItems(6) = tRs.Fields("RETENCION")
                Else
                    tLi.SubItems(6) = "0"
                End If
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    tLi.SubItems(7) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
                Else
                    tLi.SubItems(7) = "0"
                End If
                If Not IsNull(tRs.Fields("ESTADO")) Then
                    If tRs.Fields("ESTADO") = "F" Then
                        tLi.SubItems(8) = "PAGADA"
                    Else
                        If tRs.Fields("ESTADO") = "A" Then
                            tLi.SubItems(8) = "PENDIENTE DE PAGO"
                        Else
                            If tRs.Fields("ESTADO") = "M" Then
                                tLi.SubItems(8) = "EN MODIFICACION"
                            Else
                                tLi.SubItems(8) = "ORDEN PERDIDA"
                            End If
                        End If
                    End If
                Else
                    tLi.SubItems(8) = "ORDEN PERDIDA"
                End If
                tRs.MoveNext
            Loop
       End If
    Else
        sBuscar = "SELECT * FROM vsorderapidaa WHERE  ID_PRODUCTO  like'%" & Text3.Text & "%'AND FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "' ORDER  BY FECHA DESC"
        Set tRs = cnn.Execute(sBuscar)
        StrRep6 = sBuscar
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                 tLi.SubItems(2) = tRs.Fields("FECHA")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView6.ListItems.Clear
    If IdProveedor1 = "" Then
        sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, PROVEEDOR.NOMBRE, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, ABONOS_PAGO_OC.BANCO, dbo.ABONOS_PAGO_OC.NUM_ORDEN, dbo.ABONOS_PAGO_OC.TIPO FROM ORDEN_COMPRA INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.NUM_ORDEN = ABONOS_PAGO_OC.NUM_ORDEN AND ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker7.Value & "' AND '" & DTPicker8.Value & "')"
    Else
        sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, PROVEEDOR.NOMBRE, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.TAX, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.OTROS_CARGOS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, ABONOS_PAGO_OC.BANCO, dbo.ABONOS_PAGO_OC.NUM_ORDEN, dbo.ABONOS_PAGO_OC.TIPO FROM ORDEN_COMPRA INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.NUM_ORDEN = ABONOS_PAGO_OC.NUM_ORDEN AND ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.ID_PROVEEDOR = " & IdProveedor1 & " AND (ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker7.Value & "' AND '" & DTPicker8.Value & "')"
    End If
    StrRep5 = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView6.ListItems.Add(, , tRs.Fields("FECHA"))
            If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(2) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("TAX")) Then tLi.SubItems(5) = Format(tRs.Fields("TAX"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("FREIGHT")) Then tLi.SubItems(6) = Format(tRs.Fields("FREIGHT"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("OTROS_CARGOS")) Then tLi.SubItems(7) = Format(tRs.Fields("OTROS_CARGOS"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("TAX")) Then tLi.SubItems(8) = Format((tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS")) - tRs.Fields("DISCOUNT"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("NUMCHEQUE")) Then tLi.SubItems(9) = tRs.Fields("NUMCHEQUE")
            If Not IsNull(tRs.Fields("NUMTRANS")) Then tLi.SubItems(10) = tRs.Fields("NUMTRANS")
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then tLi.SubItems(11) = Format(tRs.Fields("CANT_ABONO"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("BANCO")) Then tLi.SubItems(12) = tRs.Fields("BANCO")
            tRs.MoveNext
        Loop
    End If
    'suma
End Sub
Private Sub SUMA()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView7.ListItems.Clear
    sBuscar = "SELECT SUM(TOTAL) AS TOTAL,NOMBRE FROM CHEQUES WHERE  NOMBRE like  '%" & Text4.Text & "%'AND FECHA BETWEEN '" & DTPicker7.Value & "' AND '" & DTPicker8.Value & "' GROUP BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView7.ListItems.Add(, , tRs.Fields(""))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command5_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sTipo As String
    If IdProveedor2 = "" Then
        sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.NO_CHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ID_ORDEN_RAPIDA, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.IVA) AS IVA, SUM(ORDEN_RAPIDA_DETALLE.IVARETENIDO) AS IVARETENIDO,  SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, SUM(ORDEN_RAPIDA_DETALLE.ISR) AS ISR FROM ORDEN_RAPIDA_DETALLE INNER JOIN ORDEN_RAPIDA ON ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ABONOS_PAGO_OC.NUM_ORDEN WHERE (ABONOS_PAGO_OC.TIPO = 'R') AND (ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker9.Value & "' AND '" & DTPicker10.Value & "') GROUP BY ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.NO_CHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ID_ORDEN_RAPIDA"
    Else
        sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.NO_CHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ID_ORDEN_RAPIDA, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.IVA) AS IVA, SUM(ORDEN_RAPIDA_DETALLE.IVARETENIDO) AS IVARETENIDO,  SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, SUM(ORDEN_RAPIDA_DETALLE.ISR) AS ISR FROM ORDEN_RAPIDA_DETALLE INNER JOIN ORDEN_RAPIDA ON ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ABONOS_PAGO_OC.NUM_ORDEN WHERE (ABONOS_PAGO_OC.TIPO = 'R') AND PROVEEDOR_CONSUMO.ID_PROVEEDOR LIKE '%" & IdProveedor2 & "%' " & _
        "(ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker9.Value & "' AND '" & DTPicker10.Value & "') GROUP BY ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.NO_CHEQUE, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.CANT_ABONO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ID_ORDEN_RAPIDA"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView11.ListItems.Add(, , tRs.Fields("FECHA"))
            If Not IsNull(tRs.Fields("ID_ORDEN_RAPIDA")) Then tLi.SubItems(1) = tRs.Fields("ID_ORDEN_RAPIDA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("NO_CHEQUE")) Then tLi.SubItems(3) = tRs.Fields("NO_CHEQUE")
            If Not IsNull(tRs.Fields("NUMTRANS")) Then tLi.SubItems(4) = tRs.Fields("NUMTRANS")
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then tLi.SubItems(5) = Format(tRs.Fields("CANT_ABONO"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(6) = Format(tRs.Fields("SUBTOTAL"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(7) = Format(tRs.Fields("IVA"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("IVARETENIDO")) Then tLi.SubItems(8) = Format(tRs.Fields("IVARETENIDO"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("ISR")) Then tLi.SubItems(9) = Format(tRs.Fields("ISR"), "###,###,###,##0.00")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(10) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command6_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView15.ListItems.Clear
    If Option10.Value = True Then
        sBuscar = "SELECT * FROM CHEQUES WHERE NOMBRE LIKE '%" & Text9.Text & "%'AND FECHA BETWEEN '" & DTPicker11.Value & "' AND '" & DTPicker12.Value & "' ORDER BY FECHA ASC"
    End If
    If Option11.Value = True Then
        sBuscar = "SELECT * FROM CHEQUES WHERE NUM_ORDEN LIKE '%" & Text9.Text & "%'AND FECHA BETWEEN '" & DTPicker11.Value & "' AND '" & DTPicker12.Value & "' ORDER BY FECHA ASC"
    End If
    StrRep5 = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView15.ListItems.Add(, , tRs.Fields("ID_CHEQUE"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("FECHA")
            tLi.SubItems(3) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
            tLi.SubItems(4) = tRs.Fields("BANCO")
            tLi.SubItems(5) = tRs.Fields("NUM_CHEQUE")
            tLi.SubItems(6) = tRs.Fields("TIPO_ORDEN")
            tLi.SubItems(7) = tRs.Fields("NUM_ORDEN")
            tLi.SubItems(8) = tRs.Fields("FECHA_REALIZADO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
    Dim sBuscar As String
    IdProveedor = ""
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    DTPicker3.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker4.Value = Format(Date, "dd/mm/yyyy")
    DTPicker5.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker6.Value = Format(Date, "dd/mm/yyyy")
    DTPicker7.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker8.Value = Format(Date, "dd/mm/yyyy")
    DTPicker9.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker10.Value = Format(Date, "dd/mm/yyyy")
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
        '.ColumnHeaders.Add , , "O.C", 1000
        '.ColumnHeaders.Add , , "Nombre", 2500
        '.ColumnHeaders.Add , , "Fecha", 1200
        '.ColumnHeaders.Add , , "Total", 0
        '.ColumnHeaders.Add , , "Idprove", 0
        '.ColumnHeaders.Add , , "Tipo", 800
        '.ColumnHeaders.Add , , "Estado", 3000
        '.ColumnHeaders.Add , , "Producto", 0
        '.ColumnHeaders.Add , , "Cantidad", 0
        '.ColumnHeaders.Add , , "Precio", 0
        '.ColumnHeaders.Add , , "Surtido", 0
        '.ColumnHeaders.Add , , "IVA", 0
        '.ColumnHeaders.Add , , "Flete", 0
        '.ColumnHeaders.Add , , "Otros", 0
        '.ColumnHeaders.Add , , "Cheque", 0
        '.ColumnHeaders.Add , , "Fecha", 0
        '.ColumnHeaders.Add , , "Factura", 0
        .ColumnHeaders.Add , , "O.C", 1000
        .ColumnHeaders.Add , , "Nombre", 2500
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 0
        .ColumnHeaders.Add , , "Idprove", 0
        .ColumnHeaders.Add , , "Tipo", 800
        .ColumnHeaders.Add , , "Producto", 0
        .ColumnHeaders.Add , , "Cantidad", 0
        .ColumnHeaders.Add , , "Precio", 0
        .ColumnHeaders.Add , , "Surtido", 0
        .ColumnHeaders.Add , , "Status", 0
        .ColumnHeaders.Add , , "IVA", 0
        .ColumnHeaders.Add , , "Flete", 0
        .ColumnHeaders.Add , , "Otros", 0
        .ColumnHeaders.Add , , "Factura", 0
    End With
    With ListView8
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 1000
        .ColumnHeaders.Add , , "Nombre", 3000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 0
        .ColumnHeaders.Add , , "Nombre", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1200
        .ColumnHeaders.Add , , "Num_orden", 1000
        .ColumnHeaders.Add , , "Entrada", 1000
        .ColumnHeaders.Add , , "Producto", 1000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Orden", 600
        .ColumnHeaders.Add , , "Proveedor.", 6000
        .ColumnHeaders.Add , , "Fecha.", 1000
        .ColumnHeaders.Add , , "Iva Retenido", 1000
        .ColumnHeaders.Add , , "ISR", 1000
        .ColumnHeaders.Add , , "Iva 10%", 1000
        .ColumnHeaders.Add , , "Retencion", 1000
        .ColumnHeaders.Add , , "Total", 1000
        .ColumnHeaders.Add , , "Estado", 1000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "id.", 0
        .ColumnHeaders.Add , , "Producto", 1200
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio", 1200
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "id.", 0
        .ColumnHeaders.Add , , "Producto", 1200
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio", 1200
    End With
    With ListView6
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Fecha.", 0
        .ColumnHeaders.Add , , "Num. Orden", 1200
        .ColumnHeaders.Add , , "Tipo", 1200
        .ColumnHeaders.Add , , "Proveedor", 1200
        .ColumnHeaders.Add , , "Subtotal", 2000
        .ColumnHeaders.Add , , "Impuesto", 2000
        .ColumnHeaders.Add , , "Flete", 2000
        .ColumnHeaders.Add , , "Otros Cargos", 2000
        .ColumnHeaders.Add , , "Total Orden", 2000
        .ColumnHeaders.Add , , "Numero de Cheque", 1200
        .ColumnHeaders.Add , , "Numero de Transaccion", 1200
        .ColumnHeaders.Add , , "Importe Abono", 1200
        .ColumnHeaders.Add , , "Banco", 1200
         
    End With
    With ListView7
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , ".", 2000
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , ".", 0
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
    End With
    With ListView9
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 0
        .ColumnHeaders.Add , , "Nombre", 5000
    End With
    With ListView10
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 1000
        .ColumnHeaders.Add , , "Nombre", 3000
    End With
    With ListView11
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "FECHA", 1000
        .ColumnHeaders.Add , , "NUM. ORDEN", 4500
        .ColumnHeaders.Add , , "PROVEEDOR", 2000
        .ColumnHeaders.Add , , "NUM. CHEQUE", 2000
        .ColumnHeaders.Add , , "NUM. TRANSACCION", 2000
        .ColumnHeaders.Add , , "IMPORTE ABONO", 3000
        .ColumnHeaders.Add , , "SUBTOTAL", 3000
        .ColumnHeaders.Add , , "IVA", 3000
        .ColumnHeaders.Add , , "IVA RETENIDO", 3000
        .ColumnHeaders.Add , , "ISR", 3000
        .ColumnHeaders.Add , , "TOTAL", 3000
    End With
    With ListView12
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 0
        .ColumnHeaders.Add , , "Nombre", 5000
    End With
    With ListView14
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , ".", 2000
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , ".", 0
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
        .ColumnHeaders.Add , , "", 1200
    End With
    With ListView15
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id_Cheque.", 0
        .ColumnHeaders.Add , , "Nombre", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1200
        .ColumnHeaders.Add , , "Banco", 2000
        .ColumnHeaders.Add , , "Numero de Cheque", 1200
        .ColumnHeaders.Add , , "Tipo de Orden", 1200
        .ColumnHeaders.Add , , "Numero de Orden", 1200
        .ColumnHeaders.Add , , "Fecha de Pago", 1200
    End With
End Sub
Private Sub Image1_Click()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim ConPag As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim NoOC As String
    Dim PROV As String
    ConPag = 1
    Cont = 1
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
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
    oDoc.WImage 70, 40, 43, 161, "Logo"
    sBuscar = "SELECT * FROM EMPRESA  "
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 40, 205, 100, 175, tRs.Fields("NOMBRE"), "F3", 8, hCenter
    oDoc.WTextBox 60, 224, 100, 175, tRs.Fields("DIRECCION"), "F3", 8, hLeft
    oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs.Fields("COLONIA"), "F3", 8, hLeft
    oDoc.WTextBox 70, 205, 100, 175, tRs.Fields("ESTADO") & "," & tRs.Fields("CD"), "F3", 8, hCenter
    oDoc.WTextBox 80, 205, 100, 175, tRs.Fields("TELEFONO"), "F3", 8, hCenter
    oDoc.WTextBox 80, 340, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
    'CAJA1
    If Option4.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE LLEGAR", "F3", 10, hCenter
    End If
    If Option3.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE PAGO CON ENTRADA", "F3", 10, hCenter
    End If
    If Option5.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE PAGO CON ENTRADA PARCIAL", "F3", 10, hCenter
    End If
    If Option6.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PAGADAS", "F3", 10, hCenter
    End If
    If Option7.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PAGADAS SIN ENTRADA", "F3", 10, hCenter
    End If
    If Option16.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA", "F3", 10, hCenter
    End If
    If Option14.Value Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA ACTIVAS", "F3", 10, hCenter
    End If
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 135
    oDoc.WLineTo 580, 135
    oDoc.LineStroke
    Posi = 135
    ' DETALLE
    Do While Cont <= ListView1.ListItems.Count
        ' ENCABEZADO DEL DETALLE
        If NoOC <> ListView1.ListItems(Cont) Then
            Posi = Posi + 20
            oDoc.WTextBox Posi, 5, 20, 40, "OC", "F2", 8, hCenter, , vbBlue
            oDoc.WTextBox Posi, 50, 20, 200, "PROVEEDOR", "F2", 8, hCenter, , vbBlue
            oDoc.WTextBox Posi, 244, 20, 70, "FECHA", "F2", 8, hCenter, , vbBlue
            oDoc.WTextBox Posi, 318, 20, 50, "TOTAL", "F2", 8, hCenter, , vbBlue
            oDoc.WTextBox Posi, 372, 20, 70, "TIPO", "F2", 8, hCenter, , vbBlue
            oDoc.WTextBox Posi, 446, 20, 120, "ESTADO", "F2", 8, hCenter, , vbBlue
            Posi = Posi + 12
            oDoc.WTextBox Posi, 5, 20, 40, ListView1.ListItems(Cont), "F3", 8, hLeft
            oDoc.WTextBox Posi, 50, 20, 190, ListView1.ListItems(Cont).SubItems(1), "F3", 8, hLeft
            oDoc.WTextBox Posi, 254, 20, 70, Format(ListView1.ListItems(Cont).SubItems(2), "dd/mm/yyyy"), "F3", 8, hLeft
            oDoc.WTextBox Posi, 318, 20, 50, Format(ListView1.ListItems(Cont).SubItems(3), "###,###,##0.00"), "F3", 8, hLeft
            oDoc.WTextBox Posi, 372, 20, 70, ListView1.ListItems(Cont).SubItems(5), "F3", 8, hLeft
            oDoc.WTextBox Posi, 446, 20, 120, ListView1.ListItems(Cont).SubItems(6), "F3", 8, hLeft
            ' Linea
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, Posi
            oDoc.WLineTo 580, Posi
            oDoc.LineStroke
            Posi = Posi + 8
            NoOC = ListView1.ListItems(Cont)
            PROV = ListView1.ListItems(Cont).SubItems(1)
            oDoc.WTextBox Posi, 5, 20, 100, "PRODUCTO", "F2", 8, hCenter
            oDoc.WTextBox Posi, 110, 20, 50, "CANTIDAD", "F2", 8, hCenter
            oDoc.WTextBox Posi, 164, 20, 70, "PRECIO", "F2", 8, hCenter
            oDoc.WTextBox Posi, 228, 20, 50, "SURTIDO", "F2", 8, hCenter
            Posi = Posi + 8
        End If
        oDoc.WTextBox Posi, 5, 20, 100, ListView1.ListItems(Cont).SubItems(7), "F3", 7, hLeft
        oDoc.WTextBox Posi, 110, 20, 50, ListView1.ListItems(Cont).SubItems(8), "F3", 7, hCenter
        oDoc.WTextBox Posi, 164, 20, 70, Format(ListView1.ListItems(Cont).SubItems(9), "###,###,##0.00"), "F3", 7, hRight
        oDoc.WTextBox Posi, 228, 20, 50, ListView1.ListItems(Cont).SubItems(10), "F3", 7, hCenter
        Posi = Posi + 12
        If Posi >= 750 Then
            oDoc.NewPage A4_Vertical
            oDoc.WImage 70, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 205, 100, 175, tRs.Fields("NOMBRE"), "F3", 8, hCenter
            oDoc.WTextBox 60, 224, 100, 175, tRs.Fields("DIRECCION"), "F3", 8, hLeft
            oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs.Fields("COLONIA"), "F3", 8, hLeft
            oDoc.WTextBox 70, 205, 100, 175, tRs.Fields("ESTADO") & "," & tRs.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 80, 205, 100, 175, tRs.Fields("TELEFONO"), "F3", 8, hCenter
            oDoc.WTextBox 80, 340, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
            'CAJA1
            If Option4.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE LLEGAR", "F3", 10, hCenter
            End If
            If Option3.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE PAGO CON ENTRADA", "F3", 10, hCenter
            End If
            If Option5.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PENDIENTES DE PAGO CON ENTRADA PARCIAL", "F3", 10, hCenter
            End If
            If Option6.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PAGADAS", "F3", 10, hCenter
            End If
            If Option7.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA PAGADAS SIN ENTRADA", "F3", 10, hCenter
            End If
            If Option16.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA", "F3", 10, hCenter
            End If
            If Option14.Value Then
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE ORDENES DE COMPRA ACTIVAS", "F3", 10, hCenter
            End If
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 135
            oDoc.WLineTo 580, 135
            oDoc.LineStroke
            Posi = 135
            ' ENCABEZADO DEL DETALLE
        End If
        Cont = Cont + 1
    Loop
    ' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 760
    oDoc.WLineTo 580, 760
    oDoc.WTextBox 780, 324, 20, 120, "Fin del Reporte", "F2", 8, hCenter
    oDoc.LineStroke
    Posi = Posi + 6
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Image10_Click()
On Error GoTo ManejaError
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    TxtExcel.Text = ""
    Me.CommonDialog1.FileName = ""
    Me.CommonDialog1.DialogTitle = "Guardar como"
    Me.CommonDialog1.Filter = "Excel (*.xls)|*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If SSTab1.Tab = 0 Then
        If ListView1.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView1.ColumnHeaders.Count
                For Con = 1 To ListView1.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView1.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    If SSTab1.Tab = 1 Then
        If ListView2.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView2.ColumnHeaders.Count
                For Con = 1 To ListView2.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView2.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    If SSTab1.Tab = 2 Then
        If ListView3.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView3.ColumnHeaders.Count
                For Con = 1 To ListView3.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView3.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    If SSTab1.Tab = 3 Then
        If ListView6.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView6.ColumnHeaders.Count
                For Con = 1 To ListView6.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView6.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView6.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView6.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    If SSTab1.Tab = 4 Then
        If ListView11.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView11.ColumnHeaders.Count
                For Con = 1 To ListView11.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView11.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView11.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView11.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    If SSTab1.Tab = 5 Then
        If ListView15.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView15.ColumnHeaders.Count
                For Con = 1 To ListView15.ListItems.Count
                    If StrCopi = "" Then
                        For Con2 = 1 To NumColum
                            StrCopi = StrCopi & ListView15.ColumnHeaders(Con2).Text & Chr(9)
                        Next Con2
                        StrCopi = StrCopi & Chr(13)
                    End If
                    StrCopi = StrCopi & ListView15.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView15.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            End If
        End If
    End If
    TxtExcel.Text = StrCopi
    foo = FreeFile
    Open Ruta For Output As #foo
    Print #foo, TxtExcel.Text
    Close #foo
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
    Dim NoOC As String
    Dim TipoOC As String
    Dim cheque As String
    Dim fecha As String
    ConPag = 1
    Total = "0"
    SUMA = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Not (ListView1.ListItems.Count = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\ReporteOrdenesCompra.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image3.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image3, "Logo", False, False
        oDoc.NewPage A4_Horizontal
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 10, 100, 760, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 10, 100, 760, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 10, 100, 760, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 10, 100, 760, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 500, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        oDoc.WTextBox 40, 500, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 10, 100, 760, "REPORTE DE ESTADO DE ORDENES DE COMPRA", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 55, "No. Orden", "F2", 8, hCenter
        oDoc.WTextBox Posi, 60, 20, 200, "Proveedor", "F2", 8, hCenter
        oDoc.WTextBox Posi, 260, 20, 20, "Tipo", "F2", 8, hCenter
        oDoc.WTextBox Posi, 280, 20, 110, "Estado", "F2", 8, hCenter
        oDoc.WTextBox Posi, 390, 20, 50, "Subtotal", "F2", 8, hCenter
        oDoc.WTextBox Posi, 440, 20, 50, "IVA", "F2", 8, hCenter
        oDoc.WTextBox Posi, 490, 20, 50, "Flete", "F2", 8, hCenter
        oDoc.WTextBox Posi, 540, 20, 50, "Otros", "F2", 8, hCenter
        oDoc.WTextBox Posi, 590, 20, 50, "Total", "F2", 8, hCenter
        oDoc.WTextBox Posi, 640, 20, 80, "Cheque", "F2", 8, hCenter
        oDoc.WTextBox Posi, 720, 20, 50, "Fecha", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        For Cont = 1 To ListView1.ListItems.Count
            If NoOC <> ListView1.ListItems(Cont) Or TipoOC <> ListView1.ListItems(Cont).SubItems(5) Then
                oDoc.WTextBox Posi, 10, 20, 55, ListView1.ListItems(Cont), "F3", 7, hLeft
                oDoc.WTextBox Posi, 60, 20, 260, ListView1.ListItems(Cont).SubItems(1), "F3", 7, hLeft
                oDoc.WTextBox Posi, 260, 20, 20, ListView1.ListItems(Cont).SubItems(5), "F3", 7, hCenter
                oDoc.WTextBox Posi, 280, 20, 110, ListView1.ListItems(Cont).SubItems(6), "F3", 7, hLeft
                oDoc.WTextBox Posi, 390, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(3)), "###,###,##0.00"), "F3", 7, hRight
                'oDoc.WTextBox Posi, 390, 20, 50, Format(ListView1.ListItems(Cont).SubItems(3), "###,###,##0.00"), "F3", 7, hRight
                oDoc.WTextBox Posi, 440, 20, 50, Format(ListView1.ListItems(Cont).SubItems(11), "###,###,##0.00"), "F3", 7, hRight
                oDoc.WTextBox Posi, 490, 20, 50, Format(ListView1.ListItems(Cont).SubItems(12), "###,###,##0.00"), "F3", 7, hRight
                oDoc.WTextBox Posi, 540, 20, 50, Format(ListView1.ListItems(Cont).SubItems(13), "###,###,##0.00"), "F3", 7, hRight
                If ListView1.ListItems(Cont).SubItems(12) <> "" Then
                    If ListView1.ListItems(Cont).SubItems(3) = "" Then
                        oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(11)) + CDbl(ListView1.ListItems(Cont).SubItems(12)) + CDbl(ListView1.ListItems(Cont).SubItems(13)), "###,###,##0.00"), "F3", 7, hRight
                    Else
                        'oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(3)) - CDbl(ListView1.ListItems(Cont).SubItems(12)) - CDbl(ListView1.ListItems(Cont).SubItems(13)) - CDbl(ListView1.ListItems(Cont).SubItems(11)), "###,###,##0.00"), "F3", 7, hRight
                        oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(11)) + CDbl(ListView1.ListItems(Cont).SubItems(12)) + CDbl(ListView1.ListItems(Cont).SubItems(13)) + CDbl(ListView1.ListItems(Cont).SubItems(3)), "###,###,##0.00"), "F3", 7, hRight
                    End If
                    If ListView1.ListItems(Cont).SubItems(6) <> "CANCELADA" Then
                        Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(3)) + CDbl(ListView1.ListItems(Cont).SubItems(11)) + CDbl(ListView1.ListItems(Cont).SubItems(13))
                    End If
                Else
                    If IsNumeric(ListView1.ListItems(Cont).SubItems(3)) Then
                        'oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(13)) + CDbl(ListView1.ListItems(Cont).SubItems(15)) + CDbl(ListView1.ListItems(Cont).SubItems(3)), "###,###,##0.00"), "F3", 7, hRight
                        oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(3)), "###,###,##0.00"), "F3", 7, hRight
                        If ListView1.ListItems(Cont).SubItems(6) <> "CANCELADA" Then
                            Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(13)) + CDbl(ListView1.ListItems(Cont).SubItems(15)) + CDbl(ListView1.ListItems(Cont).SubItems(3))
                        End If
                    Else
                        If ListView1.ListItems(Cont).SubItems(15) = "" Then
                            oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(13)), "###,###,##0.00"), "F3", 7, hRight
                        Else
                            oDoc.WTextBox Posi, 590, 20, 50, Format(CDbl(ListView1.ListItems(Cont).SubItems(13)) + CDbl(ListView1.ListItems(Cont).SubItems(15)), "###,###,##0.00"), "F3", 7, hRight
                        End If
                        If ListView1.ListItems(Cont).SubItems(6) <> "CANCELADA" Then
                            If ListView1.ListItems(Cont).SubItems(15) = "" Then
                                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(13))
                            Else
                                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(13)) + CDbl(ListView1.ListItems(Cont).SubItems(15))
                            End If
                        End If
                    End If
                End If
                sBuscar = "SELECT * FROM CHEQUES WHERE (NUM_ORDEN LIKE '" & ListView1.ListItems(Cont) & ", %' OR NUM_ORDEN LIKE '% " & ListView1.ListItems(Cont) & ", %' ) AND TIPO_ORDEN LIKE '" & TipoOC & "%'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        cheque = cheque & " " & tRs.Fields("NUM_CHEQUE")
                        fecha = tRs.Fields("FECHA")
                        tRs.MoveNext
                    Loop
                    oDoc.WTextBox Posi, 640, 20, 80, cheque, "F3", 7, hCenter
                    oDoc.WTextBox Posi, 720, 20, 50, fecha, "F3", 7, hCenter
                End If
                cheque = ""
                Posi = Posi + 12
                If Posi >= 520 Then
                    oDoc.NewPage A4_Horizontal
                    oDoc.WImage 70, 40, 43, 161, "Logo"
                    sBuscar = "SELECT * FROM EMPRESA"
                    Set tRs1 = cnn.Execute(sBuscar)
                    oDoc.WTextBox 40, 10, 100, 760, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                    oDoc.WTextBox 60, 10, 100, 760, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
                    oDoc.WTextBox 70, 10, 100, 760, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                    oDoc.WTextBox 80, 10, 100, 760, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                    oDoc.WTextBox 30, 500, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                    oDoc.WTextBox 40, 500, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                    ' ENCABEZADO DEL DETALLE
                    oDoc.WTextBox 100, 10, 100, 760, "REPORTE DE ESTADO DE ORDENES DE COMPRA", "F3", 8, hCenter
                    Posi = 120
                    oDoc.WTextBox Posi, 10, 20, 55, "No. Orden", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 60, 20, 200, "Proveedor", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 260, 20, 20, "Tipo", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 280, 20, 110, "Estado", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 390, 20, 50, "Subtotal", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 440, 20, 50, "IVA", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 490, 20, 50, "Flete", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 540, 20, 50, "Otros", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 590, 20, 50, "Total", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 640, 20, 80, "Cheque", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 720, 20, 50, "Fecha", "F2", 8, hCenter
                    Posi = Posi + 12
                    ' Linea
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 760, Posi
                    oDoc.LineStroke
                    Posi = Posi + 6
                End If
                NoOC = ListView1.ListItems(Cont)
                TipoOC = ListView1.ListItems(Cont).SubItems(5)
            End If
        Next
        ' Linea
        Posi = Posi + 15
        oDoc.WTextBox Posi, 500, 20, 140, Format(Total, "###,###,##0.00"), "F3", 10, hRight
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
         Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 10, 100, 760, "COMENTARIOS", "F3", 8, hCenter
        Posi = Posi + 20
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text5.Text = Item
    Text6.Text = ListView1.SelectedItem.SubItems(5)
    frmordpendiente.oc = Item
    frmordpendiente.art = ListView1.SelectedItem.SubItems(1)
    frmordpendiente.fecha = ListView1.SelectedItem.SubItems(2)
    frmordpendiente.txttipo = ListView1.SelectedItem.SubItems(5)
    frmordpendiente.Show vbModal
End Sub
Private Sub ListView10_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = ListView10.SelectedItem.SubItems(1)
    Text2.SetFocus
End Sub
Private Sub ListView8_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item.SubItems(1)
    Text1.SetFocus
    IdProveedor = Item
    cmdBuscar.Value = True
End Sub
Private Sub ListView9_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProveedor1 = Item
End Sub
Private Sub ListView12_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProveedor2 = Item
End Sub
Private Sub Option1_Click()
    Label15.Visible = False
    Label11.Visible = True
End Sub
Private Sub SSTab1_dbClick()
    Image1.Visible = False
    Image10.Visible = False
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView4.ListItems.Clear
    pro = Item
    sBuscar = "SELECT * FROM ORDEN_RAPIDA_DETALLE WHERE  ID_ORDEN_RAPIDA= '" & pro & "' ORDER BY ID_PRODUCTO DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tLi.SubItems(3) = tRs.Fields("PRECIO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView5.ListItems.Clear
    sBuscar = "SELECT * FROM ORDEN_RAPIDA_DETALLE WHERE  ID_ORDEN_RAPIDA= '" & Item & "' ORDER BY ID_PRODUCTO DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tLi.SubItems(3) = tRs.Fields("PRECIO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Option2_Click()
    Label15.Visible = True
    Label11.Visible = False
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        Frame4.Visible = True
    Else
        Frame5.Visible = False
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView9.ListItems.Clear
        sBuscar = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text4.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView9.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub SSTab0_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        Frame11.Visible = True
        Frame4.Visible = True
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView8.ListItems.Clear
    If KeyAscii = 13 Then
        If Check6.Value = 1 Then
            cmdBuscar.Value = True
        Else
            sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
            Set tRs = cnn.Execute(sqlQuery)
            With tRs
                If Not (.BOF And .EOF) Then
                    Do While Not .EOF
                        Set tLi = ListView8.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                        If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                        .MoveNext
                    Loop
                End If
            End With
         End If
   End If
   Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Me.Command2.Value = True
   End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView10.ListItems.Clear
    If KeyAscii = 13 Then
        sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.BOF And .EOF) Then
                Do While Not .EOF
                    Set tLi = ListView10.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                    .MoveNext
                Loop
            End If
        End With
   End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command3.Value = True
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView12.ListItems.Clear
        sBuscar = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR_CONSUMO WHERE NOMBRE LIKE '%" & Text4.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView12.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command6.Value = True
   End If
End Sub
