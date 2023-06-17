VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form orpen 
   Caption         =   "Reporte de Ordenes Pendiente de Pago"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   18
      Top             =   2640
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "orpen.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "orpen.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   15
      Top             =   1320
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "orpen.frx":1EDC
         MousePointer    =   99  'Custom
         Picture         =   "orpen.frx":21E6
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   13
      Top             =   3960
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "orpen.frx":3D28
         MousePointer    =   99  'Custom
         Picture         =   "orpen.frx":4032
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "orpen.frx":6114
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   720
            TabIndex        =   6
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   58458113
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   7
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   58458113
            CurrentDate     =   39576
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Internacional."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Nacional"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Indirecta"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "orpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
