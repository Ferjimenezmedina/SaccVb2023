VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Programadas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Programadas"
   ClientHeight    =   8415
   ClientLeft      =   2580
   ClientTop       =   660
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   69
      Top             =   2280
      Width           =   975
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
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
         TabIndex        =   70
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "Programadas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":030A
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   67
      Top             =   3480
      Width           =   975
      Begin VB.Label Label20 
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
         TabIndex        =   68
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "Programadas.frx":2434
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":273E
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   66
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   8400
      TabIndex        =   65
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   63
      Top             =   4680
      Width           =   975
      Begin VB.Label Label19 
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
         TabIndex        =   64
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Programadas.frx":44C0
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":47CA
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   44
      Top             =   7080
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
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Programadas.frx":630C
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":6616
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8400
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   34
      Top             =   5880
      Width           =   975
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Programadas.frx":86F8
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":8A02
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   32
      Top             =   600
      Width           =   975
      Begin VB.Label Label11 
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "Programadas.frx":A3C4
         MousePointer    =   99  'Custom
         Picture         =   "Programadas.frx":A6CE
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Capturar"
      TabPicture(0)   =   "Programadas.frx":C2A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LvwProd"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ListView1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmdQuitar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CmdAceptar"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtCantidad"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtCanExis"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtClvProd"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Option2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Option1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtBusProd"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDes"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTipo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtMin"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text6"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtMarca"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Check1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCredito"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDescuento"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text7"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Command5"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Cerrar"
      TabPicture(1)   =   "Programadas.frx":C2BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Option6"
      Tab(1).Control(1)=   "Option7"
      Tab(1).Control(2)=   "Option8"
      Tab(1).Control(3)=   "Option11"
      Tab(1).Control(4)=   "Option12"
      Tab(1).Control(5)=   "Option13"
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(7)=   "TreeView1"
      Tab(1).Control(8)=   "Frame2"
      Tab(1).Control(9)=   "Command1"
      Tab(1).Control(10)=   "TxtNoPed"
      Tab(1).Control(11)=   "CmdCerrar"
      Tab(1).Control(12)=   "ELLI"
      Tab(1).Control(13)=   "Lvw2"
      Tab(1).Control(14)=   "Lvw1"
      Tab(1).Control(15)=   "Label10"
      Tab(1).Control(16)=   "Label9"
      Tab(1).Control(17)=   "Label8"
      Tab(1).ControlCount=   18
      Begin VB.CommandButton Command5 
         Caption         =   "Licitación"
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
         Left            =   6840
         Picture         =   "Programadas.frx":C2D8
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Efectivo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   76
         Top             =   6240
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Cheque"
         Height          =   195
         Left            =   -73920
         TabIndex        =   75
         Top             =   6240
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   "T. Credito"
         Height          =   195
         Left            =   -72960
         TabIndex        =   74
         Top             =   6240
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "T. Electrónica"
         Height          =   195
         Left            =   -70920
         TabIndex        =   73
         Top             =   6240
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "T. Debito"
         Height          =   195
         Left            =   -71880
         TabIndex        =   72
         Top             =   6240
         Width           =   975
      End
      Begin VB.OptionButton Option13 
         Caption         =   "No Aplica"
         Height          =   195
         Left            =   -69600
         TabIndex        =   71
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cancelar"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   57
         Top             =   6720
         Width           =   5775
         Begin VB.TextBox Text9 
            Height          =   735
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   59
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Cancelar"
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
            Picture         =   "Programadas.frx":ECAA
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Comentario"
            Height          =   255
            Left            =   1320
            TabIndex        =   62
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "No. Venta"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text7 
         Height          =   615
         Left            =   1320
         MaxLength       =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   7440
         Width           =   5415
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   30
         Left            =   -69600
         TabIndex        =   54
         Top             =   2280
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.TextBox txtDescuento 
         Height          =   285
         Left            =   6840
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCredito 
         Height          =   285
         Left            =   6840
         TabIndex        =   51
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Credito"
         Height          =   195
         Left            =   6840
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   7440
         TabIndex        =   49
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   48
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   7320
         TabIndex        =   46
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pedidos no cerrados"
         Height          =   1335
         Left            =   -68880
         TabIndex        =   40
         Top             =   6720
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CommandButton Command2 
            Caption         =   "Cancelar"
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
            Left            =   360
            Picture         =   "Programadas.frx":1167C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   360
            TabIndex        =   41
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Numero :"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
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
         Left            =   -68040
         Picture         =   "Programadas.frx":1404E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   7080
         TabIndex        =   38
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtDes 
         Height          =   285
         Left            =   6840
         TabIndex        =   31
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtNoPed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72720
         TabIndex        =   14
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Cerrar"
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
         Left            =   -71640
         Picture         =   "Programadas.frx":16A20
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton ELLI 
         Caption         =   "Eliminar"
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
         Left            =   -69240
         Picture         =   "Programadas.frx":193F2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtBusProd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   3000
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   2880
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripcion"
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TxtClvProd 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox TxtCanExis 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
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
         Left            =   6840
         Picture         =   "Programadas.frx":1BDC4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuitar 
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
         Left            =   6840
         Picture         =   "Programadas.frx":1E796
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
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
         Left            =   6840
         Picture         =   "Programadas.frx":21168
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7680
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50790401
         CurrentDate     =   38807
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2566
         LabelEdit       =   1
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
         Height          =   2055
         Left            =   240
         TabIndex        =   10
         Top             =   5280
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3625
         LabelEdit       =   1
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
      Begin MSComctlLib.ListView LvwProd 
         Height          =   1455
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2566
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Lvw2 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   13
         Top             =   3360
         Width           =   7935
         _ExtentX        =   13996
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
      Begin MSComctlLib.ListView Lvw1 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   12
         Top             =   840
         Width           =   7935
         _ExtentX        =   13996
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
      Begin VB.Label Label16 
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   7440
         Width           =   975
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   7560
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "No. Orden"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Pendientes de Cerrar"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "No. de Pedido Seleccionado :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Buscar Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "ID Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Existencia :"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad "
         Height          =   255
         Left            =   5040
         TabIndex        =   24
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Entrega"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. Venta"
      Height          =   255
      Left            =   8400
      TabIndex        =   36
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Programadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim SELE As String
Dim ind As Integer
Dim guia As String
Dim elind As String
Dim ClvVenta As String
Dim CapturoVenta As String
Dim ClienteVenta As String
Dim DesClien As String
Dim sRastreaProd As String
Dim IdPedido As String
Dim sRastreaNoPed As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As Double
    sBuscar = "SELECT P.ID_PRODUCTO, P.NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_PENDIENTE, ISNULL(CANTIDAD,0) AS CANTIDAD FROM PED_CLIEN AS C JOIN PED_CLIEN_DETALLE AS P ON C.NO_PEDIDO = P.NO_PEDIDO LEFT JOIN EXISTENCIAS AS E ON E.ID_PRODUCTO = P.ID_PRODUCTO AND E.SUCURSAL ='" & VarMen.Text4(0).Text & "' WHERE C.ESTADO = 'C' AND P.NO_PEDIDO = " & TxtNoPed.Text
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not (.EOF)
                cant = CDbl(.Fields("CANTIDAD_PEDIDA")) - CDbl(.Fields("CANTIDAD_PENDIENTE"))
                If cant > 0 Then
                    If CDbl(.Fields("CANTIDAD")) > 0 Then
                        sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(.Fields("CANTIDAD")) + CDbl(cant) & " WHERE ID_PRODUCTO = '" & .Fields("ID_PRODUCTO") & "' AND SUCURSAL ='" & VarMen.Text4(0).Text & "'"
                        cnn.Execute (sBuscar)
                    Else
                        sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & cant & ", '" & .Fields("ID_PRODUCTO") & "', '" & VarMen.Text4(0).Text & "');"
                        cnn.Execute (sBuscar)
                    End If
                End If
                .MoveNext
            Loop
            sBuscar = "UPDATE PED_CLIEN SET ESTADO = 'X' WHERE NO_PEDIDO = " & TxtNoPed.Text
            cnn.Execute (sBuscar)
        End If
    End With
    TxtNoPed.Text = ""
    Lvw2.ListItems.Clear
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As Double
    If Text5.Text <> "" Then
        sBuscar = "SELECT P.ID_PRODUCTO, P.NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_PENDIENTE, ISNULL(CANTIDAD,0) AS CANTIDAD FROM PED_CLIEN AS C JOIN PED_CLIEN_DETALLE AS P ON C.NO_PEDIDO = P.NO_PEDIDO LEFT JOIN EXISTENCIAS AS E ON E.ID_PRODUCTO = P.ID_PRODUCTO AND E.SUCURSAL = 'BODEGA' WHERE C.ESTADO = 'I' AND P.NO_PEDIDO = " & Text5.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If MsgBox("DESEA CANCELAR UNA VENTA NO CERRADA", vbYesNo, "SACC") = vbYes Then
                Do While Not (tRs.EOF)
                    cant = CDbl(tRs.Fields("CANTIDAD_PEDIDA")) - CDbl(tRs.Fields("CANTIDAD_PENDIENTE"))
                    If cant > 0 Then
                        If CDbl(tRs.Fields("CANTIDAD")) > 0 Then
                        '''08/0982009  este  upddate  es cuando   cancela la venta programada
                            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) + CDbl(cant) & " WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = 'BODEGA'"
                            cnn.Execute (sBuscar)
                        Else
                            sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & cant & ", '" & tRs.Fields("ID_PRODUCTO") & "', 'BODEGA');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                    tRs.MoveNext
                Loop
                sBuscar = "UPDATE PED_CLIEN SET ESTADO = 'X' WHERE NO_PEDIDO = " & Text5.Text
                cnn.Execute (sBuscar)
            End If
        Else
            MsgBox "LA VENTA NO EXISTE O YA FUE CERRADA", vbInformation, "SACC"
        End If
        Text5.Text = ""
    End If
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    TxtBusProd.Enabled = True
    LvwProd.Enabled = True
    CmdAceptar.Enabled = True
    TxtBusProd.SetFocus
    Command3.Enabled = False
    'Command5.Enabled = False
    Me.ListView1.Enabled = False
    DTPicker1.Enabled = False
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    If Text8 <> "" Then
        If Text9.Text <> "" Then
            If MsgBox("¿DESEA CANCELAR LA VENTA PROGRAMADA?. " & Text8.Text, vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO FROM ORDEN_COMPRA, ORDEN_COMPRA_DETALLE WHERE ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA AND ORDEN_COMPRA_DETALLE.NO_PEDIDO = " & Text8.Text
                Set tRs = cnn.Execute(sBuscar)
                If (tRs.EOF And tRs.BOF) Then
                    sBuscar = "SELECT * FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & Text8.Text
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        Do While Not (tRs.EOF)
                            sBusca = "SELECT * FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                            Set tRs2 = cnn.Execute(sBusca)
                            If Not (tRs2.EOF And tRs2.BOF) Then
                            ''''update   cuando
                                If CDbl(tRs.Fields("CANTIDAD_PENDIENTE")) = 0 Then
                                    sBusca = "UPDATE EXISTENCIAS SET CANTIDAD = " & (CDbl(tRs2.Fields("CANTIDAD"))) + (CDbl(tRs.Fields("CANTIDAD_PEDIDA"))) & " WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                                    cnn.Execute (sBusca)
                                Else
                                    If tRs.Fields("CANTIDAD_PENDIENTE") <> tRs.Fields("CANTIDAD_PEDIDA") Then
                                        sBusca = "UPDATE EXISTENCIAS SET CANTIDAD = " & (CDbl(tRs2.Fields("CANTIDAD"))) + (CDbl(tRs.Fields("CANTIDAD_PEDIDA")) - CDbl(tRs.Fields("CANTIDAD_PENDIENTE"))) & " WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                                        cnn.Execute (sBusca)
                                    End If
                                End If
                            Else
                                If CDbl(tRs.Fields("CANTIDAD_PENDIENTE")) = 0 Then
                                    sBusca = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & tRs.Fields("ID_PRODUCTO") & "', " & tRs.Fields("CANTIDAD_PEDIDA") & ", '" & VarMen.Text4(0).Text & "' );"
                                    cnn.Execute (sBusca)
                                Else
                                    If tRs.Fields("CANTIDAD_PENDIENTE") <> tRs.Fields("CANTIDAD_PEDIDA") Then
                                        sBusca = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & tRs.Fields("ID_PRODUCTO") & "', " & CDbl(tRs.Fields("CANTIDAD_PEDIDA")) - CDbl(tRs.Fields("CANTIDAD_PENDIENTE")) & ", '" & VarMen.Text4(0).Text & "' );"
                                        cnn.Execute (sBusca)
                                    End If
                                End If
                            End If
                            tRs.MoveNext
                        Loop
                    Else
                        MsgBox "LA VENTA PROGRAMADA #" & Text8.Text & " YA FUE CANCELADA ANTERIORMENTE", vbInformation, "SACC"
                    End If
                    sBuscar = "DELETE FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & Text8.Text & " "
                    cnn.Execute (sBuscar)
                    sBuscar = "DELETE FROM PED_CLIEN WHERE NO_PEDIDO = " & Text8.Text & " "
                    cnn.Execute (sBuscar)
                    sBuscar = "DELETE FROM REQUISICION WHERE NO_PEDIDO = " & Text8.Text & " "
                    cnn.Execute (sBuscar)
                    sBuscar = "DELETE FROM COTIZA_REQUI WHERE NO_PEDIDO = " & Text8.Text & " "
                    cnn.Execute (sBuscar)
                Else
                    If tRs.Fields("TIPO") = "N" Then
                        MsgBox "LA VENTA PROGRAMADA NO SE PUEDE CANCELAR DEBIDO A QUE TIENE PRODUCTOS EN LA ORDEN DE COMPRA " & tRs.Fields("NUM_ORDEN") & " NACIONAL", vbExclamation, "SACC"
                    Else
                        If tRs.Fields("TIPO") = "I" Then
                            MsgBox "LA VENTA PROGRAMADA NO SE PUEDE CANCELAR DEBIDO A QUE TIENE PRODUCTOS EN LA ORDEN DE COMPRA " & tRs.Fields("NUM_ORDEN") & " INTERNACIONAL", vbExclamation, "SACC"
                        Else
                            MsgBox "LA VENTA PROGRAMADA NO SE PUEDE CANCELAR DEBIDO A QUE TIENE PRODUCTOS EN LA ORDEN DE COMPRA " & tRs.Fields("NUM_ORDEN") & " INDIRECTA", vbExclamation, "SACC"
                        End If
                    End If
                End If
           End If
       Else
           MsgBox "INGRESAR  MOTIVO DE CANCELACION,COMENTARIO", vbInformation, "SACC"
       End If
    Else
       MsgBox "INGRESAR  MOTIVO DE CANCELACION,COMENTARIO", vbInformation, "SACC"
    End If
    Text8.Text = ""
    Text9.Text = ""
    Text8.SetFocus
End Sub
Private Sub Command5_Click()
    If SELE <> "" Then
        FrmLicitados.IdCliente = SELE
        FrmLicitados.Show vbModal
    Else
        MsgBox "Debe seleccionar un cliente", vbExclamation, "SACC"
    End If
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Label3.Caption = VarMen.Text4(0).Text
    If VarMen.Text1(49).Text = "N" Then
        Me.SSTab1.TabEnabled(0) = False
    End If
    If VarMen.Text1(50).Text = "N" Then
        Me.SSTab1.TabEnabled(1) = False
    End If
    Me.Command3.Enabled = False
    'Command5.Enabled = False
    Me.Text2.Text = VarMen.Text1(1).Text
    DTPicker1.Value = Format(Date + 1, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
        "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    If VarMen.Text1(77).Text = "N" Then
       Frame4.Enabled = False
       Text9.Enabled = False
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE DEL CLIENTE", 1800
        .ColumnHeaders.Add , , "CLIENTE", 7450
        .ColumnHeaders.Add , , "RFC", 2450
        .ColumnHeaders.Add , , "DIAS DE CREDITO", 500
        .ColumnHeaders.Add , , "CREDITO DISPONIBLE", 1500
        .ColumnHeaders.Add , , "DESCUENTO", 1500
    End With
    Me.CmdQuitar.Enabled = False
    Me.CmdAceptar.Enabled = False
    With LvwProd
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2700
        .ColumnHeaders.Add , , "Descripcion", 7450
        .ColumnHeaders.Add , , "TIPO", 0
        .ColumnHeaders.Add , , "CANT_MIN", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2700
        .ColumnHeaders.Add , , "CANTIDAD PEDIDA", 2400
        .ColumnHeaders.Add , , "CANTIDAD EN EXISTENCIA", 2400
        .ColumnHeaders.Add , , "CANTIDAD PENDIENTE", 2400
        .ColumnHeaders.Add , , "Descripcion", 0
        .ColumnHeaders.Add , , "TIPO", 0
        .ColumnHeaders.Add , , "CANT_MIN", 0
        .ColumnHeaders.Add , , "MARCA", 0
        .ColumnHeaders.Add , , "PRECIO", 0
    End With
    Me.ELLI.Enabled = False
    With Lvw1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Id Capturo", 0
        .ColumnHeaders.Add , , "Cliente", 6500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "No de Orden", 1500
        .ColumnHeaders.Add , , "Capturo", 2000
        .ColumnHeaders.Add , , "Id Cliente", 0
    End With
    With Lvw2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 4500
        .ColumnHeaders.Add , , "Cantidad Pedida", 2000
        .ColumnHeaders.Add , , "Cantidad en Existencia", 0
        .ColumnHeaders.Add , , "Cantidad Pendiente", 0
        .ColumnHeaders.Add , , "Precio Unitario", 2000
    End With
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub BuscarClien()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    If IsNumeric(Text1.Text) Then
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, DIAS_CREDITO, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE LIMITE_CREDITO <> 0 AND VALORACION = 'A' AND NOMBRE LIKE '%" & Text1.Text & "%' OR LIMITE_CREDITO <> 0 AND VALORACION = 'A' AND NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' OR LIMITE_CREDITO <> 0 AND VALORACION = 'A' AND ID_CLIENTE = '" & Text1.Text & "'"
    Else
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, DIAS_CREDITO, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE (LIMITE_CREDITO <> 0) AND (VALORACION = 'A') AND (NOMBRE LIKE '%" & Text1.Text & "%') OR (LIMITE_CREDITO <> 0) AND (VALORACION = 'A') AND (NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%')"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "NO SE ENCONTRO EL CLIENTE", vbInformation, "SACC"
        Else
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(2) = .Fields("RFC") & ""
                If Not IsNull(.Fields("DIAS_CREDITO")) Then tLi.SubItems(3) = .Fields("DIAS_CREDITO") & ""
                If Not IsNull(.Fields("DIAS_CREDITO")) Then
                    If .Fields("DIAS_CREDITO") <> "" Then
                        If CDbl(.Fields("DIAS_CREDITO")) > 0 Then
                            sBuscar = "SELECT ISNULL(SUM(TOTAL_COMPRA), 0) AS TOTAL FROM CUENTAS WHERE ID_CLIENTE = " & .Fields("ID_CLIENTE")
                            Set tRs2 = cnn.Execute(sBuscar)
                            sBuscar = "SELECT ISNULL(SUM(CANT_ABONO), 0) AS ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & .Fields("ID_CLIENTE")
                            Set tRs1 = cnn.Execute(sBuscar)
                            If Not IsNull(.Fields("LIMITE_CREDITO")) And Not IsNull(tRs2.Fields("TOTAL")) Then tLi.SubItems(4) = CDbl(.Fields("LIMITE_CREDITO")) - CDbl(tRs2.Fields("TOTAL")) + CDbl(tRs1.Fields("abono"))
                            tRs2.Close
                        Else
                            tLi.SubItems(4) = 0
                        End If
                    End If
                End If
                If Not IsNull(.Fields("DESCUENTO")) Then
                    If Not IsNull(.Fields("DESCUENTO")) Then tLi.SubItems(5) = .Fields("DESCUENTO")
                Else
                    tLi.SubItems(5) = 0
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
    Imprimir_Recibo
End Sub
Private Sub Image10_Click()
    If ListView2.ListItems.Count > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Text10.Text = Text1.Text & Chr(13) & Date & Chr(13)
        Ruta = Me.CommonDialog1.FileName
        If ListView2.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView2.ColumnHeaders.Count
                For Con = 1 To ListView2.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
                Next
                ProgressBar1.Value = 0
                ProgressBar1.Visible = True
                ProgressBar1.Min = 0
                ProgressBar1.Max = ListView2.ListItems.Count
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView2.ListItems.Count
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                    ProgressBar1.Value = Con
                Next
                'archivo TXT
                Dim foo As Integer
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    End If
End Sub
Private Sub Image3_Click()
    FrmRepVentProg.Show vbModal
End Sub
Private Sub Image7_Click()
    frmShowPediC.Command1.Visible = False
    frmShowPediC.Command4.Visible = False
    frmShowPediC.Command5.Visible = False
    frmShowPediC.Label4.Visible = False
    frmShowPediC.Combo1.Visible = False
    frmShowPediC.Show vbModal
End Sub
Private Sub Image8_Click()
0 On Error GoTo ManejaError
    Me.Text2.Text = VarMen.Text1(0).Text
    If SELE <> "" And Text2.Text <> "" And ListView2.ListItems.Count > 0 Then
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim tRs As ADODB.Recordset
        Dim NumeroRegistros As Integer
        Dim Urgente As String
        Dim Pedido As String
        NumeroRegistros = ListView2.ListItems.Count
        Dim Conta As Integer
        Dim Proceder As Boolean
        Dim TotVenta As Double
        Dim IVA As Double
        Dim IdRequi As String
        Dim CanPed As String
        Dim NoRe As Integer
        Dim Cont As Integer
        Dim nComanda As Integer
        Dim cTipo As String
        Proceder = True
        CanPed = 0
        If Check1.Value = 1 Then
            TotVenta = 0
            For Conta = 1 To NumeroRegistros
                TotVenta = TotVenta + (Val(Replace(ListView2.ListItems.Item(Conta).SubItems(1), ",", "")) * Val(Replace(ListView2.ListItems.Item(Conta).SubItems(8), ",", "")))
            Next Conta
            IVA = TotVenta * CDbl(CDbl(VarMen.Text4(7).Text) / 100)
            If (TotVenta + IVA) > Val(Replace(txtCredito.Text, ",", "")) Then
                Proceder = False
            End If
        End If
        If Proceder Then
            sBuscar = "INSERT INTO PED_CLIEN (ID_CLIENTE, USUARIO, FECHA, ESTADO, FECHA_CAPTURA, NO_ORDEN, COMENTARIO) VALUES (" & SELE & ", '" & VarMen.Text1(0).Text & "', '" & DTPicker1.Value & "', 'I', DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), '" & Text6.Text & "','V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT NO_PEDIDO FROM PED_CLIEN ORDER BY NO_PEDIDO DESC"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then Pedido = tRs.Fields("NO_PEDIDO")
            Text4.Text = Pedido
            IdPedido = Pedido
            tRs.Close
            Conta = 1
            For Conta = 1 To NumeroRegistros
                sBuscar = "INSERT INTO PED_CLIEN_DETALLE (ID_PRODUCTO, NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_EXISTENCIA, CANTIDAD_PENDIENTE) VALUES ('" & ListView2.ListItems(Conta) & "', " & Pedido & ", " & CDbl(ListView2.ListItems(Conta).SubItems(1)) & ", " & CDbl(ListView2.ListItems(Conta).SubItems(2)) & ", " & CDbl(ListView2.ListItems(Conta).SubItems(1)) & ");"
                cnn.Execute (sBuscar)
                ' SE SUBE TODO... SI QUISIERA SOLO SUBIRSE LO PENDIENTE AQUI SE CHECARIA LA CANTIDAD PENDIENTE...
                ' CAMBIO POR ARMANDO H VALDEZ ARRAS A 12/OCT/2010
                If ListView2.ListItems(Conta).SubItems(5) = "COMPUESTO" Then
                    ' Modo anterior... guardaba el pedido de los productos compuestos
                    ' para que almacen los subiera en Orden de produccion
                    'sBuscar = "SELECT ID_PEDIDO FROM PEDIDO WHERE SUCURSAL = '" & Label3.Caption & "' ORDER BY ID_PEDIDO DESC"
                    'Set tRs = cnn.Execute(sBuscar)
                    'sBuscar = "INSERT INTO DETALLE_PEDIDO (ID_PEDIDO, CANTIDAD, ID_PRODUCTO, ENTREGADO, DESCRIPCION, ALMACEN, MARCA) VALUES ('" & tRs.Fields("ID_PEDIDO") & "', '" & ListView2.ListItems(Conta).SubItems(3) & "', '" & ListView2.ListItems(Conta) & "', 0, '" & ListView2.ListItems(Conta).SubItems(4) & "', 'A3', '" & ListView2.ListItems(Conta).SubItems(7) & "')"
                    'cnn.Execute (sBuscar)
                    ' Modo nuevo... Crea la orden de produccion al momento de guardar el pedido
                    ' CAMBIO POR ARMANDO H VALDEZ ARRAS A 11/NOV/2011
                    sBuscar = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_AGENTE, ID_SUCURSAL, TIPO, COMENTARIO, SUCURSAL) VALUES (DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & VarMen.Text1(0).Text & ", " & VarMen.Text1(5).Text & ", 'P', 'V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "', '" & VarMen.Text4(0).Text & "')"
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    nComanda = tRs.Fields("ID_COMANDA")
                    If Mid(ListView2.ListItems(Conta), 3, 1) = "T" Then
                        cTipo = "T" 'Toner
                    Else
                        If Mid(ListView2.ListItems(Conta), 3, 1) = "I" Then
                            cTipo = "I" 'Tinta
                        Else
                            cTipo = "X" 'Error
                        End If
                    End If
                    Cont = Cont + 1
                    sBuscar = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO, CLASIFICACION) VALUES (" & nComanda & ", " & Cont & ", '" & ListView2.ListItems(Conta) & "', " & ListView2.ListItems(Conta).SubItems(1) & ", '" & cTipo & "','P');"
                    cnn.Execute (sBuscar)
                    sBuscar = "INSERT INTO PRODPEND (ID_COMANDA, ARTICULO) VALUES (" & nComanda & ", " & Cont & ");"
                    cnn.Execute (sBuscar)
                    Imprimir_Ticket (nComanda)
                    Imprimir_Ticket (nComanda)
                Else
                    sBuscar = "SELECT ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0 AND CONTADOR = 0 AND COTIZADA = 0 AND ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If (tRs.BOF And tRs.EOF) Or (DTPicker1.Value <= (Date + 10)) Then
                        Urgente = "N"
                        'Pide todo lo encargado :: Modificado el 22/Sep/2010
                        'Para pedir solo los pendientes reemplazar el ListView2.ListItems(Conta).SubItems(1) por el ListView2.ListItems(Conta).SubItems(3)
                        If DTPicker1.Value <= (Date + 10) Then
                            Urgente = "S"
                        End If
                        'QUITA COMENTARIOS DEL SIGUIENTE IF SI QUIERES QUE SOLO SE SUBA LO QUE NO HAY EN EXISTENCIA
                        'If CDbl(ListView2.ListItems(Conta).SubItems(3)) > 0 Then
                            ' ACTUALIZACION QUE CHECA LA CANTIDAD DISPONIBLE DE LA EXISTENCIA SEGUN PEDIDOS QUE AUN NO
                            ' TIENEN APARTADO Y PIDE SOLO LO NECESARIO PARA SURTIR
                            sBuscar = "SELECT PED_CLIEN_DETALLE.ID_PRODUCTO, EXISTENCIAS.CANTIDAD, SUM(PED_CLIEN_DETALLE.CANTIDAD_PENDIENTE) AS PENDIENTES FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN EXISTENCIAS ON PED_CLIEN_DETALLE.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (EXISTENCIAS.SUCURSAL = 'BODEGA') AND (PED_CLIEN.ESTADO NOT IN ('X', 'T', 'C')) AND PED_CLIEN_DETALLE.ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "' GROUP BY PED_CLIEN_DETALLE.ID_PRODUCTO, EXISTENCIAS.CANTIDAD"
                            Set tRs = cnn.Execute(sBuscar)
                            If Not (tRs.EOF And tRs.BOF) Then
                                'If (tRs.Fields("CANTIDAD") > tRs.Fields("PENDIENTES")) Then
                                '    CanPed = ListView2.ListItems(Conta).SubItems(1) - (tRs.Fields("CANTIDAD") - tRs.Fields("PENDIENTES"))
                                'Else
                                    CanPed = ListView2.ListItems(Conta).SubItems(1)
                                'End If
                            Else
                                CanPed = ListView2.ListItems(Conta).SubItems(1)
                            End If
                            sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, ACTIVO, CONTADOR, COTIZADA, ALMACEN, URGENTE, MARCA, COMENTARIO, NO_PEDIDO) VALUES (DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), '" & ListView2.ListItems(Conta) & "' , '" & ListView2.ListItems(Conta).SubItems(4) & "'," & CDbl(CanPed) & ", 0, 0, 0, 'A3', '" & Urgente & "', '" & ListView2.ListItems(Conta).SubItems(7) & "','V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "', " & IdPedido & ")"
                            cnn.Execute (sBuscar)
                            sBuscar = "SELECT ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0 AND URGENTE = 'S' AND ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "'"
                            Set tRs = cnn.Execute(sBuscar)
                            If Not (tRs.EOF And tRs.BOF) Then
                                sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & CanPed & ", 'V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "', DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"
                                cnn.Execute (sBuscar2)
                            End If
                        'End If
                    Else
                        ' ACTUALIZACION QUE CHECA LA CAN TIDAD DISPONIBLE DE LA EXISTENCIA SEGUN PEDIDOS QUE AUN NO
                        ' TIENEN APARTADO Y PIDE SOLO LO NECESARIO PARA SURTIR
                        sBuscar = "SELECT PED_CLIEN_DETALLE.ID_PRODUCTO, EXISTENCIAS.CANTIDAD, SUM(PED_CLIEN_DETALLE.CANTIDAD_PENDIENTE) AS PENDIENTES FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN EXISTENCIAS ON PED_CLIEN_DETALLE.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (EXISTENCIAS.SUCURSAL = 'BODEGA') AND (PED_CLIEN.ESTADO NOT IN ('X', 'T', 'C')) AND PED_CLIEN_DETALLE.ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "' GROUP BY PED_CLIEN_DETALLE.ID_PRODUCTO, EXISTENCIAS.CANTIDAD"
                        Set tRs = cnn.Execute(sBuscar)
                        If (tRs.Fields("CANTIDAD") > tRs.Fields("PENDIENTES")) Then
                            CanPed = ListView2.ListItems(Conta).SubItems(1) - (tRs.Fields("CANTIDAD") - tRs.Fields("PENDIENTES"))
                        Else
                            CanPed = ListView2.ListItems(Conta).SubItems(1)
                        End If
                        IdRequi = tRs.Fields("ID_REQUISICION")
                        sBuscar = "UPDATE REQUISICION SET CANTIDAD = CANTIDAD + " & CanPed & ", NO_PEDIDO = " & IdPedido & " WHERE ID_REQUISICION = " & tRs.Fields("ID_REQUISICION") & " AND ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "'"
                        Set tRs = cnn.Execute(sBuscar)
                        sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & IdRequi & ", '" & VarMen.Text1(1).Text & "'  ," & CanPed & ", 'V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "', DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"
                        cnn.Execute (sBuscar2)
                    End If
                    If sBuscar2 = "" Then
                        sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
                        Set tRs = cnn.Execute(sBuscar)
                        sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & CanPed & ", 'V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text & "',DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())))"
                        cnn.Execute (sBuscar2)
                        tRs.Close
                    End If
                End If
                'APARTADO DE EXISTENCIAS (CANCELADO TEMPORALMENTE POR PEDIDO DEL LIC. EL DIA 19/AGOSTO/2010 POR ARMANDO H VALDEZ ARRAS
                'If (Val(Replace(ListView2.ListItems(Conta).SubItems(1), ",", "")) - Val(Replace(ListView2.ListItems(Conta).SubItems(3), ",", ""))) > 0 Then
                '    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & Val(Replace(ListView2.ListItems(Conta).SubItems(1), ",", "")) - Val(Replace(ListView2.ListItems(Conta).SubItems(3), ",", "")) & " WHERE ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "' AND SUCURSAL = 'BODEGA'"
                '    cnn.Execute (sBuscar)
                'End If
            Next Conta
            SELE = ""
            Text1.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            TxtClvProd.Text = ""
            txtDes.Text = ""
            txtTipo.Text = ""
            txtMin.Text = ""
            TxtBusProd.Text = ""
            txtMarca.Text = ""
            Check1.Value = 1
            txtCredito.Text = ""
            txtDescuento.Text = ""
            TxtCantidad.Text = ""
            TxtBusProd.Enabled = False
            Me.Command3.Enabled = True
            LvwProd.Enabled = False
            ListView1.Enabled = True
            ListView1.ListItems.Clear
            LvwProd.ListItems.Clear
            ListView2.ListItems.Clear
            Command3.Enabled = True
            Command5.Enabled = True
            CmdAceptar.Enabled = False
            DTPicker1.Enabled = True
            DTPicker1.Value = Date
            Imprimir_Recibo
            'Unload Me
        Else
            MsgBox "EL TOTAL SUPERA EL LIMITE DE CREDITO, NO SE PUEDE CERRAR LA VENTA", vbInformation, "SACC"
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    If ListView2.ListItems.Count = 0 Then
        Unload Me
    Else
        MsgBox "TIENE UNA CAPTURA PENDIENTE, DEBE ELIMINAR O GUARDAR LA CAPTURA PARA PODER SALIR", vbExclamation, "SACC"
    End If
End Sub
Private Sub ListView1_DblClick()
    Command3.Value = True
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SELE = Item
    Text1.Text = Item.SubItems(1)
    Text7.Text = Item.SubItems(1)
    If Val(Item.SubItems(3)) > 0 Then
        Check1.Value = 1
        txtCredito.Text = Item.SubItems(4)
    Else
        Check1.Value = 0
        txtCredito.Text = 0
    End If
    txtDescuento.Text = Item.SubItems(5)
    Me.Command3.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And SELE <> "" Then
        Me.Command3.Enabled = True
        Command5.Enabled = True
        Me.Command3.SetFocus
    End If
End Sub
Private Sub Lvw2_DblClick()
    FrmRastreoVentProg.Text2.Text = sRastreaProd
    FrmRastreoVentProg.Text1.Text = sRastreaNoPed
    FrmRastreoVentProg.Show vbModal
End Sub
Private Sub Lvw2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    sRastreaProd = Item
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscarClien
        Me.ListView1.SetFocus
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
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Command4.Value = True
    End If
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++ < Cerrar venta Prog. > +++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub CmdCerrar_Click()
'/luis
On Error GoTo ManejaError
    Dim ClvCliente As String
    Dim NomCliente As String
    Dim Clasificacion As String
    Dim TotVenta As String
    Dim TotLicitacion As String
    Dim DesClente As String
    Dim fecha As String
    Dim ClvUsuario As String
    Dim ClvProducto As String
    Dim DesProducto As String
    Dim CantProducto As String
    Dim PreVenta As String
    Dim precosto As String
    Dim GanProducto As String
    Dim sBuscar As String
    Dim IVA As String
    Dim cant As Integer
    Dim i As Integer
    Dim IdCta As Integer
    Dim DiasC As Integer
    Dim LimiteC As Double
    Dim DES As Double
    Dim DES2 As Double
    Dim Proceder As Boolean
    Dim SiLicitacion As Boolean
    Dim IdDescuento As String
    Dim IdDescuento2 As String
    'Dim IVA As String
    Dim IMPUESTO1 As String
    Dim IMPUESTO2 As String
    Dim RETENCION As String
    Dim TPago As String
    Dim FormaPagoSAT As String
    DesClente = "0"
    TotLicitacion = "0"
    IdDescuento = ""
    IdDescuento2 = "0"
    TotVenta = "0"
    PreVenta = "0"
    precosto = "0"
    GanProducto = "0"
    CantProducto = "0"
    IMPUESTO1 = "0"
    IMPUESTO2 = "0"
    RETENCION = "0"
    IVA = "0"
    SiLicitacion = False
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Proceder = True
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
    sBuscar = "SELECT ID_CLIENTE, USUARIO FROM PED_CLIEN WHERE NO_PEDIDO = " & TxtNoPed.Text
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        If CapturoVenta = "" Then
            ClvUsuario = VarMen.Text1(0).Text
        Else
            ClvUsuario = CapturoVenta
        End If
        sBuscar = "SELECT ID_CLIENTE, DESCUENTO, ID_DESCUENTO FROM CLIENTES WHERE NOMBRE = '" & ClienteVenta & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            ClvCliente = tRs.Fields("ID_CLIENTE")
            If Not IsNull(tRs.Fields("DESCUENTO")) Then DesClente = tRs.Fields("DESCUENTO")
            If Not IsNull(tRs.Fields("ID_DESCUENTO")) Then IdDescuento = tRs.Fields("ID_DESCUENTO")
        End If
    Else
        ClvCliente = tRs.Fields("ID_CLIENTE")
        ClvUsuario = tRs.Fields("USUARIO")
    End If
    sBuscar = "SELECT NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO, ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & ClvCliente
    Set tRs = cnn.Execute(sBuscar)
    If tRs.Fields("DESCUENTO") <> "" Then
        DesClente = tRs.Fields("DESCUENTO")
    Else
        DesClente = "0"
    End If
    If Not IsNull(tRs.Fields("ID_DESCUENTO")) Then IdDescuento = tRs.Fields("ID_DESCUENTO")
    NomCliente = tRs.Fields("NOMBRE")
    If Not IsNull(tRs.Fields("DIAS_CREDITO")) Then
        If tRs.Fields("DIAS_CREDITO") <> "" Then
            DiasC = tRs.Fields("DIAS_CREDITO")
        End If
    End If
    LimiteC = tRs.Fields("LIMITE_CREDITO")
    fecha = Format(Date, "dd/mm/yyyy")
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD_PEDIDA FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & TxtNoPed.Text & " AND FINALIZADA NOT IN ('E')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE PED_CLIEN_DETALLE SET FINALIZADA = 'S' WHERE NO_PEDIDO = " & TxtNoPed.Text & " AND FINALIZADA  NOT IN ('E')"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & ClvCliente
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.BOF And tRs2.EOF) Then
                sBuscar = "SELECT PRECIO_COSTO, GANANCIA, CLASIFICACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2, RETENCION FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                IVA = 0
                'TotVenta = Format(Val(Replace(TotVenta, ",", "")) + (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) * CDbl(tRs.Fields("CANTIDAD_PEDIDA")), "0.00")
                IVA = CDbl(IVA) + (CDbl(tRs1.Fields("IVA")) * TotVenta)
                IMPUESTO1 = CDbl(IMPUESTO1) + (CDbl(tRs1.Fields("IMPUESTO1")) * TotVenta)
                IMPUESTO2 = CDbl(IMPUESTO2) + (CDbl(tRs1.Fields("IMPUESTO2")) * TotVenta)
                RETENCION = CDbl(RETENCION) + (CDbl(tRs1.Fields("RETENCION")) * TotVenta)
                TotVenta = Format(CDbl(TotLicitacion) + (CDbl(tRs2.Fields("PRECIO_VENTA")) * (1 - CDbl(DesClente))) * CDbl(tRs.Fields("CANTIDAD_PEDIDA")), "0.00")
                SiLicitacion = True
                DES = 0
                DES2 = 0
            Else
                sBuscar = "SELECT PRECIO_COSTO, GANANCIA, CLASIFICACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2, RETENCION FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                TotVenta = Format(Val(Replace(TotVenta, ",", "")) + (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) * CDbl(tRs.Fields("CANTIDAD_PEDIDA")), "0.00")
                IVA = CDbl(IVA) + (CDbl(tRs1.Fields("IVA")) * TotVenta)
                IMPUESTO1 = CDbl(IMPUESTO1) + (CDbl(tRs1.Fields("IMPUESTO1")) * TotVenta)
                IMPUESTO2 = CDbl(IMPUESTO2) + (CDbl(tRs1.Fields("IMPUESTO2")) * TotVenta)
                RETENCION = CDbl(RETENCION) + (CDbl(tRs1.Fields("RETENCION")) * TotVenta)
                If (DesClente <> "" Or DesClente > 0) Then
                    DES = Val(Replace(TotVenta, ",", "")) * (Val(Replace(DesClente, ",", "")) / 100)
                    If (IdDescuento <> "" And IdDescuento <> "0") Then
                        sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & IdDescuento & "' AND CLASIFICACION = '" & tRs1.Fields("CLASIFICACION") & "'"
                        Set tRs5 = cnn.Execute(sBuscar)
                        If Not IsNull(tRs5.Fields("PORCENTAJE")) Then IdDescuento2 = tRs5.Fields("PORCENTAJE")
                        DES2 = Val(Replace(TotVenta, ",", "")) * (Val(Replace(IdDescuento2, ",", "")) / 100)
                    Else
                        DES2 = 0
                    End If
                Else
                    DES = 0
                End If
            End If
            tRs.MoveNext
        Loop
    End If
    If DES2 > DES Then
        DES = DES2
    End If
    If SiLicitacion = False Then
        TotVenta = Val(Replace(TotVenta, ",", "")) - DES
    End If
    SiLicitacion = False
    DES = 0
    DES2 = 0
    TotVenta = Replace(TotVenta, ",", "")
    IVA = Format(Val(TotVenta) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
    If DiasC > 0 Then
        sBuscar = "SELECT SUM(TOTAL_COMPRA) AS TOTAL FROM CUENTAS WHERE ID_CLIENTE = " & ClvCliente
        Set tRs2 = cnn.Execute(sBuscar)
        sBuscar = "SELECT SUM(CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & ClvCliente
        Set tRs3 = cnn.Execute(sBuscar)
        If Not IsNull(tRs2.Fields("TOTAL")) Then LimiteC = LimiteC - CDbl(tRs2.Fields("TOTAL"))
        If Not IsNull(tRs3.Fields("TOTAL")) Then LimiteC = LimiteC + CDbl(tRs3.Fields("TOTAL"))
        tRs2.Close
    End If
    If (Val(TotVenta) + Val(IVA)) > LimiteC And DiasC > 0 Then
        Proceder = False
    End If
    If Proceder = False Then
         MsgBox "El cliente no tiene límite de crédito disponible, no se puede cerrar la venta a crédito?", vbExclamation, "SACC"
    End If
    If Proceder Then
        If DiasC > 0 Then
            sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, IVA, SUBTOTAL, TOTAL, DESCUENTO, ID_USUARIO, FECHA, SUCURSAL, DIAS_CREDITO, FECHA_VENCE, UNA_EXIBICION, TIPO_PAGO, FORMA_PAGO, FormaPagoSAT) VALUES (" & ClvCliente & ", '" & NomCliente & "', " & Replace(IVA, ",", "") & ", " & Replace(TotVenta, ",", "") & ", " & Replace(CDbl(TotVenta) + CDbl(IVA), ",", "") & ", " & Replace(DesClente, ",", "") & ", '" & ClvUsuario & "', '" & fecha & "', '" & VarMen.Text4(0).Text & "', " & DiasC & ", '" & Format(Date + DiasC, "dd/mm/yyyy") & "', 'N', '" & TPago & "', 'PAGO EN PARCIALIDADES', '" & FormaPagoSAT & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_VENTA FROM VENTAS ORDER BY ID_VENTA DESC"
            Set tRs = cnn.Execute(sBuscar)
            ClvVenta = tRs.Fields("ID_VENTA")
            sBuscar = "INSERT INTO CUENTAS (PAGADA, ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, SUCURSAL, TOTAL_COMPRA, DEUDA, ID_VENTA) VALUES ( 'N', " & ClvCliente & ", '" & VarMen.Text1(0).Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & DiasC & ", '" & Format(Date + DiasC, "dd/mm/yyyy") & "', " & DesClente & ", '" & VarMen.Text4(0).Text & "', " & Replace(CDbl(TotVenta) + CDbl(IVA), ",", "") & ", " & Replace(Val(Replace(TotVenta, ",", "")) + Val(Replace(IVA, ",", "")), ",", "") & ", " & ClvVenta & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_CUENTA FROM CUENTAS ORDER BY ID_CUENTA DESC"
            Set tRs = cnn.Execute(sBuscar)
            IdCta = tRs.Fields("ID_CUENTA")
            sBuscar = "INSERT INTO CUENTA_VENTA (ID_VENTA, ID_CUENTA) VALUES (" & ClvVenta & ", " & IdCta & ");"
            cnn.Execute (sBuscar)
        Else
            sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, IVA, SUBTOTAL, TOTAL, DESCUENTO, ID_USUARIO, FECHA, SUCURSAL, DIAS_CREDITO, FECHA_VENCE, UNA_EXIBICION, FORMA_PAGO, TIPO_PAGO, FormaPagoSAT) VALUES (" & ClvCliente & ", '" & NomCliente & "', " & Replace(IVA, ",", "") & ", " & Replace(TotVenta, ",", "") & ", " & Replace(CDbl(TotVenta) + CDbl(IVA), ",", "") & ", " & Replace(DesClente, ",", "") & ", '" & ClvUsuario & "', '" & fecha & "', '" & VarMen.Text4(0).Text & "', 0, '', 'S','PAGO EN UNA SOLA EXHIBICION', '" & TPago & "', '" & FormaPagoSAT & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_VENTA FROM VENTAS ORDER BY ID_VENTA DESC"
            Set tRs = cnn.Execute(sBuscar)
            ClvVenta = tRs.Fields("ID_VENTA")
        End If
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD_PEDIDA, Descripcion FROM vsPEDCDA3 WHERE NO_PEDIDO = " & TxtNoPed.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                ClvProducto = tRs.Fields("ID_PRODUCTO")
                CantProducto = tRs.Fields("CANTIDAD_PEDIDA")
                DesProducto = tRs.Fields("Descripcion")
                DesProducto = Replace(DesProducto, ",", "")
                sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & ClvCliente
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    PreVenta = Format(tRs2.Fields("PRECIO_VENTA"), "0.00")
                    precosto = tRs2.Fields("PRECIO_VENTA")
                    GanProducto = "0"
                    SiLicitacion = True
                Else
                    sBuscar = "SELECT PRECIO_COSTO, GANANCIA, CLASIFICACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        PreVenta = Format(CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA"))), "0.00")
                        precosto = tRs1.Fields("PRECIO_COSTO")
                        IVA = CDbl(tRs1.Fields("IVA")) * PreVenta
                        IMPUESTO1 = CDbl(tRs1.Fields("IMPUESTO1")) * PreVenta
                        IMPUESTO2 = CDbl(tRs1.Fields("IMPUESTO2")) * PreVenta
                        RETENCION = CDbl(tRs1.Fields("P_RETENCION")) * PreVenta
                        GanProducto = tRs1.Fields("GANANCIA")
                        If (DesClente <> "" Or DesClente > 0) Then
                            DES = Val(Replace(PreVenta, ",", "")) * (Val(Replace(DesClente, ",", "")) / 100)
                            If (IdDescuento <> "") Then
                                sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & IdDescuento & "' AND CLASIFICACION = '" & tRs1.Fields("CLASIFICACION") & "'"
                                Set tRs5 = cnn.Execute(sBuscar)
                                If Not (tRs5.EOF And tRs5.BOF) Then
                                    IdDescuento2 = tRs5.Fields("PORCENTAJE")
                                    DES2 = Val(Replace(PreVenta, ",", "")) * (Val(Replace(IdDescuento2, ",", "")) / 100)
                                Else
                                    DES2 = 0
                                End If
                            Else
                                DES2 = 0
                            End If
                        Else
                            DES = 0
                        End If
                    End If
                    If DES2 > DES Then
                        DES = DES2
                    End If
                    If SiLicitacion = False Then
                        PreVenta = Val(Replace(PreVenta, ",", "")) - DES
                    End If
                    If DiasC > 0 Then
                        sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, CANTIDAD, ID_PRODUCTO, PRECIO_VENTA) VALUES (" & IdCta & ", " & CantProducto & ", '" & ClvProducto & "', " & Replace(Val(Replace(PreVenta, ",", "")) * Val(CantProducto), ",", "") & ");"
                        cnn.Execute (sBuscar)
                    End If
                End If
                sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE, IVA, IMPUESTO1, IMPUESTO2, RETENCION) VALUES (" & ClvVenta & ", '" & ClvProducto & "', '" & Replace(DesProducto, ",", "") & "', " & Replace(CantProducto, ",", "") & ", " & Replace(PreVenta, ",", "") & ", " & Replace(precosto, ",", "") & ", " & Replace(GanProducto, ",", "") & ", " & CDbl(PreVenta) * CDbl(CantProducto) & ", " & CDbl(IVA) & ", " & CDbl(IMPUESTO1) & ", " & CDbl(IMPUESTO2) & ", " & CDbl(RETENCION) & ");"
                cnn.Execute (sBuscar)
                tRs.MoveNext
            Loop
            sBuscar = "SELECT PED_CLIEN_DETALLE.NO_PEDIDO, SUM((PED_CLIEN_DETALLE.CANTIDAD_PEDIDA * ALMACEN3.PRECIO_COSTO) * (ALMACEN3.GANANCIA + 1)) AS T_VENTA, SUM((PED_CLIEN_DETALLE.CANTIDAD_PEDIDA * ALMACEN3.PRECIO_COSTO) * (ALMACEN3.GANANCIA + 1) * ALMACEN3.IVA) AS T_IVA, SUM((PED_CLIEN_DETALLE.CANTIDAD_PEDIDA * ALMACEN3.PRECIO_COSTO) * (ALMACEN3.GANANCIA + 1) * ALMACEN3.IMPUESTO1) AS T_IMP1, SUM((PED_CLIEN_DETALLE.CANTIDAD_PEDIDA * ALMACEN3.PRECIO_COSTO) * (ALMACEN3.GANANCIA + 1) * ALMACEN3.IMPUESTO2) AS T_IMP2, SUM((PED_CLIEN_DETALLE.CANTIDAD_PEDIDA * ALMACEN3.PRECIO_COSTO) * (ALMACEN3.GANANCIA + 1) * ALMACEN3.P_RETENCION) AS T_RET FROM PED_CLIEN_DETALLE INNER JOIN ALMACEN3 ON PED_CLIEN_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE PED_CLIEN_DETALLE.NO_PEDIDO = " & TxtNoPed.Text & " GROUP BY PED_CLIEN_DETALLE.NO_PEDIDO"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "UPDATE VENTAS SET SUBTOTAL = " & Format(tRs.Fields("T_VENTA"), "0.00") & ", IVA = " & Format(tRs.Fields("T_IVA"), "0.00") & ", IMPUESTO1 = " & Format(tRs.Fields("T_IMP1"), "0.00") & ", IMPUESTO2 = " & Format(tRs.Fields("T_IMP2"), "0.00") & ", RETENCION = " & Format(tRs.Fields("T_RET"), "0.00") & ", TOTAL = " & Format(CDbl(tRs.Fields("T_VENTA")) + CDbl(tRs.Fields("T_IVA")) + CDbl(tRs.Fields("T_IMP1")) + CDbl(tRs.Fields("T_IMP2")) - CDbl(tRs.Fields("T_RET")), "0.00") & " WHERE ID_VENTA = " & ClvVenta
                cnn.Execute (sBuscar)
            End If
            sBuscar = "SELECT ID_VENTA, ROUND(SUM(IMPORTE), 2) AS SUB, ROUND(SUM(IMPORTE + IVA + IMPUESTO1 + IMPUESTO2 - RETENCION), 2) AS TOT, ROUND(SUM(IVA), 2) AS IVA, ROUND(SUM(IMPUESTO1), 2) AS IMP1, ROUND(SUM(IMPUESTO2), 2) AS IMP2, ROUND(SUM(RETENCION), 2) AS RET FROM VENTAS_DETALLE GROUP BY ID_VENTA HAVING (ROUND(SUM(IMPORTE + IVA + IMPUESTO1 + IMPUESTO2 - RETENCION), 2) <> ROUND ((SELECT TOTAL FROM VENTAS WHERE (ID_VENTA = dbo.VENTAS_DETALLE.ID_VENTA)), 2))"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    sBuscar = "UPDATE VENTAS SET SUBTOTAL = " & Format(tRs.Fields("SUB"), "0.00") & ", TOTAL = " & Format(tRs.Fields("TOT"), "0.00") & ", IVA = " & Format(tRs.Fields("IVA"), "0.00") & ", IMPUESTO1 = " & Format(tRs.Fields("IMP1"), "0.00") & ", IMPUESTO2 = " & Format(tRs.Fields("IMP2"), "0.00") & ", RETENCION = " & Format(tRs.Fields("RET"), "0.00") & " WHERE (ID_VENTA= " & tRs.Fields("ID_VENTA") & ")"
                    cnn.Execute (sBuscar)
                    tRs.MoveNext
                Loop
            End If
            'IMPRIME TICKER DE VENTA...
            ImpTicket
            sBuscar = "INSERT INTO PEDIDO_VENTA (ID_VENTA, NO_PEDIDO) VALUES (" & ClvVenta & ", " & TxtNoPed.Text & ");"
            cnn.Execute (sBuscar)
            sBuscar = "UPDATE PED_CLIEN SET ESTADO = 'T', FECHA_FACTURACION = DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) WHERE NO_PEDIDO = '" & TxtNoPed.Text & "'"
            cnn.Execute (sBuscar)
        End If
        TxtNoPed.Text = ""
        Lvw2.ListItems.Clear
    Else
        MsgBox "EL CLIENTE NO CUENTA CON CREDITO SUFICIENTE PARA CERRAR LA VENTA", vbInformation, "SACC"
    End If
    Actualizar
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ELLI_Click()
On Error GoTo ManejaError
    Dim sEliminar As String
    sEliminar = "DELETE FROM PED_CLIEN WHERE NO_PEDIDO = " & guia
    cnn.Execute (sEliminar)
    Me.ELLI.Enabled = False
    Lvw1.ListItems.Clear
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualizar()
On Error GoTo ManejaError
    CmdCerrar.Enabled = False
    Command1.Enabled = False
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT NO_PEDIDO, USUARIO, NOMBRE, FECHA, NO_ORDEN, P.ID_CLIENTE FROM PED_CLIEN AS P JOIN CLIENTE AS C ON C.ID_CLIENTE = P.ID_CLIENTE WHERE P.ESTADO = 'C' ORDER BY NO_PEDIDO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Lvw1.ListItems.Clear
            Do While Not .EOF
                Set tLi = Lvw1.ListItems.Add(, , .Fields("NO_PEDIDO"))
                If Not IsNull(.Fields("USUARIO")) Then tLi.SubItems(1) = .Fields("USUARIO")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE")
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA")
                If Not IsNull(.Fields("NO_ORDEN")) Then tLi.SubItems(4) = .Fields("NO_ORDEN")
                If Not IsNull(.Fields("USUARIO")) Then
                    sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Val(.Fields("USUARIO"))
                    If Val(.Fields("USUARIO")) >= 0 Then
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            tLi.SubItems(5) = tRs2.Fields("NOMBRE") & " " & tRs2.Fields("APELLIDOS")
                        Else
                            tLi.SubItems(5) = .Fields("USUARIO")
                        End If
                    Else
                        tLi.SubItems(5) = .Fields("USUARIO")
                    End If
                End If
                If Not IsNull(.Fields("ID_CLIENTE")) Then tLi.SubItems(6) = .Fields("ID_CLIENTE")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim DesClente As String
    Dim IdDescuento As String
    Dim IdDescuento2 As String
    sRastreaNoPed = Item
    CapturoVenta = Item.SubItems(1)
    ClienteVenta = Item.SubItems(2)
    sBuscar = "SELECT DESCUENTO, ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & Item.SubItems(6)
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("DESCUENTO")) Then
            DesClente = tRs.Fields("DESCUENTO")
        Else
            DesClente = 0
        End If
        If Not IsNull(tRs.Fields("ID_DESCUENTO")) Then
            IdDescuento = tRs.Fields("ID_DESCUENTO")
        Else
            IdDescuento = 0
        End If
    End If
    sBuscar = "SELECT * FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & CDbl(Item)
    Set tRs = cnn.Execute(sBuscar)
    If (tRs.BOF And tRs.EOF) Then
        Lvw2.ListItems.Clear
        TxtNoPed.Text = ""
        MsgBox "PEDIDO VACIO!", vbInformation, "SACC"
        guia = Item
        Me.ELLI.Enabled = True
        elind = Item.Index
    Else
        Lvw2.ListItems.Clear
        TxtNoPed.Text = Item
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = Lvw2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
            tLi.SubItems(1) = tRs.Fields("CANTIDAD_PEDIDA") & ""
            tLi.SubItems(2) = tRs.Fields("CANTIDAD_EXISTENCIA") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD_PENDIENTE") & ""
            '*********************************** AGREGAR PRECIO DE VENTA *************************************
            ' POR :   H VALDEZ
            ' FECHA:  20 DE MAYO DE 2011
            '*************************************************************************************************
            sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & Item.SubItems(6)
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.BOF And tRs2.EOF) Then
                tLi.SubItems(4) = tRs2.Fields("PRECIO_VENTA")
            Else
                sBuscar = "SELECT PRECIO_COSTO, GANANCIA, CLASIFICACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    tLi.SubItems(4) = Format((CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))), "0.00")
                Else
                    tLi.SubItems(4) = "0.00"
                End If
                If (DesClente <> "" And CDbl(DesClente) > 0) Then
                    tLi.SubItems(4) = CDbl(Replace(tLi.SubItems(3), ",", "")) * (CDbl(Replace(DesClente, ",", "")) / 100)
                    If (IdDescuento <> "") Then
                        sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & IdDescuento & "' AND CLASIFICACION = '" & tRs1.Fields("CLASIFICACION") & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        IdDescuento2 = tRs2.Fields("PORCENTAJE")
                        tLi.SubItems(4) = (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) - (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) * CDbl(IdDescuento2) / 100
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub TxtBusProd_GotFocus()
    Me.TxtBusProd.BackColor = &HFFE1E1
    TxtBusProd.SelStart = 0
    TxtBusProd.SelLength = Len(TxtBusProd.Text)
End Sub
Private Sub TxtBusProd_LostFocus()
    TxtBusProd.BackColor = &H80000005
End Sub
Private Sub TxtCantidad_LostFocus()
    TxtCantidad.BackColor = &H80000005
End Sub
Private Sub TxtNoPed_Change()
    If TxtNoPed.Text = "" Then
        CmdCerrar.Enabled = False
        Command1.Enabled = False
    Else
        CmdCerrar.Enabled = True
        Command1.Enabled = True
    End If
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++ < Detalle Pedido > +++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub AgreLis()
On Error GoTo ManejaError
    Dim LI As ListItem
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    Dim Exp As Integer
    Dim TotVenta As Double
    Dim sBuscar As String
    Dim tRs2 As ADODB.Recordset
    Exp = 0
    NumeroRegistros = ListView2.ListItems.Count
    sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & TxtClvProd.Text & "' AND FECHA_FIN >= '" & DTPicker1.Value & "' AND ID_CLIENTE = " & SELE
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        TotVenta = CDbl(tRs2.Fields("PRECIO_VENTA"))
    Else
        sBuscar = "SELECT PRECIO_COSTO, GANANCIA, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & TxtClvProd.Text & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            TotVenta = CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA")))
            TotVenta = TotVenta * ((100 - Val(Replace(txtDescuento, ",", ""))) / 100)
        Else
            TotVenta = "0.00"
        End If
    End If
    tRs2.Close
    For Conta = 1 To NumeroRegistros
        If ListView2.ListItems.Item(Conta) = TxtClvProd.Text Then
            ListView2.ListItems.Item(Conta).SubItems(1) = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) + CDbl(TxtCantidad.Text), "0.00")
            If (Date + 10) >= DTPicker1.Value Then 'V.P. Urgente, Toma toda la existencia
                If CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) <= CDbl(ListView2.ListItems.Item(Conta).SubItems(2)) Then
                    ListView2.ListItems.Item(Conta).SubItems(3) = "0.00"
                Else
                    ListView2.ListItems.Item(Conta).SubItems(3) = CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) - CDbl(ListView2.ListItems.Item(Conta).SubItems(2))
                End If
            Else 'V.P. Normal, Toma la existencia dejando siempre el minimo
                If CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) <= (CDbl(ListView2.ListItems.Item(Conta).SubItems(2)) - CDbl(ListView2.ListItems.Item(Conta).SubItems(6))) Then
                    ListView2.ListItems.Item(Conta).SubItems(3) = "0.00"
                Else
                    ListView2.ListItems.Item(Conta).SubItems(3) = CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) - (CDbl(ListView2.ListItems.Item(Conta).SubItems(2)) - CDbl(ListView2.ListItems.Item(Conta).SubItems(6)))
                End If
            End If
            ListView2.ListItems.Item(Conta).SubItems(5) = txtTipo.Text
            ListView2.ListItems.Item(Conta).SubItems(6) = txtMin.Text
            ListView2.ListItems.Item(Conta).SubItems(8) = TotVenta
            Exp = 1
        End If
    Next Conta
    If Exp = 0 Then
        Set LI = ListView2.ListItems.Add(, , TxtClvProd.Text & "")
            LI.SubItems(1) = TxtCantidad.Text & ""
            LI.SubItems(2) = TxtCanExis.Text & ""
        If CDbl(TxtCantidad.Text) - CDbl(TxtCanExis.Text) <= 0 Then
            LI.SubItems(3) = "0.00"
        Else
            LI.SubItems(3) = CDbl(TxtCantidad.Text) - CDbl(TxtCanExis.Text)
        End If
        LI.SubItems(4) = txtDes.Text
        LI.SubItems(5) = txtTipo.Text
        LI.SubItems(6) = txtMin.Text
        LI.SubItems(7) = txtMarca.Text
        LI.SubItems(8) = TotVenta
    End If
    TxtClvProd.Text = ""
    TxtCanExis.Text = ""
    TxtCantidad.Text = ""
    txtDes.Text = ""
    txtTipo.Text = ""
    txtMarca.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub CmdAceptar_Click()
    AgreLis
    LvwProd.SetFocus
End Sub
Private Sub CmdQuitar_Click()
    If ind <> 0 Then
        ListView2.ListItems.Remove (ind)
        ind = 0
        Me.CmdQuitar.Enabled = False
        ListView2.SetFocus
    End If
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim PreFinVen As Double
    
    If Me.Option1.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, TIPO, C_MINIMA, MARCA FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & TxtBusProd.Text & "%' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, TIPO, C_MINIMA, MARCA FROM ALMACEN3 WHERE Descripcion LIKE '%" & TxtBusProd.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se ha encontrado el producto"
        Else
            LvwProd.ListItems.Clear
            '.MoveFirst
            Do While Not .EOF
                Set tLi = LvwProd.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")) & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("C_MINIMA") & ""
                tLi.SubItems(4) = .Fields("MARCA") & ""
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub LvwProd_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs2 As ADODB.Recordset
    TxtClvProd.Text = Item
    txtDes.Text = Item.SubItems(1)
    txtTipo.Text = Item.SubItems(2)
    txtMin.Text = Item.SubItems(3)
    txtMarca.Text = Item.SubItems(4)
    sBuscar = "SELECT CANTIDAD, SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & TxtClvProd.Text & "' AND SUCURSAL = 'BODEGA'"
    Set tRs2 = cnn.Execute(sBuscar)
    With tRs2
        If (.BOF And .EOF) Then
            TxtCanExis.Text = "0.00"
        Else
            .MoveFirst
            If Not IsNull(.Fields("CANTIDAD")) Then
                TxtCanExis.Text = .Fields("CANTIDAD")
            Else
                TxtCanExis.Text = "0.00"
            End If
            .MoveNext
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub LvwProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCantidad.SetFocus
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.CmdQuitar.Enabled = True
    ind = Item.Index
    Me.CmdQuitar.SetFocus
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.CmdQuitar.SetFocus
    End If
End Sub
Private Sub TxtBusProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtBusProd.Text <> "" Then
        Buscar
        LvwProd.SetFocus
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
Private Sub TxtClvProd_Change()
    If TxtCantidad.Text = "" Or TxtClvProd.Text = "" Then
        Me.CmdAceptar.Enabled = False
    Else
        Me.CmdAceptar.Enabled = True
    End If
End Sub
Private Sub TxtCantidad_Change()
    If TxtCantidad.Text = "" Or TxtClvProd.Text = "" Then
        Me.CmdAceptar.Enabled = False
    Else
        Me.CmdAceptar.Enabled = True
    End If
End Sub
Private Sub txtCantidad_GotFocus()
    Me.TxtCantidad.BackColor = &HFFE1E1
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCantidad.Text <> "" And TxtClvProd.Text <> "" Then
        AgreLis
        TxtBusProd.SetFocus
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
Private Sub TxtNoPed_GotFocus()
    Me.TxtNoPed.BackColor = &HFFE1E1
End Sub
Private Sub TxtNoPed_LostFocus()
    TxtNoPed.BackColor = &H80000005
End Sub
Private Sub Imprimir_Recibo()
On Error GoTo CancelaError
    If Text4.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim POSY As Integer
        POSY = 2800
        Dim NomClien As String
        Dim NoPed As String
        Dim fecha As String
        Dim NoOrden As String
        sBuscar = "SELECT * FROM VsVentProg WHERE NO_PEDIDO = " & Text4.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            CommonDialog1.Flags = 64
            CommonDialog1.CancelError = True
            CommonDialog1.ShowPrinter
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
            Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print "     Cliente : " & tRs.Fields("NOMBRE")
            Printer.Print "     No. de Venta Programada : " & tRs.Fields("NO_PEDIDO") & "                                No. Orden : " & tRs.Fields("NO_ORDEN")
            Printer.Print "     VENTA PROGRAMADA"
            Printer.Print "     Fecha de Entrega : " & tRs.Fields("FECHA")
            Printer.Print "     No. de Orden : " & Text6.Text
            Printer.Print "     Fecha : " & Now
            NomClien = tRs.Fields("NOMBRE")
            NoPed = tRs.Fields("NO_PEDIDO")
            fecha = tRs.Fields("FECHA")
            NoOrden = tRs.Fields("NO_ORDEN")
            Printer.Print ""
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print ""
            sBuscar = "SELECT * FROM VsVentProgDet WHERE NO_PEDIDO = " & Text4.Text
            Set tRs = cnn.Execute(sBuscar)
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print "Descripcion"
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print "C. PEDIDA"
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print "C. PENDIENTE"
            POSY = POSY + 400
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print Mid(tRs.Fields("ID_PRODUCTO"), 1, 25)
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2200
                    Printer.Print tRs.Fields("Descripcion")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 8800
                    Printer.Print tRs.Fields("CANTIDAD_PEDIDA")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 10000
                    Printer.Print tRs.Fields("CANTIDAD_PENDIENTE")
                    tRs.MoveNext
                    POSY = POSY + 200
                    If POSY >= 14200 Then
                        Printer.NewPage
                        POSY = 2800
                        Printer.Print ""
                        Printer.Print ""
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                        Printer.Print VarMen.Text5(0).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        Printer.Print ""
                        Printer.Print "     Cliente : " & NomClien
                        Printer.Print "     No. Pedido : " & NoPed & "                                No. Orden : " & NoOrden
                        Printer.Print "     VENTA PROGRAMADA"
                        Printer.Print "     Fecha de Entrega : " & fecha
                        Printer.Print "     No. de Orden : " & Text6.Text
                        Printer.Print ""
                        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        Printer.Print ""
                        Printer.Print ""
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 100
                        Printer.Print "Producto"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 2200
                        Printer.Print "Descripcion"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 8800
                        Printer.Print "C. PEDIDA"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 10000
                        Printer.Print "C. PENDIENTE"
                        POSY = POSY + 400
                    End If
                Loop
            End If
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "FIN DEL LISTADO"
            Printer.EndDoc
        Else
            MsgBox "NO EXISTE UNA VENTA CON ESE FOLIO!", vbInformation, "SACC"
        End If
    Else
        MsgBox "DEBE DAR EL NUMERO DE VENTA A IMPRIMIR!", vbInformation, "SACC"
    End If
    Exit Sub
CancelaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ImpTicket()
    Dim sBuscar As String
    Dim Acum As String
    Dim tRs As ADODB.Recordset
    Dim POSY As Integer
    Dim Usuario As String
    Dim Cliente As String
    Dim sExibicion As String
    Dim sTipoVenta As String
    Dim IVA As String
    Dim IMPUESTO1 As String
    Dim IMPUESTO2 As String
    Dim RETENCION As String
    sBuscar = "SELECT ID_USUARIO, NOMBRE, UNA_EXIBICION, TIPO_PAGO, IVA, IMPUESTO1, IMPUESTO2, RETENCION FROM VENTAS WHERE ID_VENTA = " & ClvVenta
    Set tRs = cnn.Execute(sBuscar)
    Usuario = tRs.Fields("ID_USUARIO")
    Cliente = tRs.Fields("NOMBRE")
    sExibicion = tRs.Fields("UNA_EXIBICION")
    sTipoVenta = tRs.Fields("TIPO_PAGO")
    IVA = tRs.Fields("IVA")
    IMPUESTO1 = tRs.Fields("IMPUESTO1")
    IMPUESTO2 = tRs.Fields("IMPUESTO2")
    RETENCION = tRs.Fields("RETENCION")
    tRs.Close
    sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Usuario
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        Usuario = VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Else
        Usuario = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
    End If
    tRs.Close
    Acum = "0"
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
    Printer.Print "No. DE VENTA : " & ClvVenta
    If sTipoVenta = "C" Then
        Printer.Print "FORMA DE PAGO : EFECTIVO"
    Else
        If sTipoVenta = "H" Then
            Printer.Print "FORMA DE PAGO : CHEQUE"
        Else
            If sTipoVenta = "T" Then
                Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
            Else
                If sTipoVenta = "E" Then
                    Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                Else
                    Printer.Print "FORMA DE PAGO : NO INDICADO"
                End If
            End If
        End If
    End If
    Printer.Print "ATENDIDO POR : " & Usuario
    Printer.Print "CLIENTE : " & Cliente
    If sExibicion = "N" Then
        Printer.Print "VENTA A CREDITO"
    Else
        Printer.Print "VENTA A CONTADO"
    End If
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                          NOTA DE FACTURA"
    Printer.Print "--------------------------------------------------------------------------------"
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
    sBuscar = "SELECT ID_PRODUCTO, PRECIO_VENTA, CANTIDAD FROM VENTAS_DETALLE WHERE ID_VENTA = " & ClvVenta
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print tRs.Fields("CANTIDAD")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Format(CDbl(tRs.Fields("PRECIO_VENTA")), "###,###,##0.00")
            Acum = CDbl(Acum) + (CDbl(tRs.Fields("PRECIO_VENTA") * CDbl(tRs.Fields("CANTIDAD"))))
            tRs.MoveNext
        Loop
    End If
    Printer.Print ""
    sBuscar = "SELECT SUM(CANTIDAD * PRECIO_VENTA) AS SUBTOT, SUM(IVA) AS IVA, SUM(IMPUESTO1) AS IMP1, SUM(IMPUESTO2) AS IMP2, SUM(RETENCION) AS RET FROM VENTAS_DETALLE WHERE ID_VENTA = " & ClvVenta
    Set tRs = cnn.Execute(sBuscar)
    Printer.Print "SUBTOTAL : " & Format(CDbl(tRs.Fields("SUBTOT")), "###,###,###,##0.00")
    Printer.Print "IVA              : " & Format(CDbl(tRs.Fields("IVA")), "###,###,###,##0.00")
    If CDbl(IMPUESTO1) > 0 Then
        Printer.Print "IMPUESTO 1       : " & Format(CDbl(tRs.Fields("IMP1")), "###,###,###,##0.00")
    End If
    If CDbl(IMPUESTO2) > 0 Then
        Printer.Print "IMPUESTO 2       : " & Format(CDbl(tRs.Fields("IMP2")), "###,###,###,##0.00")
    End If
    If CDbl(RETENCION) > 0 Then
        Printer.Print "RETENCIÓN        : " & Format(CDbl(tRs.Fields("RET")), "###,###,###,##0.00")
    End If
    Printer.Print "TOTAL        : " & Format(CDbl(tRs.Fields("SUBTOT")) + CDbl(tRs.Fields("IVA")) + CDbl(tRs.Fields("IMP1")) + CDbl(tRs.Fields("IMP2")) - CDbl(tRs.Fields("RET")), "###,###,###,##0.00")
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
    Printer.Print "     DESPUES DE HABER EFECTUADO SU "
    Printer.Print "                                COMPRA"
    Printer.Print "                APLICA RESTRICCIONES"
    Printer.Print "--------------------------------------------------------------------------------"
    'sBuscar = "UPDATE VENTAS SET SUBTOTAL = " & Acum & ", IVA = " & Format(CDbl(Acum) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00") & ", TOTAL = " & Format(CDbl(Acum) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1), "0.00") & " WHERE ID_VENTA = " & ClvVenta
    'cnn.Execute (sBuscar)
    Printer.EndDoc
End Sub
Function Imprimir_Ticket(cNoCom As Integer)
On Error GoTo ManejaError
    Printer.Print "        " & VarMen.Text5(0).Text
    Printer.Print "           ORDEN DE PRODUCCIÓN"
    Printer.Print "FECHA : " & Now
    Printer.Print "No. DE ORDEN DE PRODUCCCION : " & cNoCom
    Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "COMENTARIO : V.Prog. No. " & IdPedido & " No.Orden " & Text6.Text & " Cliente: " & Text7.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TINTA"
    Dim NRegistros As Integer
    Dim Con As Integer
    Dim POSY As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & cNoCom
    Set tRs = cnn.Execute(sBuscar)
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    Do While Not tRs.EOF
        If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) = "I" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print tRs.Fields("CANTIDAD")
        End If
        tRs.MoveNext
    Loop
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TONER"
    POSY = POSY + 600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    tRs.MoveFirst
    Do While Not tRs.EOF
        If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) = "T" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print tRs.Fields("CANTIDAD")
        End If
        tRs.MoveNext
    Loop
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.EndDoc
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
