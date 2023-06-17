VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Creditos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Creditos"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UUID"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   85
      Top             =   4800
      Width           =   735
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   82
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label23 
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
         TabIndex        =   83
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Creditos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Creditos.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   855
      Left            =   600
      TabIndex        =   74
      Top             =   9000
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1508
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   5880
      TabIndex        =   73
      Top             =   8400
      Width           =   3015
   End
   Begin VB.TextBox Text21 
      Height          =   495
      Left            =   600
      TabIndex        =   72
      Top             =   8400
      Width           =   5055
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   255
      Left            =   360
      TabIndex        =   71
      Top             =   7800
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text19 
      Height          =   195
      Left            =   1665
      TabIndex        =   68
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   58
      Top             =   6360
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Creditos.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "Creditos.frx":2156
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
         TabIndex        =   59
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Busca"
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
      Left            =   2160
      Picture         =   "Creditos.frx":4238
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5160
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Factuas"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   4800
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Notas de Venta"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   54
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   53
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
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
      Left            =   2160
      Picture         =   "Creditos.frx":6C0A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Abonar"
      TabPicture(0)   =   "Creditos.frx":95DC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTitulo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNoMov"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView2(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView2(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "SSTab2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSaldo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtIdCuenta"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "texidventa"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text22"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command5"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Historial"
      TabPicture(1)   =   "Creditos.frx":95F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text11"
      Tab(1).Control(1)=   "Text8"
      Tab(1).Control(2)=   "CmdImprimir"
      Tab(1).Control(3)=   "CmdVer"
      Tab(1).Control(4)=   "LVDeuda"
      Tab(1).Control(5)=   "DTFechaAl"
      Tab(1).Control(6)=   "DTFechade"
      Tab(1).Control(7)=   "CommonDialog1"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "De"
      Tab(1).Control(10)=   "Label1"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton Command5 
         Caption         =   "Incobrable"
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
         Left            =   4560
         Picture         =   "Creditos.frx":9614
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   6945
         TabIndex        =   75
         Text            =   "Text22"
         Top             =   4080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   120
         TabIndex        =   70
         Top             =   6960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox texidventa 
         Height          =   285
         Left            =   6720
         TabIndex        =   69
         Top             =   4080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74160
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1080
         Width           =   6135
      End
      Begin VB.TextBox txtIdCuenta 
         Height          =   285
         Left            =   6480
         TabIndex        =   45
         Top             =   4080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSaldo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   43
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Movimiento"
         Height          =   735
         Left            =   3840
         TabIndex        =   29
         Top             =   1200
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Notas de Venta"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Factuas"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Registrar"
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
         Left            =   5880
         Picture         =   "Creditos.frx":BFE6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6960
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2445
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4313
         _Version        =   393216
         TabOrientation  =   2
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   2
         TabHeight       =   723
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Cheque"
         TabPicture(0)   =   "Creditos.frx":E9B8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(1)=   "Label5"
         Tab(0).Control(2)=   "Label6"
         Tab(0).Control(3)=   "Label8"
         Tab(0).Control(4)=   "Label12"
         Tab(0).Control(5)=   "Label17"
         Tab(0).Control(6)=   "DTPicker1"
         Tab(0).Control(7)=   "Text6"
         Tab(0).Control(8)=   "Combo1"
         Tab(0).Control(9)=   "Text4"
         Tab(0).Control(10)=   "Text15"
         Tab(0).Control(11)=   "Text16"
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Efectivo"
         TabPicture(1)   =   "Creditos.frx":E9D4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label9"
         Tab(1).Control(1)=   "Label20"
         Tab(1).Control(2)=   "Text5"
         Tab(1).Control(3)=   "DTPicker2"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Transferencia"
         TabPicture(2)   =   "Creditos.frx":E9F0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label10"
         Tab(2).Control(1)=   "Label11"
         Tab(2).Control(2)=   "Label13"
         Tab(2).Control(3)=   "Label18"
         Tab(2).Control(4)=   "Label19"
         Tab(2).Control(5)=   "Label22"
         Tab(2).Control(6)=   "Text9"
         Tab(2).Control(7)=   "Text10"
         Tab(2).Control(8)=   "Combo2"
         Tab(2).Control(9)=   "Text17"
         Tab(2).Control(10)=   "Text18"
         Tab(2).Control(11)=   "DTPicker4"
         Tab(2).ControlCount=   12
         TabCaption(3)   =   "Nota de Credito"
         TabPicture(3)   =   "Creditos.frx":EA0C
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label15"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label16"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label21"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "ListView3"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "Text12"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "Text13"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "DTPicker3"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).ControlCount=   7
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   320
            Left            =   -71760
            TabIndex        =   81
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17235969
            CurrentDate     =   39860
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   310
            Left            =   5520
            TabIndex        =   79
            Top             =   2030
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17235969
            CurrentDate     =   39860
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   -72000
            TabIndex        =   77
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   17235969
            CurrentDate     =   39860
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   -69360
            TabIndex        =   67
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -71760
            TabIndex        =   65
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   -69720
            MaxLength       =   35
            TabIndex        =   63
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   -72120
            TabIndex        =   61
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            MaxLength       =   18
            TabIndex        =   49
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   18
            TabIndex        =   48
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   -71760
            TabIndex        =   38
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   -71760
            MaxLength       =   18
            TabIndex        =   36
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   -71760
            MaxLength       =   18
            TabIndex        =   34
            Top             =   1200
            Width           =   3570
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   -72000
            MaxLength       =   18
            TabIndex        =   32
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   -72120
            MaxLength       =   18
            TabIndex        =   27
            Top             =   1200
            Width           =   3915
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Creditos.frx":EA28
            Left            =   -72120
            List            =   "Creditos.frx":EA2A
            TabIndex        =   21
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   -72120
            MaxLength       =   18
            TabIndex        =   20
            Top             =   480
            Width           =   3900
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   -72120
            TabIndex        =   22
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   17235969
            CurrentDate     =   38833
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   1575
            Left            =   1080
            TabIndex        =   52
            Top             =   360
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
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha :"
            Height          =   255
            Left            =   -73680
            TabIndex        =   80
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha :"
            Height          =   255
            Left            =   4920
            TabIndex        =   78
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Fecha :"
            Height          =   255
            Left            =   -73080
            TabIndex        =   76
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Referencia :"
            Height          =   255
            Left            =   -70320
            TabIndex        =   66
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Total del Deposito :"
            Height          =   255
            Left            =   -73680
            TabIndex        =   64
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Referencia :"
            Height          =   255
            Left            =   -70680
            TabIndex        =   62
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Total de Deposito :"
            Height          =   255
            Left            =   -73800
            TabIndex        =   60
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Folio Nota :"
            Height          =   255
            Left            =   960
            TabIndex        =   51
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Banco :"
            Height          =   255
            Left            =   -73680
            TabIndex        =   39
            Top             =   765
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Numero de Tranferencia:"
            Height          =   255
            Left            =   -73680
            TabIndex        =   37
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   -73680
            TabIndex        =   35
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   -73080
            TabIndex        =   33
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   -73800
            TabIndex        =   28
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha para Depositar :"
            Height          =   255
            Left            =   -73800
            TabIndex        =   25
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Banco :"
            Height          =   255
            Left            =   -73800
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Numero de Cheque :"
            Height          =   255
            Left            =   -73800
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   1005
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   6000
         Width           =   6975
      End
      Begin VB.CommandButton CmdImprimir 
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
         Height          =   375
         Left            =   -69000
         Picture         =   "Creditos.frx":EA2C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdVer 
         Caption         =   "Ver"
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
         Picture         =   "Creditos.frx":113FE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6960
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "Creditos.frx":13DD0
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   6135
      End
      Begin MSComctlLib.ListView LVDeuda 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   9
         Top             =   2520
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5106
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
      Begin MSComCtl2.DTPicker DTFechaAl 
         Height          =   375
         Left            =   -71280
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   38828
      End
      Begin MSComCtl2.DTPicker DTFechade 
         Height          =   375
         Left            =   -73800
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   38828
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -69240
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Visible         =   0   'False
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Saldo pendiente:"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   4125
         Width           =   1335
      End
      Begin VB.Label lblNoMov 
         Caption         =   "Factura No:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   4125
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Facturas Pendientes de Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label De 
         Caption         =   "De la Fecha :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "A la Fecha :"
         Height          =   255
         Left            =   -72240
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad Pendiente :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblTipoMov 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Numero de Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "Creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ClvCliente As Integer
Dim VarLimCred As Double
Dim ClvUsuario As String
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim NOM As String
Dim CLVCLIEN As Integer
Dim LimCred As String
Dim fact As String
Dim idven As Integer
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        DTPicker1.SetFocus
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
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim idcuentas As Integer
    Dim IdVenta As String
    If Option2(0).Value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, ID_CLIENTE, SUM(DEUDA) AS DEUDA, ID_CUENTA FROM VsCxC WHERE FOLIO = '" & Text14.Text & "' AND PAGADA = 'N' GROUP BY FOLIO, ID_CLIENTE, ID_CUENTA, ID_VENTA"
        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE FOLIO = '" & Text14.Text & "' "
    Else
        If Option2(1).Value Then
            sBuscar = "SELECT ID_VENTA, ID_CLIENTE, DEUDA, ID_CUENTA FROM VsCxC WHERE ID_VENTA = " & Text14.Text & " AND PAGADA = 'N' "
            sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE FOLIO = '" & Text14.Text & "'"
        Else
            sBuscar = "SELECT ID_VENTA, FOLIO, ID_CLIENTE, SUM(DEUDA) AS DEUDA, ID_CUENTA FROM VsCxC WHERE UUID LIKE '%" & Text14.Text & "%' AND PAGADA = 'N' GROUP BY FOLIO, ID_CLIENTE, ID_CUENTA, ID_VENTA"
            sBusca = "SELECT SUM(ABONOS_CUENTA.CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA INNER JOIN VENTAS ON ABONOS_CUENTA.ID_VENTA = VENTAS.ID_VENTA WHERE VENTAS.UUID LIKE '%" & Text14.Text & "%'"
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    Set tRs1 = cnn.Execute(sBusca)
    If Not (tRs.EOF And tRs.BOF) Then
        ClvCliente = tRs.Fields("ID_CLIENTE")
        Text19.Text = tRs.Fields("ID_CUENTA")
        Do While Not tRs.EOF
            IdVenta = IdVenta & tRs.Fields("ID_VENTA") & ", "
            tRs.MoveNext
        Loop
        IdVenta = Mid(IdVenta, 1, Len(IdVenta) - 2)
        tRs.MoveFirst
        ActualVenta (IdVenta)
        If Option2(0).Value Then
            Option1(0).Value = True
            txtNo.Text = tRs.Fields("FOLIO")
        Else
            If Option2(1).Value Then
                Option1(1).Value = True
                txtNo.Text = tRs.Fields("ID_VENTA")
            Else
                Option2(2).Value = True
                txtNo.Text = tRs.Fields("FOLIO")
            End If
        End If
        txtIdCuenta.Text = tRs.Fields("ID_CUENTA")
        sBuscar = "SELECT NOMBRE, LIMITE_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            If Not IsNull(tRs1.Fields("NOMBRE")) Then Text1.Text = tRs1.Fields("NOMBRE")
            If Not IsNull(tRs1.Fields("NOMBRE")) Then Text3.Text = tRs1.Fields("NOMBRE")
            If Not IsNull(tRs1.Fields("NOMBRE")) Then Text11.Text = tRs1.Fields("NOMBRE")
            LimCred = tRs1.Fields("LIMITE_CREDITO")
        Else
            MsgBox "EL CLIENTE FUE ELIMINADO! DEBE REASIGNAR TODOS LOS MOVIMIENTOS A UN CLIENTE NUEVO!", vbExclamation, "SACC"
        End If
    Else
        If Option2(0).Value Then
            sBuscar = "SELECT PAGADA FROM VsCxC WHERE FOLIO = '" & Text14.Text & "' GROUP BY FOLIO, ID_CLIENTE, ID_CUENTA, PAGADA"
        Else
            If Option2(1).Value Then
                sBuscar = "SELECT PAGADA FROM VsCxC WHERE ID_VENTA = " & Text14.Text
            Else
                sBuscar = "SELECT PAGADA FROM VsCxC WHERE UUID LIKE '%" & Text14.Text & "%'"
            End If
        End If
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            If tRs1.Fields("PAGADA") = "S" Then
                MsgBox "EL MOVIMIENTO YA FUE PAGADO!", vbExclamation, "SACC"
            Else
                MsgBox "EL MOVIMIENTO TIENE UN PROBLEMA, FAVOR DE NOTIFICAR AL ADMINISTRADOR DEL SISTEMA!", vbExclamation, "SACC"
            End If
        Else
            MsgBox "EL MOVIMIENTO NO ESTA CAPTURADO!", vbExclamation, "SACC"
        End If
    End If
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim continuar As Boolean
    Dim completo As Boolean
    Dim pagada As String
    Dim deuda As Double
    Dim pago As Double
    Dim efe As String
    Dim NoCheq As String
    Dim banco As String
    Dim fechacheq As String
    Dim vFecha As String
    Dim CUENTAS As Integer
    vFecha = Format(Date, "dd/mm/yyyy")
    completo = True
    If SSTab2.Tab = 0 Then
        If Text4.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or txtSaldo = "" Then
            MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
            completo = False
        Else
            Text4.Text = Replace(Text4.Text, ",", "")
            pago = CDbl(Text4.Text)
            efe = "C"
            NoCheq = Text6.Text
            banco = Combo1.Text
            fechacheq = DTPicker1.Value
            vFecha = DTPicker1.Value
        End If
    ElseIf SSTab2.Tab = 1 Then
        If Text5.Text = "" Or txtSaldo = "" Then
            MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
            completo = False
        Else
            Text5.Text = Replace(Text5.Text, ",", "")
            pago = CDbl(Text5.Text)
            vFecha = DTPicker2.Value
            efe = "E"
            NoCheq = ""
            banco = ""
            fechacheq = ""
        End If
    ElseIf SSTab2.Tab = 2 Then
        If Text10.Text = "" Or Text9.Text = "" Or Combo2.Text = "" Or txtSaldo = "" Then
            MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
            completo = False
        Else
            Text9.Text = Replace(Text9.Text, ",", "")
            pago = CDbl(Text9.Text)
            efe = "T"
            NoCheq = Text10.Text
            banco = Combo2.Text
            vFecha = DTPicker4.Value
            fechacheq = ""
        End If
    Else
        If Text12.Text = "" Or txtSaldo = "" Then
            MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
            completo = False
        Else
            Text12.Text = Replace(Text12.Text, ",", "")
            pago = CDbl(Text12.Text)
            efe = "N"
            vFecha = DTPicker3.Value
            NoCheq = Text13.Text
            banco = ""
            fechacheq = ""
        End If
    End If
    deuda = CDbl(txtSaldo.Text) - pago
    Text2.Text = CDbl(Text2.Text) - pago
    Text2.Text = Replace(Text2.Text, ",", "")
    If txtSaldo.Text <> "" Then
        If CDbl(pago) + 0.1 >= CDbl(txtSaldo.Text) Then
            pagada = "S"
            deuda = 0
            Text2.Text = CDbl(Text2.Text) - CDbl(txtSaldo.Text)
        Else
            pagada = "N"
            deuda = CDbl(txtSaldo.Text) - pago
            Text2.Text = CDbl(Text2.Text) - pago
        End If
        If CDbl(deuda <= 0.1) Then 'modificacion
            pagada = "S"
            deuda = 0
        End If
        continuar = True
        If CDbl(pago) > CDbl(txtSaldo.Text) Then
            If MsgBox("Si continua se creara una Nota de Credito por el valor del exedente de pago " & Chr(13) & _
               "                               Desea Continuar?", vbYesNo, "SACC") = vbYes Then
                sBuscar = "INSERT INTO NOTA_CREDITO (IMPORTE, NOMBRE, TOTAL, FECHA, MOTIVOCAMBIO, ID_VENTA, ID_USUARIO, ID_CLIENTE) VALUES (" & pago - CDbl(txtSaldo.Text) & ", '" & Text1.Text & "', " & pago - CDbl(txtSaldo.Text) & ", '" & Format(vFecha, "dd/mm/yyyy") & "', 'SALDO A FAVOR AL ABONAR VENTA/FACTURA #" & txtNo.Text & "', 0, '" & VarMen.Text1(0).Text & "', " & ClvCliente & ");"
                cnn.Execute (sBuscar)
            Else
                continuar = False
            End If
        Else
            continuar = True
        End If
    Else
        continuar = False
    End If
    If continuar Then
        If Text15.Text = "" Then
            If Text17.Text <> "" Then
                Text15.Text = Text17.Text
            Else
                Text15.Text = 0
            End If
        End If
        If Text16.Text = "" Then
            If Text18.Text <> "" Then
                Text16.Text = Text18.Text
            Else
                Text16.Text = 0
            End If
        End If
        If Option1(1).Value Then
            sBuscar = "INSERT INTO ABONOS_CUENTA (ID_CLIENTE, CANT_ABONO, FECHA, ID_USUARIO, NO_CHEQUE, BANCO, FECHA_CHEQUE, EFECTIVO, ID_CUENTA, REFERENCIA, DEPOSITO_TOTAL,ID_VENTA,FOLIO) VALUES(" & ClvCliente & ", " & pago & ", '" & vFecha & "', '" & VarMen.Text1(0).Text & "', '" & NoCheq & "' , '" & banco & "' , '" & fechacheq & "' , '" & efe & "' , " & txtIdCuenta.Text & ", '" & Text16.Text & "', " & Text15.Text & ",'" & Text22.Text & "','V');"
            cnn.Execute (sBuscar)
        End If
        If Option1(0).Value Then
            sBuscar = "INSERT INTO ABONOS_CUENTA (ID_CLIENTE, CANT_ABONO, FECHA, ID_USUARIO, NO_CHEQUE, BANCO, FECHA_CHEQUE, EFECTIVO, ID_CUENTA, REFERENCIA, DEPOSITO_TOTAL,FOLIO,ID_VENTA) VALUES(" & ClvCliente & ", " & pago & ", '" & vFecha & "', '" & VarMen.Text1(0).Text & "', '" & NoCheq & "' , '" & banco & "' , '" & fechacheq & "' , '" & efe & "' , " & txtIdCuenta.Text & ", '" & Text16.Text & "', " & Text15.Text & ",'" & Text22.Text & "'," & texidventa & ");"
            cnn.Execute (sBuscar)
        End If
        If Option1(1).Value Then
            sBuscar = "SELECT * FROM VsCxC WHERE ID_VENTA = '" & Text22.Text & "' AND PAGADA = 'N'  "
            Set tRs = cnn.Execute(sBuscar)
        End If
        If Option1(0).Value Then
            sBuscar = "SELECT * FROM VsCxC WHERE FOLIO = '" & Text22.Text & "' AND PAGADA = 'N'  "
            Set tRs = cnn.Execute(sBuscar)
        End If
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
            CUENTAS = tRs.Fields("ID_CUENTA")
            sBuscar = "UPDATE CUENTAS SET DEUDA = " & deuda & ", PAGADA = '" & pagada & "'  WHERE ID_CUENTA =' " & CUENTAS & "'"
            cnn.Execute (sBuscar)
              tRs.MoveNext
            Loop
        End If
        If efe = "N" Then
            sBuscar = "UPDATE NOTA_CREDITO SET APLICADA = 'S' WHERE ID_NOTA = " & Text13.Text
            cnn.Execute (sBuscar)
        End If
        Actual (ClvCliente)
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.Text5.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim ID As String
    Dim foli As String
    Dim ven As Double
    sBuscar = "SELECT * FROM TINTAS WHERE ID_PRODUCTO LIKE '%" & Text21.Text & "%'  "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ID = tRs.Fields("ID_PRODUCTO")
            ven = tRs.Fields("PRECIO")
            sBusca = "UPDATE ALMACEN3 SET PRECIO_COSTO =' " & ven & "' WHERE ID_PRODUCTO = ' " & ID & "' "
            cnn.Execute (sBusca)
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command5_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim CUENTAS As Integer
    Dim Cont As Integer
    For Cont = 1 To ListView2(0).ListItems.Count
        If ListView2(0).ListItems(Cont).Checked Then
            If Option1(1).Value Then
                sBuscar = "SELECT * FROM VsCxC WHERE ID_VENTA = '" & ListView2(0).ListItems(Cont).SubItems(6) & "' AND PAGADA = 'N'  "
                Set tRs = cnn.Execute(sBuscar)
            End If
            If Option1(0).Value Then
                sBuscar = "SELECT * FROM VsCxC WHERE FOLIO = '" & ListView2(0).ListItems(Cont).SubItems(6) & "' AND PAGADA = 'N'  "
                Set tRs = cnn.Execute(sBuscar)
            End If
            sBuscar = "UPDATE CUENTAS PAGADA = 'I' WHERE ID_CUENTA =' " & ListView2(0).ListItems(Cont).SubItems(5) & "'"
            cnn.Execute (sBuscar)
        End If
    Next Cont
    For Cont = 1 To ListView2(1).ListItems.Count
        If ListView2(0).ListItems(Cont).Checked Then
            If Option1(1).Value Then
                sBuscar = "SELECT * FROM VsCxC WHERE ID_VENTA = '" & ListView2(1).ListItems(Cont).SubItems(6) & "' AND PAGADA = 'N'  "
                Set tRs = cnn.Execute(sBuscar)
            End If
            If Option1(0).Value Then
                sBuscar = "SELECT * FROM VsCxC WHERE FOLIO = '" & ListView2(1).ListItems(Cont).SubItems(6) & "' AND PAGADA = 'N'  "
                Set tRs = cnn.Execute(sBuscar)
            End If
            sBuscar = "UPDATE CUENTAS PAGADA = 'I' WHERE ID_CUENTA =' " & ListView2(1).ListItems(Cont).SubItems(5) & "'"
            cnn.Execute (sBuscar)
        End If
    Next Cont
    Actual (ClvCliente)
End Sub
Private Sub DTFechade_Change()
    DTFechaAl.MinDate = DTFechade.Value
End Sub
Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    DTPicker3.Value = Format(Date, "dd/mm/yyyy")
    DTPicker4.Value = Format(Date, "dd/mm/yyyy")
    Combo1.AddItem "AFIRME"
    Combo1.AddItem "AZTECA"
    Combo1.AddItem "BANAMEX"
    Combo1.AddItem "BANCOMER"
    Combo1.AddItem "BANORTE"
    Combo1.AddItem "HSBC"
    Combo1.AddItem "SANTANDER"
    Combo1.AddItem "SCOTIABANK"
    Combo1.AddItem "Otros"
    Combo2.AddItem "AFIRME"
    Combo1.AddItem "AZTECA"
    Combo2.AddItem "BANAMEX"
    Combo2.AddItem "BANCOMER"
    Combo2.AddItem "BANORTE"
    Combo2.AddItem "HSBC"
    Combo2.AddItem "SANTANDER"
    Combo2.AddItem "SCOTIABANK"
    Combo2.AddItem "Otros"
    DTFechade.Value = Format(Date, "dd/mm/yyyy")
    DTFechade.Value = DTFechade.Value - 30
    DTFechaAl.MinDate = DTFechade.Value
    DTFechaAl.Value = Format(Date, "dd/mm/yyyy")
    Me.cmdVer.Enabled = False
    Me.CmdImprimir.Enabled = False
    ClvUsuario = VarMen.Text1(0).Text
    Command3.Enabled = False
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
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Cliente", 1800
        .ColumnHeaders.Add , , "Nombre", 7450
        .ColumnHeaders.Add , , "RFC", 2450
        .ColumnHeaders.Add , , "Limite de credito", 2450
        .ColumnHeaders.Add , , "Descuento", 1550
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "FOLIO", 1800
        .ColumnHeaders.Add , , "TOTAL", 1800
        .ColumnHeaders.Add , , "FECHA EXPEDICION", 7450
    End With
    With LVDeuda
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CONCEPTO", 1000
        .ColumnHeaders.Add , , "FECHA", 1500
        .ColumnHeaders.Add , , "IMPORTE", 2000
        .ColumnHeaders.Add , , "PENDIENTE", 2000
        .ColumnHeaders.Add , , "FACTURA", 2000
    End With
    With ListView2(0)
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Fecha Vence", 1500
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "Deuda", 2000
        .ColumnHeaders.Add , , "Id_Cuenta", 50
        .ColumnHeaders.Add , , "Id_venta", 50
    End With
    With ListView2(1)
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Nota", 1000
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Fecha Vence", 1500
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "Deuda", 2000
        .ColumnHeaders.Add , , "Id_Cuenta", 50
        .ColumnHeaders.Add , , "Id_venta", 50
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Fecha Vence", 1500
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "Deuda", 2000
        .ColumnHeaders.Add , , "Id_Cuenta", 50
        .ColumnHeaders.Add , , "Id_venta", 50
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Fecha Vence", 1500
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "Deuda", 2000
        .ColumnHeaders.Add , , "Id_Cuenta", 50
        .ColumnHeaders.Add , , "Id_venta", 50
    End With
    Me.Command3.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If IsNumeric(Text1.Text) Then
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0 OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0 OR ID_CLIENTE = '" & Text1.Text & "' AND LIMITE_CREDITO <> 0"
    Else
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0 OR NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
            MsgBox "No se encontro cliente con credito registrado a ese nombre"
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
                tLi.SubItems(2) = .Fields("RFC") & ""
                tLi.SubItems(3) = .Fields("LIMITE_CREDITO") & ""
                If .Fields("DESCUENTO") = "" Then
                    tLi.SubItems(4) = "0.00"
                Else
                    tLi.SubItems(4) = .Fields("DESCUENTO") & ""
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
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
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = LVDeuda.ColumnHeaders.Count
            For Con = 1 To LVDeuda.ColumnHeaders.Count
                StrCopi = StrCopi & LVDeuda.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To LVDeuda.ListItems.Count
                StrCopi = StrCopi & LVDeuda.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & LVDeuda.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    ClvCliente = Item
    Text3.Text = Item.SubItems(1)
    Text11.Text = Item.SubItems(1)
    LimCred = Item.SubItems(3)
    Actual (Item)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actual(Item As String)
    Dim Acum As Double
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sBusc As String
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim tLi2 As ListItem
    Command3.Enabled = False
    ListView2(0).ListItems.Clear
    ListView2(1).ListItems.Clear
    ListView3.ListItems.Clear
    LVDeuda.ListItems.Clear
    txtNo.Text = ""
    txtSaldo.Text = ""
    Text10.Text = ""
    Text9.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text4.Text = ""
    sBuscar = "SELECT ID_NOTA, TOTAL, FECHA FROM NOTA_CREDITO WHERE APLICADA = 'N' AND ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView3.ListItems.Add(, , .Fields("ID_NOTA"))
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(1) = .Fields("TOTAL")
                If Not IsNull(.Fields("Fecha")) Then tLi.SubItems(2) = .Fields("Fecha")
                .MoveNext
            Loop
        End If
    End With
    sBuscar = "SELECT SUM(TOTAL_COMPRA) AS TOTAL FROM CUENTAS WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If IsNull(tRs.Fields("TOTAL")) Then
        Acum = 0
    Else
        Acum = tRs.Fields("TOTAL")
    End If
    sBuscar = "SELECT SUM(CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not IsNull(tRs.Fields("TOTAL")) Then
        Acum = Acum - tRs.Fields("TOTAL")
    End If
    Text2.Text = Acum
    sBuscar = "SELECT * FROM VsCxC WHERE ID_CLIENTE = " & Item & " AND PAGADA = 'N' AND (FACTURADO IN (0, 1))"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            MsgBox "El cliente no tiene adeudos"
        Else
            Do While Not .EOF
                Text20.Text = tRs.Fields("ID_VENTA")
                sBusc = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE FOLIO='" & tRs.Fields("FOLIO") & "' "
                Set tRs2 = cnn.Execute(sBusc)
                If (Not IsNull(.Fields("Folio"))) And (.Fields("Folio") <> "") Then
                    Set tLi = ListView2(0).ListItems.Add(, , .Fields("Folio"))
                Else
                    Set tLi = ListView2(1).ListItems.Add(, , .Fields("Id_Venta"))
                End If
                    If Not IsNull(.Fields("Fecha")) Then tLi.SubItems(1) = .Fields("Fecha")
                    If Not IsNull(.Fields("Fecha_Vence")) Then tLi.SubItems(2) = .Fields("Fecha_Vence")
                    If Not IsNull(.Fields("Total_Compra")) Then
                        tLi.SubItems(3) = .Fields("Total_Compra")
                    Else
                        tLi.SubItems(3) = 0
                    End If
                    If Not IsNull(.Fields("Deuda")) Then
                        If Not IsNull(tRs2.Fields("CANT_ABONO")) Then
                            tLi.SubItems(4) = .Fields("Deuda")
                        Else
                             tLi.SubItems(4) = .Fields("Deuda")
                        End If
                    Else
                        tLi.SubItems(4) = 0
                    End If
                tLi.SubItems(5) = .Fields("Id_Cuenta")
                tLi.SubItems(6) = .Fields("ID_VENTA")
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ActualVenta(Venta As String)
    Dim sBuscar As String
    Dim sBusc As String
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    Dim tLi2 As ListItem
    Dim Acum As Double
    Command3.Enabled = False
    ListView2(0).ListItems.Clear
    ListView2(1).ListItems.Clear
    ListView3.ListItems.Clear
    LVDeuda.ListItems.Clear
    txtNo.Text = ""
    txtSaldo.Text = ""
    Text10.Text = ""
    Text9.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text4.Text = ""
    sBuscar = "SELECT ID_NOTA, TOTAL, FECHA FROM NOTA_CREDITO WHERE APLICADA = 'N' AND ID_VENTA IN (" & Venta & ")"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView3.ListItems.Add(, , .Fields("ID_NOTA"))
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(1) = .Fields("TOTAL")
                If Not IsNull(.Fields("Fecha")) Then tLi.SubItems(2) = .Fields("Fecha")
                .MoveNext
            Loop
        End If
    End With
    sBuscar = "SELECT SUM(TOTAL_COMPRA) AS TOTAL FROM CUENTAS WHERE ID_VENTA IN (" & Venta & ")"
    Set tRs = cnn.Execute(sBuscar)
    If IsNull(tRs.Fields("TOTAL")) Then
        Acum = 0
    Else
        Acum = tRs.Fields("TOTAL")
    End If
    sBuscar = "SELECT SUM(CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ID_VENTA IN (" & Venta & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not IsNull(tRs.Fields("TOTAL")) Then
        Acum = Acum - tRs.Fields("TOTAL")
    End If
    Text2.Text = Acum
    sBuscar = "SELECT * FROM VsCxC WHERE ID_VENTA IN (" & Venta & ") AND PAGADA = 'N' AND (FACTURADO IN (0, 1))"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            MsgBox "El cliente no tiene adeudos"
        Else
            Do While Not .EOF
                Text20.Text = tRs.Fields("ID_VENTA")
                sBusc = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE FOLIO = '" & tRs.Fields("FOLIO") & "'"
                Set tRs2 = cnn.Execute(sBusc)
                If (Not IsNull(.Fields("Folio"))) And (.Fields("Folio") <> "") Then
                    Set tLi = ListView2(0).ListItems.Add(, , .Fields("Folio"))
                Else
                    Set tLi = ListView2(1).ListItems.Add(, , .Fields("Id_Venta"))
                End If
                    If Not IsNull(.Fields("Fecha")) Then tLi.SubItems(1) = .Fields("Fecha")
                    If Not IsNull(.Fields("Fecha_Vence")) Then tLi.SubItems(2) = .Fields("Fecha_Vence")
                    If Not IsNull(.Fields("Total_Compra")) Then
                        tLi.SubItems(3) = .Fields("Total_Compra")
                    Else
                        tLi.SubItems(3) = 0
                    End If
                    If Not IsNull(.Fields("Deuda")) Then
                        If Not IsNull(tRs2.Fields("CANT_ABONO")) Then
                            tLi.SubItems(4) = .Fields("Deuda")
                        Else
                             tLi.SubItems(4) = .Fields("Deuda")
                        End If
                    Else
                        tLi.SubItems(4) = 0
                    End If
                tLi.SubItems(5) = .Fields("Id_Cuenta")
                tLi.SubItems(6) = .Fields("ID_VENTA")
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2(Index).SortKey = ColumnHeader.Index - 1
    ListView2(Index).Sorted = True
    ListView2(Index).SortOrder = 1 Xor ListView2(Index).SortOrder
End Sub
Private Sub ListView2_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sBusc As String
    Dim tRs8 As ADODB.Recordset
    Dim todueda As Double
    Dim tocom As Double
    Dim dven As String
    txtSaldo.Text = ""
    txtIdCuenta = Item.SubItems(5)
    Text22.Text = Item
    dven = Item.SubItems(6)
    If Option1(0).Value Then
        sBuscar = "SELECT FOLIO, SUM(DEUDA) AS DEUDA, SUM(TOTAL_COMPRA) AS TOTAL_COMPRA  FROM VsCxC WHERE FOLIO = '" & Item & "' AND PAGADA = 'N'  GROUP BY FOLIO"
        Set tRs = cnn.Execute(sBuscar)
        sBusc = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE FOLIO = '" & Item & "'"
        Set tRs8 = cnn.Execute(sBusc)
        txtNo.Text = Item
        If Not IsNull(tRs8.Fields("CANT_ABONO")) Then
            If Not IsNull(tRs.Fields("DEUDA")) Then
                txtSaldo.Text = Format(CDbl(tRs.Fields("TOTAL_COMPRA")) - CDbl(tRs8.Fields("CANT_ABONO")), "0.00")
            End If
        Else
            txtSaldo.Text = Format(CDbl(tRs.Fields("DEUDA")), "0.00")
        End If
        texidventa = Item.SubItems(6)
    End If
    If Option1(1).Value Then
        sBuscar = "SELECT FOLIO, SUM(DEUDA) AS DEUDA, SUM(TOTAL_COMPRA) AS TOTAL_COMPRA,ID_VENTA  FROM VsCxC WHERE ID_VENTA = '" & Item & "' AND PAGADA = 'N'  GROUP BY FOLIO,ID_VENTA"
        Set tRs = cnn.Execute(sBuscar)
        texidventa = tRs.Fields("ID_VENTA")
        sBusc = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_CUENTA WHERE ID_VENTA = '" & Item & "'"
        Set tRs8 = cnn.Execute(sBusc)
        txtNo.Text = Item
        If Not IsNull(tRs8.Fields("CANT_ABONO")) Then
            If Not IsNull(tRs.Fields("DEUDA")) Then
                txtSaldo.Text = Format(CDbl(tRs.Fields("TOTAL_COMPRA")) - CDbl(tRs8.Fields("CANT_ABONO")), "0.00")
            End If
        Else
            txtSaldo.Text = Format(CDbl(tRs.Fields("DEUDA")), "0.00")
        End If
        texidventa = Item.SubItems(6)
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text13.Text = Item
    Text12.Text = Item.SubItems(1)
    Command3.Enabled = True
End Sub
Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value Then
        ListView2(1).Visible = True
        ListView2(0).Visible = False
        lblNoMov.Caption = "Nota V. No:"
        lblTitulo.Caption = "Notas Pendientes de Pago"
    Else
        ListView2(1).Visible = False
        ListView2(0).Visible = True
        lblNoMov.Caption = "Factura No:"
        lblTitulo.Caption = "Facturas Pendientes de Pago"
    End If
End Sub
Private Sub Option2_Click(Index As Integer)
    If Option2(1).Value Then
        lblTipoMov.Caption = "Numero de Nota V.:"
        Command1.Caption = "Buscar Nota V."
    Else
        lblTipoMov.Caption = "Numero de Factura:"
        Command1.Caption = "Buscar Factura"
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        Frame11.Visible = True
    Else
        Frame11.Visible = False
    End If
End Sub
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        Me.Command2.Enabled = False
    Else
        Me.Command2.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Text1.Text <> "" Then
        If KeyAscii = 13 Then
            Buscar
            ListView1.SetFocus
        End If
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
Private Sub Text10_GotFocus()
    Text10.BackColor = &HFFE1E1
End Sub
Private Sub Text10_LostFocus()
    Text10.BackColor = &H80000005
End Sub
Private Sub Text11_Change()
    If Text11.Text <> "" Then
        cmdVer.Enabled = True
    Else
        cmdVer.Enabled = False
    End If
End Sub
Private Sub Text14_Change()
    If Text14.Text = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text14.Text <> "" Then
            Command1.Value = True
        End If
    End If
End Sub
Private Sub Text14_GotFocus()
    Text14.BackColor = &HFFE1E1
End Sub
Private Sub Text14_LostFocus()
    Text14.BackColor = &H80000005
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text15_GotFocus()
    Text15.BackColor = &HFFE1E1
End Sub
Private Sub Text15_LostFocus()
    Text15.BackColor = &H80000005
End Sub
Private Sub Text16_GotFocus()
    Text16.BackColor = &HFFE1E1
End Sub
Private Sub Text16_LostFocus()
    Text16.BackColor = &H80000005
End Sub
Private Sub Text4_Change()
On Error GoTo ManejaError
    If Text4.Text <> "" Then
        Command3.Enabled = True
    Else
        Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text4.Text <> "" Then
            Me.Command3.SetFocus
        Else
            Text5.SetFocus
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
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    If Text5.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Command3.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HFFE1E1
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &H80000005
End Sub
Private Sub cmdImprimir_Click()
On Error GoTo ManejaError
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    ImprimeDeuda
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub cmdVer_Click()
On Error GoTo ManejaError
    Dim vuelta As Integer
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sBuscar2 As String
    Dim tRs2 As ADODB.Recordset
    Dim fech As String
    Dim fech2 As String
    Dim fech3 As String
    Dim ACUMABONO As Double
    Dim ACUMDEUDA As Double
    Dim TOTCREDDIS As Double
    Dim LIM As Integer
    LVDeuda.ListItems.Clear
    sBuscar = "SELECT FECHA, TOTAL_COMPRA, DEUDA,FOLIO FROM VSCXC WHERE ID_CLIENTE = " & ClvCliente
    If DTFechade.Value = DTFechaAl.Value Then
        sBuscar = sBuscar & " AND FECHA = '" & DTFechade.Value & "'"
    Else
        sBuscar = sBuscar & " AND (FECHA >= '" & DTFechade.Value & "' AND FECHA <= '" & DTFechaAl.Value & "' )"
    End If
    Set tRs = cnn.Execute(sBuscar)
    sBuscar2 = "SELECT FECHA, CANT_ABONO,FOLIO FROM VSABONOS WHERE ID_CLIENTE = " & ClvCliente
    Set tRs2 = cnn.Execute(sBuscar2)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            fech = tRs.Fields("FECHA")
            fech2 = tRs.Fields("FECHA")
            Do While (Not (tRs.EOF)) ' And (fech = fech2)
                Set tLi = LVDeuda.ListItems.Add(, , "COMPRA")
                    If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(1) = tRs.Fields("FECHA")
                    If Not IsNull(tRs.Fields("TOTAL_COMPRA")) Then tLi.SubItems(2) = tRs.Fields("TOTAL_COMPRA")
                    If Not IsNull(tRs.Fields("DEUDA")) Then tLi.SubItems(3) = tRs.Fields("DEUDA")
                    If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(4) = tRs.Fields("FOLIO")
                If Not (tRs.EOF And tRs.BOF) Then
                    fech2 = tRs.Fields("FECHA")
                Else
                    fech2 = ""
                End If
                tRs.MoveNext
            Loop
            If Not (tRs2.EOF And tRs2.BOF) Then
                fech3 = tRs2.Fields("FECHA")
                Do While (Not (tRs2.EOF)) And (fech = fech3)
                    Set tLi = LVDeuda.ListItems.Add(, , "ABONO")
                        If Not IsNull(tRs2.Fields("FECHA")) Then tLi.SubItems(1) = tRs2.Fields("FECHA")
                        If Not IsNull(tRs2.Fields("CANT_ABONO")) Then tLi.SubItems(2) = tRs2.Fields("CANT_ABONO")
                        tLi.SubItems(3) = ""
                        'tLi.SubItems(4) = tRs2.Fields("FOLIO")
                        tRs2.MoveNext
                    If Not (tRs2.EOF) Then
                        fech3 = tRs2.Fields("FECHA")
                    Else
                        fech3 = ""
                    End If
                Loop
            End If
        Loop
        If Not (tRs2.EOF And tRs2.BOF) Then
            Do While (Not (tRs2.EOF))
                Set tLi = LVDeuda.ListItems.Add(, , "ABONO")
                    If Not IsNull(tRs2.Fields("FECHA")) Then tLi.SubItems(1) = tRs2.Fields("FECHA")
                    If Not IsNull(tRs2.Fields("CANT_ABONO")) Then tLi.SubItems(2) = tRs2.Fields("CANT_ABONO")
                    tLi.SubItems(3) = ""
                    If Not IsNull(tRs2.Fields("FOLIO")) Then tLi.SubItems(4) = tRs2.Fields("FOLIO")
                tRs2.MoveNext
            Loop
        End If
        sBuscar = " SELECT SUM(DEUDA) AS TOT FROM CUENTAS WHERE ID_CLIENTE = " & ClvCliente
        Set tRs = cnn.Execute(sBuscar)
        sBuscar = " SELECT SUM(CANT_ABONO) AS TOT FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & ClvCliente
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            ACUMDEUDA = CDbl(tRs.Fields("TOT")) - CDbl(tRs2.Fields("TOT"))
        Else
            ACUMDEUDA = 0
        End If
        If LimCred <> "" Then
            TOTCREDDIS = CDbl(LimCred) - ACUMDEUDA
        Else
            TOTCREDDIS = 0
        End If
        Text8.Text = "TOTAL DE DEUDA   : $ " & ACUMDEUDA & "    LIMITE DE CREDITO  : $ " & LimCred & "    CREDITO DISPONIBLE  : $ " & TOTCREDDIS
    Else
        MsgBox "NO EXISTEN MOVIMIENTOS EN ESTE CLIENTE"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ImprimeDeuda()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim X As Integer
    Dim Y As Integer
    Dim fech As String
    Dim ACUMABONO As Double
    Dim ACUMDEUDA As Double
    Dim LIM As Integer
    ACUMABONO = 0
    fech = DTFechade.Value
    X = 3000
    Y = 20
    LIM = 0
    sBuscar = "SELECT CANT_ABONO, FECHA FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                ACUMABONO = CDbl(ACUMABONO) + CDbl(tRs.Fields("CANT_ABONO"))
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    LIM = 0
    sBuscar = "SELECT TOTAL_COMPRA, FECHA FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_CUENTA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                If tRs.Fields("TOTAL_COMPRA") <> Null Then
                    ACUMDEUDA = CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA"))
                End If
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    ACUMDEUDA = ACUMDEUDA - ACUMABONO
    Enca
    Printer.CurrentY = 2800
    Printer.CurrentX = 100
    Printer.Print "CONCEPTO"
    Printer.CurrentY = 2800
    Printer.CurrentX = 3000
    Printer.Print "FECHA"
    Printer.CurrentY = 2800
    Printer.CurrentX = 6000
    Printer.Print "IMPORTE"
    Printer.CurrentY = 2800
    Printer.CurrentX = 9000
    Printer.Print "PENDIENTE"
    Do While fech <> (DTFechaAl.Value + 1)
        sBuscar = "SELECT FECHA, CANT_ABONO FROM ABONOS_CUENTA WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Printer.CurrentY = X
                Printer.CurrentX = 100
                Printer.Print "ABONO"
                Printer.CurrentY = X
                Printer.CurrentX = 3000
                Printer.Print fech
                Printer.CurrentY = X
                Printer.CurrentX = 6000
                Printer.Print tRs.Fields("CANT_ABONO")
                Printer.CurrentY = X
                Printer.CurrentX = 9000
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) - CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                Printer.Print CDbl(ACUMDEUDA)
                X = X + 200
                Y = Y + 1
                tRs.MoveNext
                If Y = 73 Then
                    Printer.NewPage
                    Enca
                    X = 200
                    Y = 20
                End If
            Loop
        End If
        sBuscar = "SELECT FECHA, TOTAL_COMPRA FROM CUENTAS WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Printer.CurrentY = X
                Printer.CurrentX = 100
                Printer.Print "COMPRA"
                Printer.CurrentY = X
                Printer.CurrentX = 3000
                Printer.Print fech
                Printer.CurrentY = X
                Printer.CurrentX = 6000
                Printer.Print Format(tRs.Fields("TOTAL_COMPRA"), "###,###,##0.00")
                Printer.CurrentY = X
                Printer.CurrentX = 9000
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA")), "0.00")
                Printer.Print CDbl(ACUMDEUDA)
                X = X + 200
                Y = Y + 1
                tRs.MoveNext
                If Y = 73 Then
                    Printer.NewPage
                    Enca
                    X = 200
                    Y = 20
                End If
            Loop
        End If
    Loop
    Printer.CurrentY = X
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentY = X + 200
    Printer.CurrentX = 6500
    Printer.Print "TOTAL DE CREDITO   : $ " & ACUMDEUDA
    Printer.CurrentY = X + 400
    Printer.CurrentX = 6500
    Printer.Print "LIMITE DE CREDITO  : $ " & LimCred
    Printer.CurrentY = X + 600
    Printer.CurrentX = 6500
    Dim TOTCREDDIS As Double
    TOTCREDDIS = CDbl(LimCred) - CDbl(ACUMDEUDA)
    Printer.Print "CREDITO DISPONIBLE : $ " & TOTCREDDIS
    fech = DTFechade.Value + 1
    Printer.EndDoc
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Enca()
On Error GoTo ManejaError
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
    Printer.Print "             ESTADO DE CUENTA"
    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "             CLIENTE : " & NOM
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentY = 2800
    Printer.CurrentX = 100
    Printer.Print "CONCEPTO"
    Printer.CurrentY = 2800
    Printer.CurrentX = 3000
    Printer.Print "FECHA"
    Printer.CurrentY = 2800
    Printer.CurrentX = 6000
    Printer.Print "IMPORTE"
    Printer.CurrentY = 2800
    Printer.CurrentX = 9000
    Printer.Print "PENDIENTE"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub
Private Sub Text9_Change()
    If Text9.Text <> "" Then
        Command3.Enabled = True
    Else
        Command3.Enabled = False
    End If
End Sub
Private Sub Text9_GotFocus()
    Text9.BackColor = &HFFE1E1
End Sub
Private Sub Text9_LostFocus()
    Text9.BackColor = &H80000005
End Sub
