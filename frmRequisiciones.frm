VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRequisiciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos a cotizar y requisiciónes"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   63
      Top             =   3720
      Width           =   975
      Begin VB.Label Label15 
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
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmRequisiciones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisiciones.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   45
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label8 
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
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmRequisiciones.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisiciones.frx":2156
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   43
      Top             =   7320
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmRequisiciones.frx":3B18
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisiciones.frx":3E22
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
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   41
      Top             =   6120
      Width           =   975
      Begin VB.Label Label7 
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
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdImprimir 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmRequisiciones.frx":5F04
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisiciones.frx":620E
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   39
      Top             =   4920
      Width           =   975
      Begin VB.Image cmdAgregar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmRequisiciones.frx":7DE0
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisiciones.frx":80EA
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Productos a Cotizar"
      TabPicture(0)   =   "frmRequisiciones.frx":9A24
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAgregado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lvwRequi2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lvwRequisiciones"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtId_Proveedor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Combo1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command11"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command12"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Check3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "DTPicker1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check4"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Generar Requisición"
      TabPicture(1)   =   "frmRequisiciones.frx":9A40
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text6"
      Tab(1).Control(1)=   "Command7"
      Tab(1).Control(2)=   "Command6"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "lvwProveedores"
      Tab(1).Control(7)=   "lvRequi"
      Tab(1).Control(8)=   "lblFolio"
      Tab(1).Control(9)=   "Label4"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Ver Requisición"
      TabPicture(2)   =   "frmRequisiciones.frx":9A5C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command9"
      Tab(2).Control(1)=   "Command8"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "lvRequiFin"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Precios  de Cotizaciones"
      TabPicture(3)   =   "frmRequisiciones.frx":9A78
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label19"
      Tab(3).Control(1)=   "Label20"
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(3)=   "Label22"
      Tab(3).Control(4)=   "Frame9"
      Tab(3).Control(5)=   "ListView1"
      Tab(3).Control(6)=   "Command10"
      Tab(3).Control(7)=   "Frame12"
      Tab(3).Control(8)=   "Frame13"
      Tab(3).ControlCount=   9
      Begin VB.CheckBox Check4 
         Caption         =   "Agrupar"
         Height          =   255
         Left            =   8760
         TabIndex        =   85
         Top             =   1080
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5760
         TabIndex        =   82
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49872897
         CurrentDate     =   44634
      End
      Begin VB.CheckBox Check3 
         Caption         =   "De la fecha"
         Height          =   255
         Left            =   4560
         TabIndex        =   81
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   79
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Frame Frame13 
         Height          =   1455
         Left            =   -67560
         TabIndex        =   75
         Top             =   600
         Width           =   1455
         Begin VB.OptionButton Option3 
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame Frame12 
         Height          =   615
         Left            =   -74760
         TabIndex        =   67
         Top             =   6960
         Width           =   9375
         Begin VB.Label Label18 
            Height          =   375
            Left            =   7920
            TabIndex        =   70
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label17 
            Height          =   375
            Left            =   3240
            TabIndex        =   69
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label16 
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.CommandButton Command12 
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
         Left            =   8520
         Picture         =   "frmRequisiciones.frx":9A94
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Eliminara todos los articulos marcados con el recuadro"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -74880
         TabIndex        =   65
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton Command11 
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
         Left            =   8520
         Picture         =   "frmRequisiciones.frx":C466
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Eliminara todos los articulos marcados con el recuadro"
         Top             =   7680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Elimar Producto Urgente"
         Height          =   195
         Left            =   8160
         TabIndex        =   61
         Top             =   4560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8760
         TabIndex        =   60
         Top             =   8040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Eliminar producto"
         Height          =   195
         Left            =   8280
         TabIndex        =   59
         Top             =   4800
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000000&
         Height          =   285
         Left            =   8280
         TabIndex        =   58
         Top             =   8040
         Visible         =   0   'False
         Width           =   375
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
         Left            =   -70560
         Picture         =   "frmRequisiciones.frx":EE38
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   56
         Top             =   2640
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7011
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame9 
         Caption         =   "Producto"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   53
         Top             =   600
         Width           =   6255
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1200
            TabIndex        =   54
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label11 
            Caption         =   "PRODUCTO :"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
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
         Left            =   -65400
         Picture         =   "frmRequisiciones.frx":1180A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "+"
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
         Left            =   -65400
         Picture         =   "frmRequisiciones.frx":141DC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   7200
         TabIndex        =   50
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmRequisiciones.frx":16BAE
         Left            =   4440
         List            =   "frmRequisiciones.frx":16BBE
         TabIndex        =   47
         Text            =   "Todos"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "-"
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
         Left            =   -65360
         Picture         =   "frmRequisiciones.frx":16BEA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4560
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
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
         Left            =   -65360
         Picture         =   "frmRequisiciones.frx":195BC
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4080
         Width           =   255
      End
      Begin VB.Frame Frame5 
         Caption         =   "PROVEEDORES ASIGNADOS A ESTA REQUISICION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -69960
         TabIndex        =   31
         Top             =   480
         Width           =   4815
         Begin VB.CommandButton cmdOtroP 
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
            Left            =   1800
            Picture         =   "frmRequisiciones.frx":1BF8E
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2880
            Width           =   1335
         End
         Begin MSComctlLib.ListView lvProvRequi 
            Height          =   2535
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4471
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   4695
         Begin VB.Frame Frame6 
            Caption         =   "REQUISICIONES NO CERRADAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   4455
            Begin MSComctlLib.ListView lvRequiPend 
               Height          =   1935
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   3413
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483633
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
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
            Left            =   3000
            Picture         =   "frmRequisiciones.frx":1E960
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   28
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Numero de Requisicion"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command5 
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
         Left            =   -66720
         Picture         =   "frmRequisiciones.frx":21332
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
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
         Left            =   9640
         Picture         =   "frmRequisiciones.frx":23D04
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
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
         Left            =   9640
         Picture         =   "frmRequisiciones.frx":266D6
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
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
         Left            =   9640
         Picture         =   "frmRequisiciones.frx":290A8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
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
         Left            =   9640
         Picture         =   "frmRequisiciones.frx":2BA7A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5760
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   -70160
         TabIndex        =   15
         Top             =   2160
         Width           =   5055
         Begin VB.TextBox txtNotas 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   480
            Width           =   4815
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "NOTAS"
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
            TabIndex        =   17
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   -70160
         TabIndex        =   8
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   1920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtTel2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtTel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTel3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "DIRECCIÓN"
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
            TabIndex        =   14
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "TELEFONOS"
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
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRequisiciones.frx":2E44C
         Left            =   1320
         List            =   "frmRequisiciones.frx":2E45C
         TabIndex        =   6
         Text            =   "Todos"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtId_Proveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   7920
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwRequisiciones 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5106
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwRequi2 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   5040
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4683
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwProveedores 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   18
         Top             =   1200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvRequi 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   24
         Top             =   3840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvRequiFin 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   26
         Top             =   3960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7435
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7320
         TabIndex        =   84
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49872897
         CurrentDate     =   44634
      End
      Begin VB.Label Label23 
         Caption         =   "al"
         Height          =   255
         Left            =   7080
         TabIndex        =   83
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Nota: Dando  click  sobre el producto que  le interesa, se desplegara  la informacion individualmente"
         Height          =   375
         Left            =   -73800
         TabIndex        =   74
         Top             =   2160
         Width           =   7695
      End
      Begin VB.Label Label21 
         Caption         =   "PRECIO DE VENTA"
         Height          =   255
         Left            =   -66960
         TabIndex        =   73
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "DECRIPCION"
         Height          =   255
         Left            =   -70680
         TabIndex        =   72
         Top             =   6720
         Width           =   3495
      End
      Begin VB.Label Label19 
         Caption         =   "PRODUCTO:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   6720
         Width           =   3855
      End
      Begin VB.Label Label10 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   6600
         TabIndex        =   49
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Marcas:"
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblFolio 
         Height          =   255
         Left            =   -74640
         TabIndex        =   34
         Top             =   8040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Proveedor"
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
         TabIndex        =   19
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Requisiciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblAgregado 
         Alignment       =   2  'Center
         Caption         =   "----------------------------------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   7920
         Width           =   5655
      End
      Begin VB.Label Label6 
         Caption         =   "COTIZACIONES URGENTES"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
End
Attribute VB_Name = "frmRequisiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim StrRep As String
Dim StrRep2 As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check4_Click()
    If Check4.Value = 1 Then
        SaveSetting "APTONER", "ConfigSACC", "AgrupaRequi", "S"
    Else
        SaveSetting "APTONER", "ConfigSACC", "AgrupaRequi", "N"
    End If
    Llenar_Lista_Proveedores
    Llenar_Lista_Requisiciones
End Sub
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    Dim ID_REQUISICION As String
    Dim ID_PROVEEDOR As Integer
    Dim ID_PRODUCTO As String
    Dim Descripcion As String
    Dim CANTIDAD As Double
    Dim DIAS_ENTREGA As Integer
    Dim Precio As Double
    Dim Cont As Integer
    Dim CONT2 As Integer
    Dim FolioR As Integer
    Dim FolioC As Integer
    Dim tLi As ListItem
    Dim Uno As Boolean
    Dim Dos As Boolean
    Cont = 1
    Do While Cont <= lvwRequisiciones.ListItems.Count
        If lvwRequisiciones.ListItems.Item(Cont).Checked Then
            Set tLi = lvRequi.ListItems.Add(, , lvwRequisiciones.ListItems.Item(Cont))
            tLi.SubItems(1) = lvwRequisiciones.ListItems.Item(Cont).ListSubItems(1)
            tLi.SubItems(2) = lvwRequisiciones.ListItems.Item(Cont).ListSubItems(2)
            tLi.SubItems(3) = lvwRequisiciones.ListItems.Item(Cont).ListSubItems(3)
            tLi.SubItems(4) = lvwRequisiciones.ListItems.Item(Cont).ListSubItems(4)
            tLi.SubItems(5) = lvwRequisiciones.ListItems.Item(Cont).ListSubItems(5)
            lvwRequisiciones.ListItems.Remove (Cont)
            lvRequi.ListItems.Item(lvRequi.ListItems.Count).ForeColor = vbBlack
            lvRequi.ListItems.Item(lvRequi.ListItems.Count).Bold = False
        Else
            Cont = Cont + 1
        End If
    Loop
    Cont = 1
    Do While Cont <= lvwRequi2.ListItems.Count
        If lvwRequi2.ListItems.Item(Cont).Checked Then
            Set tLi = lvRequi.ListItems.Add(, , lvwRequi2.ListItems.Item(Cont))
            tLi.SubItems(1) = lvwRequi2.ListItems.Item(Cont).ListSubItems(1)
            tLi.SubItems(2) = lvwRequi2.ListItems.Item(Cont).ListSubItems(2)
            tLi.SubItems(3) = lvwRequi2.ListItems.Item(Cont).ListSubItems(3)
            tLi.SubItems(4) = lvwRequi2.ListItems.Item(Cont).ListSubItems(4)
            tLi.SubItems(5) = lvwRequi2.ListItems.Item(Cont).ListSubItems(5)
            lvwRequi2.ListItems.Remove (Cont)
            lvRequi.ListItems.Item(lvRequi.ListItems.Count).ForeColor = vbRed
            lvRequi.ListItems.Item(lvRequi.ListItems.Count).Bold = True
        Else
            Cont = Cont + 1
        End If
    Loop
    lblFolio.Caption = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    lvRequiFin.ListItems.Clear
    lvProvRequi.ListItems.Clear
    If Text1.Text <> "" Then
        sBuscar = "SELECT * FROM VsREQUISICION WHERE ACTIVO = 0 AND FOLIO = " & Text1.Text
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                Do While Not .EOF
                    Set tLi = lvRequiFin.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")))
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                    If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = Trim(.Fields("FECHA"))
                    If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(5) = Trim(.Fields("ID_REQUISICION"))
                    tLi.SubItems(4) = Trim(.Fields("CONTADOR"))
                    If .Fields("URGENTE") = "S" Then
                        lvRequiFin.ListItems(lvRequiFin.ListItems.Count).ForeColor = vbRed
                        lvRequiFin.ListItems(lvRequiFin.ListItems.Count).Bold = True
                    Else
                        lvRequiFin.ListItems(lvRequiFin.ListItems.Count).ForeColor = vbBlack
                        lvRequiFin.ListItems(lvRequiFin.ListItems.Count).Bold = False
                    End If
                    .MoveNext
                Loop
                sBuscar = "SELECT C.ID_PROVEEDOR, P.NOMBRE FROM COTIZA_REQUI AS C JOIN PROVEEDOR AS P ON C.ID_PROVEEDOR = P.ID_PROVEEDOR WHERE C.FOLIOREQUI = " & Text1.Text & " GROUP BY NOMBRE, C.ID_PROVEEDOR"
                Set tRs2 = cnn.Execute(sBuscar)
                Do While Not tRs2.EOF
                    Set tLi = lvProvRequi.ListItems.Add(, , tRs2.Fields("ID_PROVEEDOR"))
                    If Not IsNull(tRs2.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(tRs2.Fields("NOMBRE"))
                    tRs2.MoveNext
                Loop
                cmdOtroP.Enabled = True
            Else
                MsgBox "EL NUMERO DE REQUISICION NO EXISTE O YA ESTA CERRADO", vbInformation, "SACC"
                cmdOtroP.Enabled = False
            End If
        End With
    Else
        MsgBox "SELECCIONE UNA REQUISION ", vbInformation, "SACC"
    End If
End Sub
Private Sub CmdGuardar_Click()
    sBuscar = "UPDATE REQUISICION SET URGENTE= 'C' WHERE ID_REQUISICION='" & Text5.Text & "'"
    cnn.Execute (sBuscar)
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    Else
        MsgBox "NO HAY PROVEEDORES", vbInformation, "SACC"
    End If
    If Hay_Requisiciones Then
        Llenar_Lista_Requisiciones
    End If
End Sub
Private Sub cmdImprimir_Click()
On Error GoTo Error
    If SSTab1.Tab = 0 Then
        CommonDialog1.Flags = 64
        CommonDialog1.ShowPrinter
        Dim NRegistros As Integer
        Dim Con As Integer
        Dim Hojas As Integer
        Dim Hojastot As Integer
        Dim HojasAprox As Double
        Hojastot = 1
        NRegistros = lvwRequisiciones.ListItems.Count
        For Con = 1 To NRegistros
            If lvwRequisiciones.ListItems.Item(Con).Checked Then Hojastot = Hojastot + 1
        Next Con
        NRegistros = lvwRequi2.ListItems.Count
        For Con = 1 To NRegistros
            If lvwRequi2.ListItems.Item(Con).Checked Then Hojastot = Hojastot + 1
        Next Con
        HojasAprox = Hojastot / 30
        Hojastot = Hojastot / 30
        If HojasAprox - Hojastot > 0 Then
            Hojastot = Hojastot + 1
        End If
        Hojas = 1
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
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR")) / 2
        Printer.Print "LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR"
        If Combo1.Text = "ALMACEN 1" Then
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 1")) / 2
            Printer.Print "ALMACEN 1"
        ElseIf Combo1.Text = "ALMACEN 2" Then
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 2")) / 2
            Printer.Print "ALMACEN 2"
        ElseIf Combo1.Text = "ALMACEN 3" Then
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 3")) / 2
            Printer.Print "ALMACEN 3"
        Else
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 1, ALMACEN 2 Y ALMACEN 3")) / 2
            Printer.Print "ALMACEN 1, ALMACEN 2 Y ALMACEN 3"
        End If
        NRegistros = lvwRequisiciones.ListItems.Count
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "PRODUCTO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9000
        Printer.Print "CANTIDAD"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10100
        Printer.Print "URGENTE"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = POSY + 200
        For Con = 1 To NRegistros
            If lvwRequisiciones.ListItems.Item(Con).Checked Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvwRequisiciones.ListItems.Item(Con)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print lvwRequisiciones.ListItems(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 9200
                Printer.Print lvwRequisiciones.ListItems(Con).SubItems(2)
                Printer.CurrentY = POSY
                Printer.CurrentX = 10300
                Printer.Print "NO"
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 0
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                If POSY >= 14200 Then
                    Printer.Print "                                                                                                                                                                                                                                                 Pagina " & Hojas & " de " & Hojastot
                    Printer.NewPage
                    Hojas = Hojas + 1
                    POSY = 100
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                    Printer.Print VarMen.Text5(0).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR")) / 2
                    Printer.Print "LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR"
                    If Combo1.Text = "ALMACEN 1" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 1")) / 2
                        Printer.Print "ALMACEN 1"
                    ElseIf Combo1.Text = "ALMACEN 2" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 2")) / 2
                        Printer.Print "ALMACEN 2"
                    ElseIf Combo1.Text = "ALMACEN 3" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 3")) / 2
                        Printer.Print "ALMACEN 3"
                    End If
                    NRegistros = lvwRequisiciones.ListItems.Count
                    POSY = 2200
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2200
                    Printer.Print "Descripcion"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 9000
                    Printer.Print "CANTIDAD"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 10100
                    Printer.Print "URGENTE"
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    POSY = POSY + 200
                End If
            End If
        Next Con
        NRegistros = lvwRequi2.ListItems.Count
        For Con = 1 To NRegistros
            If lvwRequi2.ListItems.Item(Con).Checked Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvwRequi2.ListItems.Item(Con)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print lvwRequi2.ListItems(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 9200
                Printer.Print lvwRequi2.ListItems(Con).SubItems(2)
                Printer.CurrentY = POSY
                Printer.CurrentX = 10200
                Printer.Print "SI"
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 0
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                If POSY >= 14200 Then
                    Printer.Print "                                                                                                                                                                                                                                                 Pagina " & Hojas & " de " & Hojastot
                    Printer.NewPage
                    Hojas = Hojas + 1
                    POSY = 100
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
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR")) / 2
                    Printer.Print "LISTADO DE PRODUCTOS PENDIENTES DE COTIZAR"
                    If Combo1.Text = "ALMACEN 1" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 1")) / 2
                        Printer.Print "ALMACEN 1"
                    ElseIf Combo1.Text = "ALMACEN 2" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 2")) / 2
                        Printer.Print "ALMACEN 2"
                    ElseIf Combo1.Text = "ALMACEN 3" Then
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("ALMACEN 3")) / 2
                        Printer.Print "ALMACEN 3"
                    End If
                    NRegistros = lvwRequisiciones.ListItems.Count
                    POSY = 2200
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2200
                    Printer.Print "Descripcion"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 9000
                    Printer.Print "CANTIDAD"
                    Printer.CurrentX = 9600
                    Printer.Print "URGENTE"
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    POSY = POSY + 200
                End If
            End If
        Next Con
        Printer.Print ""
        Printer.Print "FIN DEL LISTADO"
        Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------      Pagina " & Hojas & " de " & Hojastot
        Printer.Print "      USUARIO: " & VarMen.lblHola.Caption & " Fecha de Impresion: " & Format(Date, "dd/mm/yyyy")
        Printer.EndDoc
        CommonDialog1.Copies = 1
    Else
        Imprimir2
    End If
Exit Sub
Error:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub cmdOtroP_Click()
    Dim Cont As Integer
    Dim tLi As ListItem
    lblFolio = Text1.Text
    lvRequi.ListItems.Clear
    For Cont = 1 To lvRequiFin.ListItems.Count
        If lvRequiFin.ListItems.Item(Cont).Checked = True Then
            Set tLi = lvRequi.ListItems.Add(, , lvRequiFin.ListItems.Item(Cont))
            tLi.SubItems(1) = lvRequiFin.ListItems.Item(Cont).SubItems(1)
            tLi.SubItems(2) = lvRequiFin.ListItems.Item(Cont).SubItems(2)
            tLi.SubItems(3) = lvRequiFin.ListItems.Item(Cont).SubItems(3)
            tLi.SubItems(4) = lvRequiFin.ListItems.Item(Cont).SubItems(4)
            tLi.SubItems(5) = lvRequiFin.ListItems.Item(Cont).SubItems(5)
        End If
    Next Cont
    Label14.Caption = "Guardar"
    cmdAgregar.Enabled = True
    CmdImprimir.Enabled = False
    SSTab1.Tab = 1
End Sub
Private Sub Combo1_Click()
    Llenar_Lista_Proveedores
    Llenar_Lista_Requisiciones
End Sub
Private Sub Combo2_Click()
    Llenar_Lista_Proveedores
    Llenar_Lista_Requisiciones
End Sub
Private Sub Command1_Click()
    Dim Cont As Integer
    For Cont = 1 To lvwRequisiciones.ListItems.Count
        lvwRequisiciones.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command10_Click()
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    On Error GoTo ManejaError
    If Text3.Text <> "" Then
        If Option4.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION FROM VSVENTAS WHERE ID_PRODUCTO LIKE '%" & Trim(Text3.Text) & "%' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"   'Cambiado 25/09/06
        End If                                                                                                                                     'Se cambio Almacen3 por VsVentas
        If Option3.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION FROM VSVENTAS WHERE Descripcion LIKE '%" & Trim(Text3.Text) & "%' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"   'Cambiado 25/09/06
        End If                                                                                                                           'Se cambio Almacen3 por VsVentas
        If Option5.Value = True Then
            sBus = "SELECT ID_PRODUCTO FROM ENTRADA_PRODUCTO WHERE CODIGO_BARAS = '" & Text3.Text & "'"
            Set tRs = cnn.Execute(sBus)
            If Not (tRs.EOF And tRs.BOF) Then
                sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD, TIPO, CLASIFICACION FROM VSVENTAS WHERE ID_PRODUCTO LIKE '%" & tRs.Fields("ID_PRODUCTO") & "%' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"   'Cambiado 25/09/06
            Else
                MsgBox "EL CODIGO DE BARRAS NO ESTA REGISTRADO, INTENTE OTRO MODO DE BUSQUEDA!", vbInformation, "SACC"
            End If
        End If
        If sBus <> "" Then
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView1.ListItems.Clear
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                            If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                                tLi.SubItems(2) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "###,###,##0.00")
                            End If
                            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = .Fields("CANTIDAD") & ""
                            If Not IsNull(.Fields("CLASIFICACION")) Then tLi.SubItems(4) = .Fields("CLASIFICACION") & ""
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
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
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command11_Click()
    Dim Cont As Integer
    Dim sBuscar As String
    For Cont = 1 To lvwRequi2.ListItems.Count
        If lvwRequi2.ListItems(Cont).Checked Then
            sBuscar = "DELETE FROM REQUISICION WHERE ID_REQUISICION IN (" & lvwRequi2.ListItems(Cont).SubItems(5) & ")"
            cnn.Execute (sBuscar)
        End If
    Next
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    Else
        MsgBox "NO HAY PROVEEDORES", vbInformation, "SACC"
    End If
    If Hay_Requisiciones Then
        Llenar_Lista_Requisiciones
    End If
End Sub
Private Sub Command12_Click()
    Dim Cont As Integer
    Dim sBuscar As String
    For Cont = 1 To lvwRequisiciones.ListItems.Count
        If lvwRequisiciones.ListItems(Cont).Checked Then
            sBuscar = "DELETE FROM REQUISICION WHERE ID_REQUISICION IN (" & lvwRequisiciones.ListItems(Cont).SubItems(5) & ")"
            cnn.Execute (sBuscar)
        End If
    Next
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    Else
        MsgBox "NO HAY PROVEEDORES", vbInformation, "SACC"
    End If
    If Hay_Requisiciones Then
        Llenar_Lista_Requisiciones
    End If
End Sub
Private Sub Command6_Click()
    Dim Cont As Integer
    For Cont = 1 To lvRequi.ListItems.Count
        lvRequi.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command7_Click()
    Dim Cont As Integer
    For Cont = 1 To lvRequi.ListItems.Count
        lvRequi.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub DTPicker2_CloseUp()
    Llenar_Lista_Requisiciones
End Sub
Private Sub Image1_Click()
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
    If lvwRequisiciones.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = lvwRequisiciones.ColumnHeaders.Count
            For Con = 1 To lvwRequisiciones.ColumnHeaders.Count
                StrCopi = StrCopi & lvwRequisiciones.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To lvwRequisiciones.ListItems.Count
                If lvwRequisiciones.ListItems(Con).Checked Then
                    StrCopi = StrCopi & lvwRequisiciones.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & lvwRequisiciones.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                End If
            Next
            'archivo TXT
        End If
    End If
    If lvwRequi2.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = lvwRequi2.ColumnHeaders.Count
            For Con = 1 To lvwRequi2.ColumnHeaders.Count
                StrCopi = StrCopi & lvwRequi2.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To lvwRequi2.ListItems.Count
                If lvwRequi2.ListItems(Con).Checked Then
                    StrCopi = StrCopi & lvwRequi2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & lvwRequi2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                End If
            Next
            'archivo TXT
            
        End If
        
    End If
    Dim foo As Integer
    foo = FreeFile
    Open Ruta For Output As #foo
        Print #foo, StrCopi
    Close #foo
    ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Label16.Caption = Item
    Label17.Caption = ListView1.SelectedItem.SubItems(1)
     Label18.Caption = ListView1.SelectedItem.SubItems(2)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command10.Value = True
    End If
End Sub
Private Sub Command2_Click()
    Dim Cont As Integer
    For Cont = 1 To lvwRequisiciones.ListItems.Count
        lvwRequisiciones.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub Command3_Click()
    Dim Cont As Integer
    For Cont = 1 To lvwRequi2.ListItems.Count
        lvwRequi2.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command4_Click()
    Dim Cont As Integer
    For Cont = 1 To lvwRequi2.ListItems.Count
        lvwRequi2.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub Command5_Click()
    Dim Cont As Integer
    Dim tLi As ListItem
    Cont = 1
    Do While Cont <= lvRequi.ListItems.Count
        If lvRequi.ListItems.Item(Cont).Checked Then
            If lvRequi.ListItems.Item(Cont).ForeColor = vbBlack Then
                Set tLi = lvwRequisiciones.ListItems.Add(, , lvRequi.ListItems.Item(Cont))
                    tLi.SubItems(1) = lvRequi.ListItems.Item(Cont).ListSubItems(1)
                    tLi.SubItems(2) = lvRequi.ListItems.Item(Cont).ListSubItems(2)
                    tLi.SubItems(3) = lvRequi.ListItems.Item(Cont).ListSubItems(3)
                    tLi.SubItems(4) = lvRequi.ListItems.Item(Cont).ListSubItems(4)
                    tLi.SubItems(5) = lvRequi.ListItems.Item(Cont).ListSubItems(5)
            Else
                Set tLi = lvwRequi2.ListItems.Add(, , lvRequi.ListItems.Item(Cont))
                    tLi.SubItems(1) = lvRequi.ListItems.Item(Cont).ListSubItems(1)
                    tLi.SubItems(2) = lvRequi.ListItems.Item(Cont).ListSubItems(2)
                    tLi.SubItems(3) = lvRequi.ListItems.Item(Cont).ListSubItems(3)
                    tLi.SubItems(4) = lvRequi.ListItems.Item(Cont).ListSubItems(4)
                    tLi.SubItems(5) = lvRequi.ListItems.Item(Cont).ListSubItems(5)
                lvwRequi2.ListItems.Item(lvwRequi2.ListItems.Count).ForeColor = vbRed
                lvwRequi2.ListItems.Item(lvwRequi2.ListItems.Count).Bold = True
            End If
            lvRequi.ListItems.Remove (Cont)
        Else
            Cont = Cont + 1
        End If
    Loop
End Sub
Private Sub Command8_Click()
    Dim Cont As Integer
    For Cont = 1 To lvRequiFin.ListItems.Count
        lvRequiFin.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command9_Click()
    Dim Cont As Integer
    For Cont = 1 To lvwRequi2.ListItems.Count
        lvwRequiFin.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1.Value = Date - 15
    DTPicker2.Value = Date
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    If VarMen.Text1(77).Text = "N" Then
       Command11.Enabled = False
       Command12.Enabled = False
    End If
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    With lvwRequisiciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Cantidad", 1000, 2
        .ColumnHeaders.Add , , "Fecha", 1100, 2
        .ColumnHeaders.Add , , "Contador", 0, 2
        .ColumnHeaders.Add , , "Id Requisición", 0
        .ColumnHeaders.Add , , "Proveedor", 0
        .ColumnHeaders.Add , , "Comentarios", 3500
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Último Proveedor", 3500
        .ColumnHeaders.Add , , "Última Compra", 1500
        .ColumnHeaders.Add , , "Preco Anterior", 1500
        .ColumnHeaders.Add , , "Días de Entrega", 0
        .ColumnHeaders.Add , , "Marca", 1500
    End With
    With lvRequi
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Cantidad", 1000, 2
        .ColumnHeaders.Add , , "Fecha", 1100, 2
        .ColumnHeaders.Add , , "Contador", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
    With lvRequiFin
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Cantidad", 1000, 2
        .ColumnHeaders.Add , , "Fecha", 1100, 2
        .ColumnHeaders.Add , , "Contador", 0, 2
        .ColumnHeaders.Add , , "Id Requisición", 0
    End With
    With lvwRequi2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Cantidad", 1100, 2
        .ColumnHeaders.Add , , "Fecha", 1000, 2
        .ColumnHeaders.Add , , "Contador", 0, 2
        .ColumnHeaders.Add , , "Id Requisición", 0
        .ColumnHeaders.Add , , "Comentarios", 3500
        .ColumnHeaders.Add , , "Id Requisición", 0
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Último Proveedor", 3500
        .ColumnHeaders.Add , , "Última Compra", 1500
        .ColumnHeaders.Add , , "Precio Anterior", 1500
        .ColumnHeaders.Add , , "Días de Entrega", 0
        .ColumnHeaders.Add , , "MArca", 1500
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clv Producto", 1600
        .ColumnHeaders.Add , , "Descripción", 3600
        .ColumnHeaders.Add , , "Precio", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Clasificación", 0
    End With
    With lvRequiPend
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Folio Requisición", 4000
    End With
    With Me.lvwProveedores
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        If VarMen.TxtEmp(12).Text = "EXTENDIDO" Then
            .MultiSelect = True
        Else
            .MultiSelect = False
        End If
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Proveedor", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
    With lvProvRequi
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Proveedor", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
    Combo2.Clear
    sBuscar = "SELECT MARCA From ALMACEN3 GROUP BY MARCA Union SELECT MARCA From ALMACEN2 GROUP BY MARCA Union SELECT MARCA From ALMACEN1 GROUP BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Combo2.AddItem ("Todas")
        Combo2.Text = "Todas"
        Do While Not (tRs.EOF)
            If Trim(tRs.Fields("MARCA")) <> "" Then Combo2.AddItem (tRs.Fields("MARCA"))
            tRs.MoveNext
        Loop
    Else
        Combo2.Text = "NO HAY MARCAS"
        Combo2.Enabled = False
    End If
    If GetSetting("APTONER", "ConfigSACC", "AgrupaRequi", "N") = "S" Then
        Check4.Value = 1
    Else
        Check4.Value = 0
        Llenar_Lista_Proveedores
        Llenar_Lista_Requisiciones
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Requisiciones()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim Cond As String
    Dim tRs3 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Cond = ""
    sqlQuery = "UPDATE REQUISICION SET CONTADOR = '0' WHERE (CONTADOR = '*')"
    On Error Resume Next
    cnn.Execute (sqlQuery)
    Me.lvwRequisiciones.ListItems.Clear
    Me.lvwRequi2.ListItems.Clear
    If Combo1.Text <> "Todos" Then
        Cond = " AND ALMACEN = 'A" & Right(Combo1.Text, 1) & "'"
    End If
    If Combo2.Text <> "Todas" And Combo2.Enabled Then
        Cond = Cond & " AND MARCA = '" & Combo2.Text & "'"
    Else
        Cond = Cond & " AND MARCA LIKE '%'"
    End If
    If Text2.Text <> "" Then
        If Cond = "" Then
            Cond = " AND ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        Else
            Cond = Cond & " AND ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        End If
    End If
    If Text7.Text <> "" Then
        If Cond = "" Then
            Cond = " AND COMENTARIO LIKE '%" & Replace(Text7.Text, " ", "%' AND COMENTARIO LIKE '%") & "%'"
        Else
            Cond = Cond & " AND COMENTARIO LIKE '%" & Replace(Text7.Text, " ", "%' AND COMENTARIO LIKE '%") & "%'"
        End If
    End If
    If Check3.Value = 1 Then
        If Cond = "" Then
            Cond = " AND FECHA BETWEEN '" & DTPicker1.Value & "'  AND '" & DTPicker2.Value & "' "
        Else
            Cond = Cond & " AND FECHA BETWEEN '" & DTPicker1.Value & "'  AND '" & DTPicker2.Value & "'"
        End If
    End If
    'CONSULTA EN COMENTARIOS ES PARA AGRUPAR POR PRODUCTO DEJANDO EL ID_REQUISICION SEPARADO POR COMAS
    If Check4.Value = 1 Then
        sqlQuery = "SELECT STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0)  FOR XML PATH('')), 1, 1, '') AS ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD, (SELECT TOP 1 FECHA FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0) ORDER BY FECHA) AS FECHA, 0 AS CONTADOR, URGENTE, STUFF(( SELECT ',' + COMENTARIO FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0)  FOR XML PATH('')), 1, 1, '') AS COMENTARIO, MARCA FROM REQUISICION RE Where (ACTIVO = 0) And (Folio = 0) " & Cond & " GROUP BY ID_PRODUCTO, DESCRIPCION, URGENTE, MARCA ORDER BY ID_PRODUCTO"
    Else
        sqlQuery = "SELECT ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, COMENTARIO, CANTIDAD, FECHA, CONTADOR, URGENTE, MARCA FROM REQUISICION WHERE ACTIVO = 0 AND FOLIO = 0" & Cond
    End If
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                If Not IsNull(.Fields("CONTADOR")) Then
                    Cont = .Fields("CONTADOR")
                End If
                If .Fields("URGENTE") = "N" Then
                    Set tLi = lvwRequisiciones.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")))
                    pro = .Fields("ID_PRODUCTO")
                    sqlQuery = "SELECT ID_PRODUCTO FROM PRODUCTOS_CONSUMIBLES WHERE ID_PRODUCTO= '" & .Fields("ID_PRODUCTO") & "'"
                    Set tRs3 = cnn.Execute(sqlQuery)
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        tLi.ForeColor = &HFF0000
                    End If
                    sqlQuery = "SELECT * FROM vscomprasproveedor2 WHERE  ID_PRODUCTO= '" & pro & "'"
                    Set tRs2 = cnn.Execute(sqlQuery)
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                    If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = Trim(.Fields("FECHA"))
                    If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(5) = Trim(.Fields("ID_REQUISICION"))
                    If Not IsNull(.Fields("CONTADOR")) Then tLi.SubItems(4) = Trim(.Fields("CONTADOR"))
                    If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(7) = Trim(.Fields("COMENTARIO"))
                    'If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(6) = Trim(.Fields("COMENTARIO"))
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        If Not IsNull(tRs2.Fields("ID_PROVEEDOR")) Then
                           tLi.SubItems(8) = tRs2.Fields("ID_PROVEEDOR")
                        Else
                           tLi.SubItems(8) = 0
                        End If
                        If Not IsNull(tRs2.Fields("NOMBRE")) Then
                            tLi.SubItems(9) = tRs2.Fields("NOMBRE")
                        Else
                            tLi.SubItems(9) = 0
                        End If
                        If Not IsNull(tRs2.Fields("FECHA")) Then
                            tLi.SubItems(10) = tRs2.Fields("FECHA")
                        Else
                            tLi.SubItems(10) = 0
                        End If
                        If Not IsNull(tRs2.Fields("PRECIO")) Then
                            tLi.SubItems(11) = tRs2.Fields("PRECIO")
                        Else
                            tLi.SubItems(11) = 0
                        End If
                        If Not IsNull(.Fields("MARCA")) Then
                            tLi.SubItems(13) = .Fields("MARCA")
                        Else
                            tLi.SubItems(13) = ""
                        End If
                    Else
                        tLi.SubItems(8) = 0
                        tLi.SubItems(9) = 0
                        tLi.SubItems(10) = 0
                        tLi.SubItems(11) = 0
                        tLi.SubItems(13) = ""
                    End If
                End If
'                .MoveNext
'            Loop
'            StrRep = sqlQuery
'        End If
'    End With
'    'CONSULTA EN COMENTARIOS ES PARA AGRUPAR POR PRODUCTO DEJANDO EL ID_REQUISICION SEPARADO POR COMAS
'    'SQLQUERY = "SELECT STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0)  FOR XML PATH('')), 1, 1, '') AS ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD, (SELECT TOP 1 FECHA FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0) ORDER BY FECHA) AS FECHA, 0 AS CONTADOR, URGENTE, STUFF(( SELECT ',' + COMENTARIO FROM REQUISICION RQ WHERE RQ.ID_PRODUCTO = RE.ID_PRODUCTO AND (ACTIVO = 0) AND (FOLIO = 0)  FOR XML PATH('')), 1, 1, '') AS COMENTARIO, MARCA FROM REQUISICION RE Where (ACTIVO = 0) And (Folio = 0) GROUP BY ID_PRODUCTO, DESCRIPCION, URGENTE, MARCA ORDER BY ID_PRODUCTO"
'    sqlQuery = "SELECT ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, CONTADOR, URGENTE, COMENTARIO, MARCA FROM REQUISICION WHERE ACTIVO = 0 AND FOLIO = 0" & Cond
'    Set tRs = cnn.Execute(sqlQuery)
'    With tRs
'        If Not (.BOF And .EOF) Then
'            Do While Not .EOF
'                If Not IsNull(.Fields("CONTADOR")) Then
'                    Cont = .Fields("CONTADOR")
'                End If
                If .Fields("URGENTE") = "S" Then
                    Set tLi = lvwRequi2.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")))
                    pro = .Fields("ID_PRODUCTO")
                    sqlQuery = "SELECT * FROM vscomprasproveedor2 WHERE  ID_PRODUCTO= '" & pro & "'"
                    Set tRs2 = cnn.Execute(sqlQuery)
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                    If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = Trim(.Fields("FECHA"))
                    If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(5) = Trim(.Fields("ID_REQUISICION"))
                    tLi.SubItems(4) = "0"
                    If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(6) = Trim(.Fields("COMENTARIO"))
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        tLi.SubItems(8) = tRs2.Fields("ID_PROVEEDOR")
                    Else
                        tLi.SubItems(8) = 0
                    End If
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        tLi.SubItems(9) = tRs2.Fields("NOMBRE")
                    Else
                        tLi.SubItems(9) = 0
                    End If
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        tLi.SubItems(10) = tRs2.Fields("FECHA")
                    Else
                        tLi.SubItems(10) = 0
                    End If
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        tLi.SubItems(11) = tRs2.Fields("PRECIO")
                    Else
                        tLi.SubItems(11) = 0
                    End If
                    If Not (tRs2.BOF And tRs2.EOF) Then
                        If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(13) = .Fields("MARCA")
                    Else
                        tLi.SubItems(13) = ""
                    End If
                    sqlQuery = "SELECT ID_PRODUCTO FROM PRODUCTOS_CONSUMIBLES WHERE ID_PRODUCTO= '" & .Fields("ID_PRODUCTO") & "'"
                    Set tRs3 = cnn.Execute(sqlQuery)
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        tLi.ForeColor = &HFF0000
                    Else
                        lvwRequi2.ListItems(lvwRequi2.ListItems.Count).ForeColor = vbRed
                    End If
                    lvwRequi2.ListItems(lvwRequi2.ListItems.Count).Bold = False
                End If
                .MoveNext
            Loop
            StrRep2 = sqlQuery
        End If
    End With
    sqlQuery = "SELECT FOLIO FROM REQUISICION WHERE ACTIVO = 0 AND FOLIO <> 0 AND ID_REQUISICION IN (SELECT ID_REQUISICION FROM COTIZA_REQUI) GROUP BY FOLIO"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvRequiPend.ListItems.Add(, , Trim(.Fields("FOLIO")))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Requisiciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_REQUISICION)ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_REQUISICION") <> 0 Then
            Hay_Requisiciones = True
        Else
            Hay_Requisiciones = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    Dim Cond As String
    Me.lvwProveedores.ListItems.Clear
    If Combo1.Text <> "Todos" Then
        If Combo1.Text = "ALMACEN 1" Then
            Cond = " AND ALMACEN1 = 'S' "
        End If
        If Combo1.Text = "ALMACEN 2" Then
            Cond = " AND ALMACEN2 = 'S' "
        End If
        If Combo1.Text = "ALMACEN 3" Then
            Cond = " AND ALMACEN3 = 'S' "
        End If
        Cond = Cond & " AND NOMBRE LIKE '%" & Text6.Text & "%'"
        Cond = Cond & " ORDER BY NOMBRE"
    Else
        Cond = "AND NOMBRE LIKE '%" & Text6.Text & "%' ORDER BY NOMBRE"
    End If
    sqlQuery = "SELECT ID_PROVEEDOR, NOMBRE, DIRECCION, COLONIA, CIUDAD, CP, RFC, TELEFONO1, TELEFONO2, TELEFONO3, NOTAS, ESTADO, PAIS FROM PROVEEDOR WHERE ELIMINADO = 'N' AND EMAIL IS NOT NULL AND EMAIL <> '' " & Cond
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwProveedores.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(.Fields("DIRECCION"))
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(.Fields("COLONIA"))
                If Not IsNull(.Fields("CP")) Then tLi.SubItems(4) = Trim(.Fields("CP"))
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(5) = Trim(.Fields("RFC"))
                If Not IsNull(.Fields("TELEFONO1")) Then tLi.SubItems(6) = Trim(.Fields("TELEFONO1"))
                If Not IsNull(.Fields("TELEFONO2")) Then tLi.SubItems(7) = Trim(.Fields("TELEFONO2"))
                If Not IsNull(.Fields("TELEFONO3")) Then tLi.SubItems(8) = Trim(.Fields("TELEFONO3"))
                If Not IsNull(.Fields("NOTAS")) Then tLi.SubItems(9) = Trim(.Fields("NOTAS"))
                If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = Trim(.Fields("ESTADO"))
                If Not IsNull(.Fields("PAIS")) Then tLi.SubItems(11) = Trim(.Fields("PAIS"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Proveedores() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_PROVEEDOR)ID_PROVEEDOR FROM PROVEEDOR WHERE ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_PROVEEDOR") <> 0 Then
            Hay_Proveedores = True
        Else
            Hay_Proveedores = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image8_Click()
    Dim NoPed As String
    If Puede_Agregar Then
        Dim NumProv As Integer
        If lblFolio.Caption <> "" Then
            FolioR = CDbl(Text1.Text)
            Uno = False
        Else
            FolioR = 0
            Uno = True
        End If
        NumProv = lvwProveedores.ListItems.Count
        For CONT2 = 1 To NumProv
            Dos = True
            If lvwProveedores.ListItems.Item(CONT2).Selected Then
                ID_PROVEEDOR = lvwProveedores.ListItems.Item(CONT2)
                For Cont = 1 To lvRequi.ListItems.Count
                    ID_REQUISICION = lvRequi.ListItems.Item(Cont).ListSubItems(5)
                    ID_PRODUCTO = lvRequi.ListItems.Item(Cont)
                    Descripcion = lvRequi.ListItems.Item(Cont).ListSubItems(1)
                    CANTIDAD = lvRequi.ListItems.Item(Cont).ListSubItems(2)
                    sqlQuery = "SELECT * FROM COTIZA_REQUI WHERE ID_REQUISICION IN (" & ID_REQUISICION & ") AND ID_PROVEEDOR = " & ID_PROVEEDOR & " AND ID_PRODUCTO = '" & ID_PRODUCTO & "' AND CANTIDAD = " & CANTIDAD & " AND FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
                    Set tRs = cnn.Execute(sqlQuery)
                    If tRs.EOF And tRs.BOF Then
                        If Dos Then
                            sqlQuery = "SELECT FOLIO FROM COTIZA_REQUI ORDER BY FOLIO DESC"
                            Set tRs = cnn.Execute(sqlQuery)
                            If tRs.EOF And tRs.BOF Then
                                FolioC = 1
                            Else
                                FolioC = CDbl(tRs.Fields("FOLIO")) + 1
                            End If
                            Dos = False
                        End If
                        sqlQuery = "SELECT FOLIO, NO_PEDIDO FROM REQUISICION ORDER BY FOLIO DESC"
                        Set tRs = cnn.Execute(sqlQuery)
                        If Not IsNull(tRs.Fields("NO_PEDIDO")) Then NoPed = tRs.Fields("NO_PEDIDO")
                        If Uno Then
                            FolioR = CDbl(tRs.Fields("FOLIO")) + 1
                            Uno = False
                        End If
                        sqlQuery = "INSERT INTO COTIZA_REQUI (ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, FOLIO, FOLIOREQUI, NO_PEDIDO) VALUES (" & Replace(ID_REQUISICION, ",", ", " & ID_PROVEEDOR & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", '" & Format(Date, "dd/mm/yyyy") & "', " & FolioC & ", " & FolioR & ", '" & NoPed & "'), (") & ", " & ID_PROVEEDOR & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", '" & Format(Date, "dd/mm/yyyy") & "', " & FolioC & ", " & FolioR & ", '" & NoPed & "');"
                        cnn.Execute (sqlQuery)
                        sqlQuery = "UPDATE REQUISICION SET CONTADOR = CONTADOR + 1, FOLIO = " & FolioR & " WHERE ID_REQUISICION IN (" & ID_REQUISICION & ")"
                        cnn.Execute (sqlQuery)
                    Else
                        MsgBox "PRODUCTO " & lvRequi.ListItems.Item(Cont) & " YA ASIGNADO A " & lvwProveedores.ListItems.Item(CONT2).SubItems(1), vbInformation, "SACC"
                    End If
                Next Cont
            End If
        Next CONT2
        If FolioR > 0 Then
            Text1.Text = FolioR
            cmdBuscar.Value = True
            cmdAgregar.Enabled = False
            CmdImprimir.Enabled = True
            SSTab1.Tab = 2
        End If
        lblFolio.Caption = ""
        lvRequi.ListItems.Clear
        Llenar_Lista_Requisiciones
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvRequiPend_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvRequiPend.ListItems.Count > 0 Then
        Text1.Text = Item
    End If
End Sub
Private Sub lvwProveedores_Click()
    Me.lblAgregado.Caption = "----------------------------------------------"
End Sub
Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtId_Proveedor.Text = Item
    Label4.Caption = Item.SubItems(1)
    Me.txtDireccion.Text = Item.SubItems(2) + " " + Item.SubItems(3) + " " + Item.SubItems(4) + " " + Item.SubItems(5) + " " + Item.SubItems(10) + " " + Item.SubItems(11)
    Me.txtTel1.Text = Item.SubItems(6)
    Me.txtTel2.Text = Item.SubItems(7)
    Me.txtTel3.Text = Item.SubItems(8)
    Me.txtNotas.Text = Item.SubItems(9)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwRequi2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRequi2.SortKey = ColumnHeader.Index - 1
    lvwRequi2.Sorted = True
    lvwRequi2.SortOrder = 1 Xor lvwRequi2.SortOrder
End Sub
Private Sub lvwRequi2_DblClick()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If lvwRequi2.ListItems.Count > 0 Then
        frmVerRastreo.Text1.Text = lvwRequi2.SelectedItem.SubItems(5)
        frmVerRastreo.Text2.Text = lvwRequi2.SelectedItem
        frmVerRastreo.Show vbModal
    End If
End Sub
Private Sub lvwRequisiciones_Click()
    If lvwRequisiciones.ListItems.Count > 0 Then
        Me.lblAgregado.Caption = "----------------------------------------------"
        Text5.Text = lvwRequisiciones.SelectedItem.SubItems(5)
        Command12.Visible = True
        If VarMen.Text1(77).Text = "N" Then
            Command12.Enabled = False
        End If
    End If
End Sub
Private Sub lvwRequi2_Click()
    If lvwRequi2.ListItems.Count > 0 Then
        Text4.Text = lvwRequi2.SelectedItem.SubItems(5)
        Command11.Visible = True
    End If
End Sub
Function Puede_Agregar() As Boolean
On Error GoTo ManejaError
    If Me.lvRequi.ListItems.Count = 0 Then
        MsgBox "NO HAY PRODUCTOS EN LA REQUISICION", vbInformation, "SACC"
        Puede_Agregar = False
        Exit Function
    End If
    If Me.txtId_Proveedor.Text = "" Then
        MsgBox "SELECCIONE EL PROVEEDOR", vbInformation, "SACC"
        Puede_Agregar = False
        Exit Function
    End If
    Puede_Agregar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub lvwRequisiciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRequisiciones.SortKey = ColumnHeader.Index - 1
    lvwRequisiciones.Sorted = True
    lvwRequisiciones.SortOrder = 1 Xor lvwRequisiciones.SortOrder
End Sub
Private Sub lvwRequisiciones_DblClick()
    If lvwRequisiciones.ListItems.Count > 0 Then
        frmVerRastreo.Text1.Text = lvwRequisiciones.SelectedItem.SubItems(5)
        frmVerRastreo.Text2.Text = lvwRequisiciones.SelectedItem
        frmVerRastreo.Show vbModal
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Frame14.Visible = True
        Frame8.Visible = False
        cmdAgregar.Enabled = True
        CmdImprimir.Enabled = True
    ElseIf SSTab1.Tab = 1 Then
        Frame14.Visible = False
        Frame8.Visible = True
        cmdAgregar.Enabled = True
        CmdImprimir.Enabled = False
    Else
        Frame14.Visible = False
        Frame8.Visible = False
        cmdAgregar.Enabled = False
        CmdImprimir.Enabled = True
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Valido As String
    Valido = "1234567890"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Else
        If (KeyAscii = 13) And (Text1.Text <> "") Then
            cmdBuscar.Value = True
        End If
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_Change()
    If Text1.Text <> "" Then
        cmdBuscar.Enabled = True
        cmdOtroP.Enabled = True
    Else
        cmdBuscar.Enabled = False
        cmdOtroP.Enabled = False
    End If
End Sub
Private Sub Imprimir2()
On Error GoTo ManjaError
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Dim NRegistros As Integer
    Dim Con As Integer
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
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("REQUISICION #" & FolioR)) / 2
    Printer.Print "REQUISICION #" & FolioR
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    NRegistros = lvRequiFin.ListItems.Count
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 2200
    Printer.Print "Descripcion"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9000
    Printer.Print "CANTIDAD"
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 200
    For Con = 1 To NRegistros
        If lvRequiFin.ListItems.Item(Con).Checked Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print lvRequiFin.ListItems.Item(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print lvRequiFin.ListItems(Con).SubItems(1)
            Printer.CurrentY = POSY
            Printer.CurrentX = 9000
            Printer.Print lvRequiFin.ListItems(Con).SubItems(2)
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 0
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            If POSY >= 14200 Then
                Printer.NewPage
                POSY = 100
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
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("REQUISICION #" & FolioR)) / 2
                Printer.Print "REQUISICION #" & FolioR
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 2200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9000
                Printer.Print "CANTIDAD"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            End If
        End If
    Next Con
    Printer.Print ""
    Printer.Print "FIN DEL LISTADO"
    Printer.EndDoc
    CommonDialog1.Copies = 1
Exit Sub
ManjaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Llenar_Lista_Proveedores
        Llenar_Lista_Requisiciones
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
Private Sub Text6_Change()
    Llenar_Lista_Proveedores
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Llenar_Lista_Requisiciones
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-_1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
