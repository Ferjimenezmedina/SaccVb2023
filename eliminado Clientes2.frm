VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EliCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar o Eliminar Cliente"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10320
      TabIndex        =   76
      Top             =   960
      Width           =   975
      Begin VB.Label Label39 
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
         TabIndex        =   77
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "eliminado Clientes2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Clientes2.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10320
      TabIndex        =   68
      Top             =   2160
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   73
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "eliminado Clientes2.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Clientes2.frx":1FD6
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label37 
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
            TabIndex        =   74
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   71
         Top             =   1320
         Width           =   975
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   0
            TabIndex        =   72
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "eliminado Clientes2.frx":3998
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Clientes2.frx":3CA2
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   69
         Top             =   0
         Width           =   975
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   0
            TabIndex        =   70
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "eliminado Clientes2.frx":54CC
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Clientes2.frx":57D6
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   75
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "eliminado Clientes2.frx":7288
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Clientes2.frx":7592
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10320
      TabIndex        =   66
      Top             =   3360
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
         TabIndex        =   67
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "eliminado Clientes2.frx":92BC
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Clientes2.frx":95C6
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
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
      Picture         =   "eliminado Clientes2.frx":B6A8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Te1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ID_CLIENTE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   4320
      TabIndex        =   32
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "eliminado Clientes2.frx":E07A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label26"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label43"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label44"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label45"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(18)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(15)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(12)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo8"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo7"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "eliminado Clientes2.frx":E096
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "COLONIA"
      Tab(1).Control(1)=   "Combo3"
      Tab(1).Control(2)=   "Text1(20)"
      Tab(1).Control(3)=   "Text1(19)"
      Tab(1).Control(4)=   "Text1(17)"
      Tab(1).Control(5)=   "Text1(14)"
      Tab(1).Control(6)=   "Text1(13)"
      Tab(1).Control(7)=   "Text1(11)"
      Tab(1).Control(8)=   "Text1(10)"
      Tab(1).Control(9)=   "Text1(9)"
      Tab(1).Control(10)=   "Text1(7)"
      Tab(1).Control(11)=   "Label36"
      Tab(1).Control(12)=   "Label35"
      Tab(1).Control(13)=   "Label34"
      Tab(1).Control(14)=   "Label33"
      Tab(1).Control(15)=   "Label32"
      Tab(1).Control(16)=   "Label31"
      Tab(1).Control(17)=   "Label30"
      Tab(1).Control(18)=   "Label29"
      Tab(1).Control(19)=   "Label11"
      Tab(1).Control(20)=   "Label8"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Credito"
      TabPicture(2)   =   "eliminado Clientes2.frx":E0B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text2"
      Tab(2).Control(1)=   "Combo5"
      Tab(2).Control(2)=   "Combo4"
      Tab(2).Control(3)=   "Check1"
      Tab(2).Control(4)=   "Combo2"
      Tab(2).Control(5)=   "Combo1"
      Tab(2).Control(6)=   "Text1(16)"
      Tab(2).Control(7)=   "Text1(23)"
      Tab(2).Control(8)=   "Label42"
      Tab(2).Control(9)=   "Label41"
      Tab(2).Control(10)=   "Label40"
      Tab(2).Control(11)=   "Label9"
      Tab(2).Control(12)=   "Label14"
      Tab(2).Control(13)=   "Label24"
      Tab(2).Control(14)=   "Comentarios"
      Tab(2).ControlCount=   15
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   4440
         TabIndex        =   86
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   3120
         TabIndex        =   85
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   120
         TabIndex        =   83
         Top             =   3600
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74760
         TabIndex        =   82
         Top             =   3720
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   -74760
         TabIndex        =   79
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   -74760
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar leyendas en Facturas"
         Height          =   255
         Left            =   -72600
         TabIndex        =   29
         Top             =   3480
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   30
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   120
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   120
         MaxLength       =   100
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   4080
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton COLONIA 
         Caption         =   "Colonia Nueva"
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
         Left            =   -72000
         Picture         =   "eliminado Clientes2.frx":E0CE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   -71040
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   22
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -72240
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -71520
         MaxLength       =   9
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -70320
         MaxLength       =   9
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -70440
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -72840
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74760
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74760
         TabIndex        =   25
         Text            =   "0"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   2085
         Index           =   23
         Left            =   -72840
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label45 
         Caption         =   "Régimen Capital"
         Height          =   255
         Left            =   4440
         TabIndex        =   88
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "Régimen Fiscal"
         Height          =   255
         Left            =   3120
         TabIndex        =   87
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "* Uso CFDi"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label42 
         Caption         =   "Cuenta Bancaria"
         Height          =   255
         Left            =   -74760
         TabIndex        =   81
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "Dia de Contra Recibo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   80
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label40 
         Caption         =   "Descuento por Tipo "
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74880
         TabIndex        =   65
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "* Estado"
         Height          =   195
         Left            =   -72840
         TabIndex        =   64
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   -72240
         TabIndex        =   63
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "* C.P."
         Height          =   195
         Left            =   -70440
         TabIndex        =   62
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Num Int"
         Height          =   195
         Left            =   -71520
         TabIndex        =   61
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Num Ext"
         Height          =   195
         Left            =   -70320
         TabIndex        =   60
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   -74880
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   -74880
         TabIndex        =   58
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Clave Cliente"
         Height          =   195
         Left            =   1800
         TabIndex        =   57
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre Comercial"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C"
         Height          =   195
         Left            =   4080
         TabIndex        =   54
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Tel. Casa"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tel. Trabajo"
         Height          =   195
         Left            =   1440
         TabIndex        =   52
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   2760
         TabIndex        =   51
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   1920
         TabIndex        =   50
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CURP"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña de Web"
         Height          =   195
         Left            =   4080
         TabIndex        =   48
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74640
         TabIndex        =   47
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -71880
         TabIndex        =   46
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Correo Electronico"
         Height          =   195
         Left            =   -71160
         TabIndex        =   45
         Top             =   2640
         Width           =   2250
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal"
         Height          =   195
         Left            =   -69240
         TabIndex        =   44
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Numero Interior"
         Height          =   195
         Left            =   -68760
         TabIndex        =   43
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Numero Exterior"
         Height          =   195
         Left            =   -70320
         TabIndex        =   42
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   -74640
         TabIndex        =   41
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   -74640
         TabIndex        =   40
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* Dirección"
         Height          =   195
         Left            =   -74880
         TabIndex        =   39
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   -71040
         TabIndex        =   38
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dias Crédito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   -74760
         TabIndex        =   36
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Limite de credito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Comentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   -72840
         TabIndex        =   34
         Top             =   960
         Width           =   870
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   2400
      TabIndex        =   30
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "EliCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub COLONIA_Click()
    FrmAgrColonia.Show vbModal
End Sub
Private Sub Combo1_DropDown()
    Combo1.Clear
    Combo1.AddItem " 0"
    Combo1.AddItem "15"
    Combo1.AddItem "30"
    Combo1.AddItem "45"
    Combo1.AddItem "60"
    Combo1.AddItem "90"
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_DropDown()
    Combo2.Clear
    Combo2.AddItem "50 %"
    Combo2.AddItem "40 %"
    Combo2.AddItem "30 %"
    Combo2.AddItem "25 %"
    Combo2.AddItem "20 %"
    Combo2.AddItem "15 %"
    Combo2.AddItem "14 %"
    Combo2.AddItem "11 %"
    Combo2.AddItem " 5 %"
    Combo2.AddItem " 0 %"
End Sub
Private Sub Combo3_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM COLONIAS ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo3.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then Combo3.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    Buscar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    On Error GoTo ManejaError
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
        .ColumnHeaders.Add , , "ID Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 4100, lvwColumnCenter
        .ColumnHeaders.Add , , "RFC", 0, lvwColumnCenter
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_DESCUENTO FROM DESCUENTOS GROUP BY ID_DESCUENTO ORDER BY ID_DESCUENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo4.AddItem tRs.Fields("ID_DESCUENTO")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT D1 FROM SEMANA  "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo5.AddItem tRs.Fields("D1")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT Descripcion FROM SATUsoCFDi ORDER BY Descripcion "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo6.AddItem tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT RegimenCapital FROM CLIENTE WHERE  RegimenCapital <> '' GROUP BY RegimenCapital ORDER BY RegimenCapital"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("RegimenCapital")) Then Combo7.AddItem tRs.Fields("RegimenCapital")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT Descripcion FROM SATRegimenFiscal ORDER BY Descripcion"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("Descripcion")) Then Combo8.AddItem tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clea
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE CLIENTE SET VALORACION = 'E' WHERE ID_CLIENTE = " & Te1.Text
    cnn.Execute (sBuscar)
    MsgBox "CLIENTE ELIMINADO!", vbInformation, "SACC"
    Te1.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim num As String
    Dim MosLey As String
    Dim ValDes As Double
    Dim UsoCFDi As String
    Dim RegimenFiscal As String
    If Text1(1).Text <> "" And Text1(15).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" And Combo3.Text <> "" And Text1(10).Text <> "" And Text1(11).Text <> "" And Text1(9).Text <> "" And Text1(7).Text <> "" And Combo6.Text <> "" Then
        If Text1(16).Text = "" Then
            Text1(16).Text = "0.00"
        End If
        If Check1.Value = 1 Then
            MosLey = "S"
        Else
            MosLey = "N"
        End If
        If Combo6.Text <> "" Then
            sBuscar = "SELECT Clave FROM SATUsoCFDi WHERE Descripcion = '" & Combo6.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    UsoCFDi = tRs.Fields("Clave")
                    tRs.MoveNext
                Loop
            End If
        End If
        If Combo8.Text <> "" Then
            sBuscar = "SELECT Clave FROM SATRegimenFiscal WHERE Descripcion = '" & Combo8.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                RegimenFiscal = tRs.Fields("Clave")
            Else
                MsgBox "NO SE ENCONTRÒ EL REGIMEN FISCAL SELECIONADO", vbExclamation, "SACC"
                Exit Sub
            End If
        Else
            MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTREO", vbExclamation, "SACC"
            Exit Sub
        End If
        Select Case Combo2.ListIndex
            Case Is = 0: ValDes = "50"
            Case Is = 1: ValDes = "40"
            Case Is = 2: ValDes = "30"
            Case Is = 3: ValDes = "25"
            Case Is = 4: ValDes = "20"
            Case Is = 5: ValDes = "15"
            Case Is = 6: ValDes = "14"
            Case Is = 7: ValDes = "11"
            Case Is = 8: ValDes = "05"
            Case Is = 9: ValDes = "0"
        End Select
        If ValDes = 0 Then
            ValDes = "0.00"
        End If
        ValDes = Replace(ValDes, ",", "")
        If VarMen.Text1(57).Text = "S" Then
            sBuscar = "UPDATE CLIENTE SET NOMBRE = '" & Text1(15).Text & "', NOMBRE_COMERCIAL = '" & Text1(1).Text & "', CURP = '" & Text1(12).Text & "', CONTACTO = '" & Text1(8).Text & "', RFC = '" & Text1(2).Text & "', TELEFONO_CASA = '" & Text1(3).Text & "', TELEFONO_TRABAJO = '" & Text1(4).Text & "', FAX = '" & Text1(5).Text & "', " _
                    & "WEB_PASSWORD = '" & Text1(18).Text & "', COLONIA = '" & Combo3.Text & "', CP = '" & Text1(10).Text & "', DIRECCION = '" & Text1(11).Text & "', NUMERO_EXTERIOR = '" & Text1(13).Text & "', NUMERO_INTERIOR = '" & Text1(14).Text & "', CIUDAD = '" & Text1(9).Text & "', ESTADO = '" & Text1(7).Text & "', PAIS = '" & Text1(20).Text & "', DIRECCION_ENVIO = '" & Text1(19).Text & "', " _
                    & "EMAIL = '" & Text1(17).Text & "', LIMITE_CREDITO = " & Text1(16).Text & ", DIAS_CREDITO = '" & Combo1.Text & "', DESCUENTO = " & ValDes & ", COMENTARIOS = '" & Text1(23).Text & "', LEYENDAS = '" & MosLey & "', ID_DESCUENTO = '" & Combo4.Text & "', RECIBO='" & Combo5.Text & " ', NUM_CUENTA_PAGO_CLIENTE = '" & Text2.Text & "', UsoCFDi = '" & UsoCFDi & "', RegimenCapital = '" & Combo7.Text & "', RegimenFiscal = '" & RegimenFiscal & "' WHERE ID_CLIENTE = " & Te1.Text & ""
        Else
            sBuscar = "UPDATE CLIENTE SET NOMBRE = '" & Text1(15).Text & "', NOMBRE_COMERCIAL = '" & Text1(1).Text & "', CURP = '" & Text1(12).Text & "', CONTACTO = '" & Text1(8).Text & "', RFC = '" & Text1(2).Text & "', TELEFONO_CASA = '" & Text1(3).Text & "', TELEFONO_TRABAJO = '" & Text1(4).Text & "', FAX = '" & Text1(5).Text & "', " _
                    & "WEB_PASSWORD = '" & Text1(18).Text & "', COLONIA = '" & Combo3.Text & "', CP = '" & Text1(10).Text & "', DIRECCION = '" & Text1(11).Text & "', NUMERO_EXTERIOR = '" & Text1(13).Text & "', NUMERO_INTERIOR = '" & Text1(14).Text & "', CIUDAD = '" & Text1(9).Text & "', ESTADO = '" & Text1(7).Text & "', PAIS = '" & Text1(20).Text & "', DIRECCION_ENVIO = '" & Text1(19).Text & "', " _
                    & "EMAIL = '" & Text1(17).Text & "', COMENTARIOS = '" & Text1(23).Text & "',RECIBO='" & Combo5.Text & "', UsoCFDi = '" & UsoCFDi & "', RegimenCapital = '" & Combo7.Text & "', RegimenFiscal = '" & RegimenFiscal & "' WHERE ID_CLIENTE = " & Te1.Text
        End If
        Set tRs = cnn.Execute(sBuscar)
        Text1(6).Text = ""
        Text1(15).Text = ""
        Text1(1).Text = ""
        Text1(12).Text = ""
        Text1(8).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Text1(18).Text = ""
        Combo3.Text = ""
        Text1(10).Text = ""
        Text1(11).Text = ""
        Text1(13).Text = ""
        Text1(14).Text = ""
        Text1(9).Text = ""
        Text1(7).Text = ""
        Text1(20).Text = ""
        Text1(19).Text = ""
        Text1(17).Text = ""
        Text1(16).Text = ""
        Combo1.Text = ""
        Combo2.Text = ""
        Combo4.Text = ""
        Combo5.Text = ""
        Combo7.Text = ""
        Combo8.Text = ""
        Text1(16).Text = ""
        Text1(23).Text = ""
    Else
        MsgBox "Debe proporcionar toda la información marcada con asteriscos (*)", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Te1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & Te1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text1(6).Text = Te1.Text
        If Not IsNull(tRs.Fields("NOMBRE")) Then
            Text1(15).Text = Trim(tRs.Fields("NOMBRE"))
        Else
            Text1(15).Text = ""
        End If
        If Not IsNull(tRs.Fields("NOMBRE_COMERCIAL")) Then
            Text1(1).Text = Trim(tRs.Fields("NOMBRE_COMERCIAL"))
        Else
            Text1(1).Text = ""
        End If
        If Not IsNull(tRs.Fields("CURP")) Then
            Text1(12).Text = Trim(tRs.Fields("CURP"))
        Else
            Text1(12).Text = ""
        End If
        If Not IsNull(tRs.Fields("CONTACTO")) Then
            Text1(8).Text = Trim(tRs.Fields("CONTACTO"))
        Else
            Text1(8).Text = ""
        End If
        If Not IsNull(tRs.Fields("RFC")) Then
            Text1(2).Text = Trim(tRs.Fields("RFC"))
        Else
            Text1(2).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO_CASA")) Then
            Text1(3).Text = Trim(tRs.Fields("TELEFONO_CASA"))
        Else
            Text1(3).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO_TRABAJO")) Then
            Text1(4).Text = Trim(tRs.Fields("TELEFONO_TRABAJO"))
        Else
            Text1(4).Text = ""
        End If
        If Not IsNull(tRs.Fields("FAX")) Then
            Text1(5).Text = Trim(tRs.Fields("FAX"))
        Else
            Text1(5).Text = ""
        End If
        If Not IsNull(tRs.Fields("WEB_PASSWORD")) Then
            Text1(18).Text = Trim(tRs.Fields("WEB_PASSWORD"))
        Else
            Text1(18).Text = ""
        End If
        If Not IsNull(tRs.Fields("COLONIA")) Then
            Combo3.Text = Trim(tRs.Fields("COLONIA"))
        Else
            Combo3.Text = ""
        End If
        If Not IsNull(tRs.Fields("CP")) Then
            Text1(10).Text = Trim(tRs.Fields("CP"))
        Else
            Text1(10).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIRECCION")) Then
            Text1(11).Text = Trim(tRs.Fields("DIRECCION"))
        Else
            Text1(11).Text = ""
        End If
        If Not IsNull(tRs.Fields("NUMERO_EXTERIOR")) Then
            Text1(13).Text = Trim(tRs.Fields("NUMERO_EXTERIOR"))
        Else
            Text1(13).Text = ""
        End If
        If Not IsNull(tRs.Fields("NUMERO_INTERIOR")) Then
            Text1(14).Text = Trim(tRs.Fields("NUMERO_INTERIOR"))
        Else
            Text1(14).Text = ""
        End If
        If Not IsNull(tRs.Fields("CIUDAD")) Then
            Text1(9).Text = Trim(tRs.Fields("CIUDAD"))
        Else
            Text1(9).Text = ""
        End If
        If Not IsNull(tRs.Fields("ESTADO")) Then
            Text1(7).Text = Trim(tRs.Fields("ESTADO"))
        Else
            Text1(7).Text = ""
        End If
        If Not IsNull(tRs.Fields("PAIS")) Then
            Text1(20).Text = Trim(tRs.Fields("PAIS"))
        Else
            Text1(20).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIRECCION_ENVIO")) Then
            Text1(19).Text = Trim(tRs.Fields("DIRECCION_ENVIO"))
        Else
            Text1(19).Text = ""
        End If
        If Not IsNull(tRs.Fields("EMAIL")) Then
            Text1(17).Text = Trim(tRs.Fields("EMAIL"))
        Else
            Text1(17).Text = ""
        End If
        If Not IsNull(tRs.Fields("LIMITE_CREDITO")) Then
            Text1(16).Text = Trim(tRs.Fields("LIMITE_CREDITO"))
        Else
            Text1(16).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIAS_CREDITO")) Then
            Combo1.Text = Trim(tRs.Fields("DIAS_CREDITO"))
        Else
            Combo1.Text = ""
        End If
        If Not IsNull(tRs.Fields("DESCUENTO")) Then
            Combo2.Text = Trim(tRs.Fields("DESCUENTO"))
        Else
            Combo2.Text = ""
        End If
        If Not IsNull(tRs.Fields("COMENTARIOS")) Then
            Text1(23).Text = Trim(tRs.Fields("COMENTARIOS"))
        Else
            Text1(23).Text = ""
        End If
        If Not IsNull(tRs.Fields("ID_DESCUENTO")) Then
            Combo4.Text = Trim(tRs.Fields("ID_DESCUENTO"))
        Else
            Combo4.Text = ""
        End If
        If Not IsNull(tRs.Fields("RECIBO")) Then
            Combo5.Text = Trim(tRs.Fields("RECIBO"))
        Else
            Combo5.Text = ""
        End If
        If Not IsNull(tRs.Fields("NUM_CUENTA_PAGO_CLIENTE")) Then
            Text2.Text = Trim(tRs.Fields("NUM_CUENTA_PAGO_CLIENTE"))
        Else
            Text2.Text = ""
        End If
        sBuscar = "SELECT Descripcion FROM SATUsoCFDi WHERE CLAVE = '" & tRs.Fields("UsoCFDi") & "'"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            Combo6.Text = tRs1.Fields("Descripcion")
        End If
        If Not IsNull(tRs.Fields("RegimenCapital")) Then
            Combo7.Text = tRs.Fields("RegimenCapital")
        Else
            Combo7.Text = ""
        End If
        sBuscar = "SELECT Descripcion FROM SATRegimenFiscal WHERE CLAVE = '" & tRs.Fields("RegimenFiscal") & "'"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            Combo8.Text = tRs1.Fields("Descripcion")
        End If
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Replace(Text5.Text, " ", "%")
    If IsNumeric(Text5.Text) Then
        sBuscar = "SELECT NOMBRE, RFC, ID_CLIENTE FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%' AND VALORACION NOT LIKE 'E' OR NOMBRE_COMERCIAL LIKE '%" & sBuscar & "%' AND VALORACION NOT LIKE 'E' OR ID_CLIENTE = '" & sBuscar & "' AND VALORACION NOT LIKE 'E'"
    Else
        sBuscar = "SELECT NOMBRE, RFC, ID_CLIENTE FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%' AND VALORACION NOT LIKE 'E' OR NOMBRE_COMERCIAL LIKE '%" & sBuscar & "%' AND VALORACION NOT LIKE 'E'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE") & "")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(2) = tRs.Fields("RFC")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 6 Then
        Text1(Index).BackColor = &HFFE1E1
    End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 16 Then
        Valido = "1234567890-()"
    Else
        If Index = 17 Then
            Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz _-@.,;&"
        Else
            Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%@!?*+"
        End If
    End If
    If Index = 17 Then
        KeyAscii = Asc(Chr(KeyAscii))
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    If Index <> 6 Then
        Text1(Index).BackColor = &H80000005
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890-ABCDEFGHIJKLMNÑOPQRSTUVWXYZ ()"
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
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
