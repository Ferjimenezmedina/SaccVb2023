VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAltaProdAlm3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de Producto de Almacén 3"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   27
      Top             =   2040
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   55
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm3.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm3.frx":030A
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label24 
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
            TabIndex        =   56
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   53
         Top             =   1320
         Width           =   975
         Begin VB.Label Label23 
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
            TabIndex        =   54
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm3.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm3.frx":1FD6
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   51
         Top             =   0
         Width           =   975
         Begin VB.Label Label21 
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
            TabIndex        =   52
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm3.frx":3800
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm3.frx":3B0A
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label25 
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
         TabIndex        =   57
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm3.frx":55BC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm3.frx":58C6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   30
      Top             =   4440
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
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm3.frx":75F0
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm3.frx":78FA
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   29
      Top             =   3240
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm3.frx":99DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm3.frx":9CE6
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label15 
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
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmAltaProdAlm3.frx":B6A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "FrmAltaProdAlm3.frx":B6C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(1)=   "Text15"
      Tab(1).Control(2)=   "Text14"
      Tab(1).Control(3)=   "Text5"
      Tab(1).Control(4)=   "Text4"
      Tab(1).Control(5)=   "Combo7"
      Tab(1).Control(6)=   "Combo6"
      Tab(1).Control(7)=   "Combo5"
      Tab(1).Control(8)=   "Combo4"
      Tab(1).Control(9)=   "Combo3"
      Tab(1).Control(10)=   "Combo2"
      Tab(1).Control(11)=   "Text6"
      Tab(1).Control(12)=   "Combo1"
      Tab(1).Control(13)=   "Text3"
      Tab(1).Control(14)=   "Text7"
      Tab(1).Control(15)=   "Text8"
      Tab(1).Control(16)=   "Text9"
      Tab(1).Control(17)=   "Text10"
      Tab(1).Control(18)=   "Text11"
      Tab(1).Control(19)=   "Text12"
      Tab(1).Control(20)=   "Text13"
      Tab(1).Control(21)=   "cmdEnviar"
      Tab(1).Control(22)=   "Label31"
      Tab(1).Control(23)=   "Label30"
      Tab(1).Control(24)=   "Label29"
      Tab(1).Control(25)=   "Label28"
      Tab(1).Control(26)=   "Label26"
      Tab(1).Control(27)=   "Label22"
      Tab(1).Control(28)=   "Label20"
      Tab(1).Control(29)=   "Label19"
      Tab(1).Control(30)=   "Label14"
      Tab(1).Control(31)=   "Label18"
      Tab(1).Control(32)=   "Label17"
      Tab(1).Control(33)=   "Label16"
      Tab(1).Control(34)=   "Label3"
      Tab(1).Control(35)=   "Label5"
      Tab(1).Control(36)=   "Label6"
      Tab(1).Control(37)=   "Label7"
      Tab(1).Control(38)=   "Label8"
      Tab(1).Control(39)=   "Label9"
      Tab(1).Control(40)=   "Label10"
      Tab(1).Control(41)=   "Label11"
      Tab(1).Control(42)=   "Label12"
      Tab(1).Control(43)=   "Label13"
      Tab(1).ControlCount=   44
      Begin VB.CheckBox Check1 
         Caption         =   "Se puede hacer pedido manual de este producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   4800
         Width           =   5295
      End
      Begin VB.TextBox Text15 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -70560
         MaxLength       =   10
         TabIndex        =   22
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Text14 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -72240
         MaxLength       =   10
         TabIndex        =   21
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Text5 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -68880
         MaxLength       =   10
         TabIndex        =   23
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Text4 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -73800
         MaxLength       =   10
         TabIndex        =   20
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm3.frx":B6E0
         Left            =   -69960
         List            =   "FrmAltaProdAlm3.frx":B6E2
         TabIndex        =   19
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm3.frx":B6E4
         Left            =   -73920
         List            =   "FrmAltaProdAlm3.frx":B6E6
         TabIndex        =   18
         Top             =   3960
         Width           =   2655
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm3.frx":B6E8
         Left            =   -69720
         List            =   "FrmAltaProdAlm3.frx":B6F2
         TabIndex        =   17
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   -73680
         TabIndex        =   16
         Top             =   3480
         Width           =   3015
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm3.frx":B706
         Left            =   -74160
         List            =   "FrmAltaProdAlm3.frx":B708
         TabIndex        =   6
         Text            =   "SIMPLE"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -71640
         TabIndex        =   7
         Text            =   "S"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   -70200
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5400
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   5040
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73680
         MaxLength       =   500
         TabIndex        =   5
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox Text7 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -70320
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "<NINGUNO>"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -70560
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "<NINGUNO>"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -70440
         MaxLength       =   7
         TabIndex        =   15
         Text            =   "0"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "0"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text12 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -73800
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text13 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -72000
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Limpiar"
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
         Picture         =   "FrmAltaProdAlm3.frx":B70A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4920
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7011
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
      Begin VB.Label Label31 
         Caption         =   "% Imp. 1"
         Height          =   255
         Left            =   -71280
         TabIndex        =   62
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "% IVA"
         Height          =   255
         Left            =   -72840
         TabIndex        =   61
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "% Imp. 2"
         Height          =   255
         Left            =   -69600
         TabIndex        =   60
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "% Retención "
         Height          =   255
         Left            =   -74760
         TabIndex        =   59
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "* Presentación :"
         Height          =   255
         Left            =   -71160
         TabIndex        =   58
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69360
         TabIndex        =   50
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "* Categoria :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Precio en :"
         Height          =   255
         Left            =   -70560
         TabIndex        =   48
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "* Clasificación "
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "* Marca"
         Height          =   255
         Left            =   -70920
         TabIndex        =   45
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "* Tipo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "* Clave del Producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "* Descripción"
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "* Venta WEB"
         Height          =   255
         Left            =   -72720
         TabIndex        =   39
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   38
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "P. Venta"
         Height          =   255
         Left            =   -71040
         TabIndex        =   37
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "C. Maxima"
         Height          =   255
         Left            =   -71280
         TabIndex        =   36
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "C. Minima"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "% Ganancia"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "P. Compra"
         Height          =   255
         Left            =   -72840
         TabIndex        =   33
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Material"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Color"
         Height          =   255
         Left            =   -71160
         TabIndex        =   31
         Top             =   2520
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAltaProdAlm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProv As String
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Sub cmdEnviar_Click()
    IdProv = ""
    Text2.Text = ""
    Text3.Text = ""
    Text11.Text = ""
    Text10.Text = ""
    Combo3.Text = ""
    Combo2.Text = ""
    Combo1.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text7.Text = ""
    Text6.Text = ""
    Text12.Enabled = True
    Text13.Enabled = True
    Combo4.Text = ""
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim Resp As Long
    Resp = SendMessageLong(Combo1.hWnd, &H14F, True, 0)
    If KeyAscii = 13 Then
        Text12.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Combo2_DropDown()
    Combo2.Clear
    Combo2.AddItem "S"
    Combo2.AddItem "N"
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
    KeyAscii = 0
End Sub
Private Sub Combo3_DropDown()
    Combo3.Clear
    Combo3.AddItem "SIMPLE"
    Combo3.AddItem "COMPUESTO"
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
    KeyAscii = 0
End Sub

Private Sub Combo4_Change()
    If Combo4.Text <> "ORIGINAL" And Combo4.Text <> "SERVICIO" Then
        Combo6.Enabled = True
    End If
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo5.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo7.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Combo7_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Text4.Text = "0"
    Text14.Text = "16"
    Text15.Text = "0"
    Text5.Text = "0"
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Proveedor", 2500
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima", 1500
        .ColumnHeaders.Add , , "Tipo", 0
        .ColumnHeaders.Add , , "Venta WEB", 0
        .ColumnHeaders.Add , , "Marca", 0
        .ColumnHeaders.Add , , "Material", 0
        .ColumnHeaders.Add , , "Color", 0
        .ColumnHeaders.Add , , "Ganancia", 0
        .ColumnHeaders.Add , , "Precio de Costo", 0
        .ColumnHeaders.Add , , "Precio de Venta", 0
        .ColumnHeaders.Add , , "Clasificación", 0
        .ColumnHeaders.Add , , "Precio en", 0
        .ColumnHeaders.Add , , "Categoria", 100
        .ColumnHeaders.Add , , "Presentación", 100
        .ColumnHeaders.Add , , "% Retención", 100
        .ColumnHeaders.Add , , "% IVA", 100
        .ColumnHeaders.Add , , "% Impuesto 1", 100
        .ColumnHeaders.Add , , "% Impuesto 2", 100
        .ColumnHeaders.Add , , "Pedido Cliente", 0
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo4.Clear
    sBuscar = "SELECT CLASIFICACION FROM CLASIFICACIONES"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            Combo4.AddItem tRs.Fields("CLASIFICACION")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT CATEGORIA FROM ALMACEN3 GROUP BY CATEGORIA ORDER BY CATEGORIA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("CATEGORIA")) Then Combo6.AddItem tRs.Fields("CATEGORIA")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT PRESENTACION FROM ALMACEN3 GROUP BY PRESENTACION ORDER BY PRESENTACION "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("PRESENTACION")) Then Combo7.AddItem tRs.Fields("PRESENTACION")
            tRs.MoveNext
        Loop
    End If
    Combo1.Clear
    sBuscar = "SELECT MARCA FROM ALMACEN3 GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo1.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
    Text14.Text = VarMen.Text4(7).Text
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim MensajeSist As String
    sBuscar = "SELECT ID_REPARACION FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Text6.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        sBuscar = "DELETE FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_TEMPORALES WHERE ID_REPARACION = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO1 = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO2 = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        Text6.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text11.Text = "0"
        Text10.Text = "0"
        Combo3.Text = "SIMPLE"
        Combo2.Text = "S"
        Combo1.Text = ""
        Combo6.Text = ""
        Text8.Text = "<NINGUNO>"
        Text9.Text = "<NINGUNO>"
        Text12.Text = ""
        Text13.Text = ""
        Text7.Text = ""
        Combo6.Text = "ORIGINAL"
        Combo5.Text = ""
        Text4.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text5.Text = ""
    Else
        If MsgBox("EL PRODUCTO ES COMPUESTO ¿DESEA ELIMINAR EL PRODUCTO?, SE ELIMINARA TAMBIEN DEL JUEGO DE REPARACIÓN", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            sBuscar = "DELETE FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_TEMPORALES WHERE ID_REPARACION = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO1 = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO2 = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            Text6.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text11.Text = "0"
            Text10.Text = "0"
            Combo3.Text = "SIMPLE"
            Combo2.Text = "S"
            Combo1.Text = ""
            Combo6.Text = ""
            Text8.Text = "<NINGUNO>"
            Text9.Text = "<NINGUNO>"
            Text12.Text = ""
            Text13.Text = ""
            Text7.Text = ""
            Combo6.Text = "ORIGINAL"
            Combo5.Text = ""
        End If
    End If
    BusProd
End Sub
Private Sub Image8_Click()
    Dim sPedido As String
    If Check1.Value = 1 Then
        sPedido = "S"
    Else
        sPedido = "N"
    End If
    If Text7.Text <> "" Then
        If CDbl(Text7.Text) = 0 Then
            Text7.Text = ""
        End If
    End If
    If Text4.Text = "" Then
        Text4.Text = "0"
    End If
    If Text12.Text <> "" Then
        If CDbl(Text12.Text) = 0 Then
            Text12.Text = ""
        End If
    End If
    If Text13.Text <> "" Then
        If CDbl(Text13.Text) = 0 Then
            Text13.Text = ""
        End If
    End If
    If Text7.Text <> "" And Text13.Text <> "" And Text12.Text = "" Then
        Text12.Text = ((CDbl(Text7.Text) / CDbl(Text13.Text)) - 1) * 100
    End If
    If Text12.Text = "" Then
        Text12.Text = "20"
    End If
    If Not (Text7.Text = "" And Text13.Text = "") Then
        If Text7.Text = "" Then
            Text7.Text = ((CDbl(Text12.Text) / 100) + 1) * CDbl(Text13.Text)
        Else
            Text13.Text = CDbl(Text7.Text) / ((CDbl(Text12.Text) / 100) + 1)
        End If
    Else
        MsgBox "DEBE DAR UN PRECIO DE COSTO O PRECIO DE VENTA PARA EL REGISTRO!", vbInformation, "SACC"
        Exit Sub
    End If
    Text12.Text = CDbl(Text12.Text) / 100
    If Combo5.Text = "" Then
        Combo5.Text = "PESOS"
    End If
    If IsNumeric(Text4.Text) Then
        If Text4.Text = "" Then
            Text4.Text = "0"
        Else
            Text4.Text = CDbl(Text4.Text) / 100
        End If
    Else
        Text4.Text = "0"
    End If
    If IsNumeric(Text14.Text) Then
        If Text14.Text = "" Then
            Text14.Text = "0"
        Else
            Text14.Text = CDbl(Text14.Text) / 100
        End If
    Else
        Text4.Text = "0"
    End If
    If IsNumeric(Text15.Text) Then
        If Text15.Text = "" Then
            Text15.Text = "0"
        Else
            Text15.Text = CDbl(Text15.Text) / 100
        End If
    Else
        Text4.Text = "0"
    End If
    If IsNumeric(Text5.Text) Then
        If Text5.Text = "" Then
            Text5.Text = "0"
        Else
            Text5.Text = CDbl(Text5.Text) / 100
        End If
    Else
        Text4.Text = "0"
    End If
    Dim Ganan As String
    Dim PCompra As String
    If Combo6.Text <> "" And Combo7.Text <> "" And Text3.Text <> "" And Text6.Text <> "" And Text12.Text <> "" And Text13.Text <> "" And Text7.Text <> "" And Combo4.Text <> "" Then
        'If Text12.Text > 10 Then
        '    Text12.Text = Val(Text12.Text) / 100
        'End If
        Ganan = CDbl(Text12.Text)
        Text12.Text = Format(Text12.Text, "0.00000")
        Text13.Text = Format(Text13.Text, "0.00000")
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            sBuscar = "INSERT INTO ALMACEN3 (ID_PRODUCTO, DESCRIPCION, TIPO, VENTA_WEB, MARCA, GANANCIA, PRECIO_COSTO, MATERIAL, COLOR, C_MINIMA, C_MAXIMA, CLASIFICACION, PRECIO_EN, USR_ALTA, FECHA_ALTA, CATEGORIA, PRESENTACION, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2, PEDIDO_SUCURSAL) VALUES ('" & Text6.Text & "', '" & Text3.Text & "', '" & Combo3.Text & "', '" & Combo2.Text & "', '" & Combo1.Text & "', '" & Ganan & "', '" & Text13.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text11.Text & "', '" & Text10.Text & "', '" & Combo4.Text & "', '" & Combo5.Text & "', " & VarMen.Text1(0).Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Combo6.Text & "', '" & Combo7.Text & "', '" & Text4.Text & "', '" & Text14.Text & "', '" & Text15.Text & "', '" & Text5.Text & "', '" & sPedido & "');"
            cnn.Execute (sBuscar)
        Else
            If MsgBox("YA EXISTE UN PRODUCTO CON LA CLAVE " & Trim(Text6.Text) & "!, ¿DESEA GUARDAR LOS CAMBIOS?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                sBuscar = "UPDATE ALMACEN3 SET ID_PRODUCTO = '" & Text6.Text & "', Descripcion = '" & Text3.Text & "', TIPO = '" & Combo3.Text & "', VENTA_WEB = '" & Combo2.Text & "', MARCA = '" & Combo1.Text & "', GANANCIA = " & Ganan & ", PRECIO_COSTO = '" & Text13.Text & "', MATERIAL = '" & Text8.Text & "', COLOR = '" & Text9.Text & "', C_MINIMA = '" & Text11.Text & "', C_MAXIMA = " & Text10.Text & ", CLASIFICACION = '" & Combo4.Text & "', PRECIO_EN = '" & Combo5.Text & "', USR_MOD =" & VarMen.Text1(0).Text & ", FECHA_MOD = '" & Format(Date, "dd/mm/yyyy") & "', CATEGORIA = '" & Combo6.Text & "', PRESENTACION = '" & Combo7.Text & "', IVA = '" & Text14.Text & "', IMPUESTO1 = '" & Text15.Text & "', P_RETENCION = '" & Text5.Text & "', PEDIDO_SUCURSAL = '" & sPedido & "' WHERE ID_PRODUCTO = '" & IdProv & "'"
                Set tRs = cnn.Execute(sBuscar)
            End If
        End If
        IdProv = ""
        Text6.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text11.Text = "0"
        Text10.Text = "0"
        Combo3.Text = "SIMPLE"
        Combo2.Text = "S"
        Combo1.Text = ""
        Combo6.Text = ""
        Text8.Text = "<NINGUNO>"
        Text9.Text = "<NINGUNO>"
        Text12.Text = ""
        Text13.Text = ""
        Text7.Text = ""
        Combo6.Text = ""
        Combo5.Text = ""
        Combo7.Text = ""
        Text4.Text = "0"
        Text14.Text = "16"
        Text15.Text = "0"
        Text5.Text = "0"
        Check1.Value = 0
        Text12.Enabled = True
        Text13.Enabled = True
        If Text1.Text <> "" Then
            BusProd
        End If
    Else
        If Text6.Text <> "" Then
            MsgBox "ES NECESARIO DAR UNA CLAVE DE PRODUCTO!", vbInformation, "SACC"
        Else
            If Text3.Text <> "" Then
                MsgBox "ES NECESARIO DAR UNA Descripcion DEL PRODUCTO!", vbInformation, "SACC"
            End If
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub ListView1_DblClick()
    SSTab1.Tab = 1
    Text6.SetFocus
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Label22.Caption = ""
    Text2.Text = Trim(Item.SubItems(1))
    Text3.Text = Trim(Item.SubItems(1))
    Text11.Text = Trim(Item.SubItems(2))
    Text10.Text = Trim(Item.SubItems(3))
    Combo3.Text = Trim(Item.SubItems(4))
    Combo2.Text = Trim(Item.SubItems(5))
    Combo1.Text = Trim(Item.SubItems(6))
    Text8.Text = Trim(Item.SubItems(7))
    Text9.Text = Trim(Item.SubItems(8))
    Combo4.Text = Trim(Item.SubItems(12))
    Combo5.Text = Trim(Item.SubItems(13))
    Combo6.Text = Trim(Item.SubItems(14))
    Combo7.Text = Trim(Item.SubItems(15))
    Text4.Text = Trim(Item.SubItems(16))
    'If Combo3.Text <> "COMPUESTO" Then
        Text12.Text = CDbl(Item.SubItems(9)) * 100
        Text13.Text = Item.SubItems(10)
    'Else
    '    Text12.Enabled = False
    '    Text13.Enabled = False
    'End If
    Text7.Text = Item.SubItems(11)
    Text6.Text = Item
    If Item.SubItems(20) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Combo5.Text = "DOLARES" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT COMPRA FROM DOLAR ORDER BY ID_DOLAR DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Label22.Caption = Format(CDbl(Text7.Text) * CDbl(tRs.Fields("COMPRA")), "###,###,##0.00")
            Label22.Caption = "$ " & Label22.Caption & " M.N."
        Else
            MsgBox "NO SE TIENE PRECIO DEL DOLAR PARA CALCULAR EL PRECIO DE VENTA EN PESOS!", vbExclamation, "SACC"
        End If
    End If
    Text4.Text = CDbl(Item.SubItems(16)) * 100
    Text14.Text = CDbl(Item.SubItems(17)) * 100
    Text15.Text = CDbl(Item.SubItems(18)) * 100
    Text5.Text = CDbl(Item.SubItems(19)) * 100
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1.Text <> "" Then
        BusProd
    End If
End Sub
Private Sub BusProd()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT * FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1.Text & "%' ORDER BY Descripcion"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(2) = tRs.Fields("C_MINIMA")
            If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(3) = tRs.Fields("C_MAXIMA")
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(4) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("VENTA_WEB")) Then tLi.SubItems(5) = tRs.Fields("VENTA_WEB")
            If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(6) = tRs.Fields("MARCA")
            If Not IsNull(tRs.Fields("MATERIAL")) Then tLi.SubItems(7) = tRs.Fields("MATERIAL")
            If Not IsNull(tRs.Fields("COLOR")) Then tLi.SubItems(8) = tRs.Fields("COLOR")
            If Not IsNull(tRs.Fields("GANANCIA")) Then tLi.SubItems(9) = tRs.Fields("GANANCIA")
            If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(10) = tRs.Fields("PRECIO_COSTO")
            If Not IsNull(tRs.Fields("GANANCIA")) And Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(11) = Format((1 + CDbl(tRs.Fields("GANANCIA"))) * CDbl(tRs.Fields("PRECIO_COSTO")), "###,###,##0.00")
            If Not IsNull(tRs.Fields("CLASIFICACION")) Then tLi.SubItems(12) = tRs.Fields("CLASIFICACION")
            If Not IsNull(tRs.Fields("PRECIO_EN")) Then tLi.SubItems(13) = tRs.Fields("PRECIO_EN")
            If Not IsNull(tRs.Fields("CATEGORIA")) Then tLi.SubItems(14) = tRs.Fields("CATEGORIA")
            If Not IsNull(tRs.Fields("PRESENTACION")) Then tLi.SubItems(15) = tRs.Fields("PRESENTACION")
            If Not IsNull(tRs.Fields("P_RETENCION")) Then tLi.SubItems(16) = tRs.Fields("P_RETENCION")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(17) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("IMPUESTO1")) Then tLi.SubItems(18) = tRs.Fields("IMPUESTO1")
            If Not IsNull(tRs.Fields("IMPUESTO2")) Then tLi.SubItems(19) = tRs.Fields("IMPUESTO2")
            If Not IsNull(tRs.Fields("PEDIDO_SUCURSAL")) Then tLi.SubItems(20) = tRs.Fields("PEDIDO_SUCURSAL")
            tRs.MoveNext
        Loop
        ListView1.SetFocus
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text10_LostFocus()
    If Text10.Text = "" Then
        Text10.Text = "0"
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text10.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text11_LostFocus()
    If Text11.Text = "" Then
        Text11.Text = "0"
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text13.SetFocus
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
Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo3.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_LostFocus()
    If Text8.Text = "" Then
        Text8.Text = "<NINGUNO>"
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.- ?¿!()%&#$*/_+,1234567890" & """"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text9_LostFocus()
    If Text9.Text = "" Then
        Text9.Text = "<NINGUNO>"
    End If
End Sub
