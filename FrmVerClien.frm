VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVerClien 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Información de Cliente"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Te1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ID_CLIENTE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
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
      Picture         =   "FrmVerClien.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10440
      TabIndex        =   9
      Top             =   2880
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmVerClien.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerClien.frx":2CDC
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
         TabIndex        =   30
         Top             =   960
         Width           =   975
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   4320
      TabIndex        =   32
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "FrmVerClien.frx":4DBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(15)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(18)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "FrmVerClien.frx":4DDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(7)"
      Tab(1).Control(1)=   "Text1(9)"
      Tab(1).Control(2)=   "Text1(10)"
      Tab(1).Control(3)=   "Text1(11)"
      Tab(1).Control(4)=   "Text1(13)"
      Tab(1).Control(5)=   "Text1(14)"
      Tab(1).Control(6)=   "Text1(17)"
      Tab(1).Control(7)=   "Text1(19)"
      Tab(1).Control(8)=   "Text1(20)"
      Tab(1).Control(9)=   "Combo3"
      Tab(1).Control(10)=   "COLONIA"
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(13)=   "Label29"
      Tab(1).Control(14)=   "Label30"
      Tab(1).Control(15)=   "Label31"
      Tab(1).Control(16)=   "Label32"
      Tab(1).Control(17)=   "Label33"
      Tab(1).Control(18)=   "Label34"
      Tab(1).Control(19)=   "Label35"
      Tab(1).Control(20)=   "Label36"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Credito"
      TabPicture(2)   =   "FrmVerClien.frx":4DF6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(23)"
      Tab(2).Control(1)=   "Text1(16)"
      Tab(2).Control(2)=   "Combo1"
      Tab(2).Control(3)=   "Combo2"
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(5)=   "Combo4"
      Tab(2).Control(6)=   "Comentarios"
      Tab(2).Control(7)=   "Label24"
      Tab(2).Control(8)=   "Label14"
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(10)=   "Label40"
      Tab(2).ControlCount=   11
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   1965
         Index           =   23
         Left            =   -72840
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74760
         TabIndex        =   25
         Text            =   "0"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74760
         TabIndex        =   26
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -72840
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -70440
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -70320
         MaxLength       =   9
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   -71520
         MaxLength       =   9
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -72240
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   22
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   -71040
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
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
         Picture         =   "FrmVerClien.frx":4E12
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   4200
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   3
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   120
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   30
         TabIndex        =   8
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   33
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar leyendas en Facturas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72600
         TabIndex        =   29
         Top             =   3240
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74760
         TabIndex        =   27
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Comentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   -72840
         TabIndex        =   66
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Limite de credito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   65
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   -74760
         TabIndex        =   64
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dias Crédito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   63
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   -71040
         TabIndex        =   62
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* Dirección"
         Height          =   195
         Left            =   -74880
         TabIndex        =   61
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   -74640
         TabIndex        =   60
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   -74640
         TabIndex        =   59
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Numero Exterior"
         Height          =   195
         Left            =   -70320
         TabIndex        =   58
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Numero Interior"
         Height          =   195
         Left            =   -68760
         TabIndex        =   57
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal"
         Height          =   195
         Left            =   -69240
         TabIndex        =   56
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Correo Electronico"
         Height          =   195
         Left            =   -71160
         TabIndex        =   55
         Top             =   2640
         Width           =   2250
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -71880
         TabIndex        =   54
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74640
         TabIndex        =   53
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña de Web"
         Height          =   195
         Left            =   4080
         TabIndex        =   52
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CURP"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   450
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   2760
         TabIndex        =   49
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tel. Trabajo"
         Height          =   195
         Left            =   1440
         TabIndex        =   48
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Tel. Casa"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C"
         Height          =   195
         Left            =   4080
         TabIndex        =   46
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Comercial"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   1320
         TabIndex        =   44
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Clave Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   -74880
         TabIndex        =   42
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   -74880
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Num Ext"
         Height          =   195
         Left            =   -70320
         TabIndex        =   40
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Num Int"
         Height          =   195
         Left            =   -71520
         TabIndex        =   39
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "* C.P."
         Height          =   195
         Left            =   -70440
         TabIndex        =   38
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   -72240
         TabIndex        =   37
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "* Estado"
         Height          =   195
         Left            =   -72840
         TabIndex        =   36
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74880
         TabIndex        =   35
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label40 
         Caption         =   "Descuento por Tipo "
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   2400
      TabIndex        =   68
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmVerClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
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
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
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
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Te1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Replace(Text5.Text, " ", "%")
    sBuscar = "SELECT NOMBRE, RFC, ID_CLIENTE FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%'"
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
        Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
