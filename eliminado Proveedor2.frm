VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form EliProveedor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Proveedor"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   59
      Top             =   720
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   60
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "eliminado Proveedor2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Proveedor2.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   51
      Top             =   1920
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   56
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "eliminado Proveedor2.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Proveedor2.frx":1FD6
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
            TabIndex        =   57
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   54
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
            TabIndex        =   55
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "eliminado Proveedor2.frx":3998
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Proveedor2.frx":3CA2
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   52
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
            TabIndex        =   53
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "eliminado Proveedor2.frx":54CC
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Proveedor2.frx":57D6
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
         TabIndex        =   58
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "eliminado Proveedor2.frx":7288
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Proveedor2.frx":7592
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   49
      Top             =   3120
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
         TabIndex        =   50
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "eliminado Proveedor2.frx":92BC
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Proveedor2.frx":95C6
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
      Picture         =   "eliminado Proveedor2.frx":B6A8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
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
   Begin VB.TextBox Te1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ID_PROVEEDOR"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   4320
      TabIndex        =   25
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "eliminado Proveedor2.frx":E07A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label21"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label28"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(6)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(19)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Bancos"
      TabPicture(1)   =   "eliminado Proveedor2.frx":E096
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Text1(12)"
      Tab(1).Control(7)=   "Text1(13)"
      Tab(1).Control(8)=   "Text1(14)"
      Tab(1).Control(9)=   "Text1(15)"
      Tab(1).Control(10)=   "Text1(16)"
      Tab(1).Control(11)=   "Text1(17)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Notas"
      TabPicture(2)   =   "eliminado Proveedor2.frx":E0B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Text1(18)"
      Tab(2).Control(2)=   "Label7"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   100
         TabIndex        =   61
         Top             =   3720
         Width           =   5895
      End
      Begin VB.Frame Frame1 
         Caption         =   "* Almacen"
         Height          =   2535
         Left            =   -70800
         TabIndex        =   48
         Top             =   960
         Width           =   1815
         Begin VB.CheckBox Check3 
            Caption         =   "Almacen 3"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Almacen 2"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Almacen 1"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   4
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   120
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   27
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3960
         TabIndex        =   26
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   120
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -70800
         MaxLength       =   100
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   18
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   15
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -72960
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   2325
         Index           =   18
         Left            =   -74760
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "* E-mail"
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre"
         Height          =   195
         Left            =   1920
         TabIndex        =   46
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* Direccion"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   2280
         TabIndex        =   43
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "* Estado"
         Height          =   195
         Left            =   4320
         TabIndex        =   42
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* C.P."
         Height          =   195
         Left            =   2520
         TabIndex        =   40
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C."
         Height          =   195
         Left            =   4080
         TabIndex        =   39
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Telefono 1"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Telefono 2"
         Height          =   195
         Left            =   2280
         TabIndex        =   37
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   4320
         TabIndex        =   36
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Clave Swift"
         Height          =   195
         Left            =   -70680
         TabIndex        =   35
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   -74880
         TabIndex        =   34
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Routing"
         Height          =   195
         Left            =   -72600
         TabIndex        =   33
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   -74760
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   -72840
         TabIndex        =   31
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   -74760
         TabIndex        =   30
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Notas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   4080
      Width           =   615
   End
End
Attribute VB_Name = "EliProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub cmdCancelar_Click()
    Unload Me
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
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 4200, lvwColumnCenter
        .ColumnHeaders.Add , , "RFC", 0, lvwColumnCenter
        .ColumnHeaders.Add , , "DIRECCION", 0, lvwColumnCenter
    End With
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    sBuscar = "DELETE FROM PROVEEDOR WHERE ID_PROVEEDOR = " & Te1.Text
    cnn.Execute (sBuscar)
    MsgBox "PROVEEDOR ELIMINADO!", vbInformation, "SACC"
    Te1.Text = ""
    Text5.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    If Te1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim num As String
        Dim ALM1 As String
        Dim ALM2 As String
        Dim ALM3 As String
        If Check1.value = 1 Then
            ALM1 = "S"
        Else
            ALM1 = "N"
        End If
        If Check2.value = 1 Then
            ALM2 = "S"
        Else
            ALM2 = "N"
        End If
        If Check3.value = 1 Then
            ALM3 = "S"
        Else
            ALM3 = "N"
        End If
        sBuscar = "UPDATE PROVEEDOR SET NOMBRE = '" & Text1(2).Text & "', DIRECCION = '" & Text1(19).Text & "', COLONIA = '" & Text1(3).Text & "', CIUDAD = '" & Text1(4).Text & "', ESTADO = '" & Text1(5).Text & "', PAIS = '" & Text1(6).Text & "', CP = '" & Text1(7).Text & "', RFC = '" & Text1(8).Text & "', TELEFONO1 = '" & Text1(9).Text & "', TELEFONO2 = '" & Text1(10).Text & "', TELEFONO3 = '" & Text1(11).Text & "', TRANS_BANCO = '" & Text1(12).Text & "', TRANS_DIRECCION = '" & Text1(13).Text & "', TRANS_CIUDAD = '" & Text1(14).Text & "', TRANS_ROUTING = '" & Text1(15).Text & "', TRANS_CLAVE_SWIFT = '" & Text1(17).Text & "', TRANS_CUENTA = '" & Text1(16).Text & "', NOTAS = '" & Text1(18).Text & "', ALMACEN1 = '" & ALM1 & "', ALMACEN2 = '" & ALM2 & "', ALMACEN3 = '" & ALM3 & "', EMAIL = '" & Text1(0).Text & "' WHERE ID_PROVEEDOR = " & Te1.Text
        Set tRs = cnn.Execute(sBuscar)
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(19).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Text1(6).Text = ""
        Text1(7).Text = ""
        Text1(8).Text = ""
        Text1(9).Text = ""
        Text1(10).Text = ""
        Text1(11).Text = ""
        Text1(12).Text = ""
        Text1(13).Text = ""
        Text1(14).Text = ""
        Text1(15).Text = ""
        Text1(17).Text = ""
        Text1(16).Text = ""
        Text1(18).Text = ""
        Te1.Text = ""
    Else
        MsgBox "DEBE SELECCIONAR UN PROVEEDOR PARA MODIFICAR!", vbModal, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Te1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & Te1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then
            Text1(1).Text = tRs.Fields("ID_PROVEEDOR")
        Else
            Text1(1).Text = ""
        End If
        If Not IsNull(tRs.Fields("NOMBRE")) Then
            Text1(2).Text = tRs.Fields("NOMBRE")
        Else
            Text1(2).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIRECCION")) Then
            Text1(19).Text = tRs.Fields("DIRECCION")
        Else
            Text1(19).Text = ""
        End If
        If Not IsNull(tRs.Fields("COLONIA")) Then
            Text1(3).Text = tRs.Fields("COLONIA")
        Else
            Text1(3).Text = ""
        End If
        If Not IsNull(tRs.Fields("CIUDAD")) Then
            Text1(4).Text = tRs.Fields("CIUDAD")
        Else
            Text1(4).Text = ""
        End If
        If Not IsNull(tRs.Fields("ESTADO")) Then
            Text1(5).Text = tRs.Fields("ESTADO")
        Else
            Text1(5).Text = ""
        End If
        If Not IsNull(tRs.Fields("PAIS")) Then
            Text1(6).Text = tRs.Fields("PAIS")
        Else
            Text1(6).Text = ""
        End If
        If Not IsNull(tRs.Fields("CP")) Then
            Text1(7).Text = tRs.Fields("CP")
        Else
            Text1(7).Text = ""
        End If
        If Not IsNull(tRs.Fields("RFC")) Then
            Text1(8).Text = tRs.Fields("RFC")
        Else
            Text1(8).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO1")) Then
            Text1(9).Text = tRs.Fields("TELEFONO1")
        Else
            Text1(9).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO2")) Then
            Text1(10).Text = tRs.Fields("TELEFONO2")
        Else
            Text1(10).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO3")) Then
            Text1(11).Text = tRs.Fields("TELEFONO3")
        Else
            Text1(11).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_BANCO")) Then
            Text1(12).Text = tRs.Fields("TRANS_BANCO")
        Else
            Text1(12).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_DIRECCION")) Then
            Text1(13).Text = tRs.Fields("TRANS_DIRECCION")
        Else
            Text1(13).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_CIUDAD")) Then
            Text1(14).Text = tRs.Fields("TRANS_CIUDAD")
        Else
            Text1(14).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_ROUTING")) Then
            Text1(15).Text = tRs.Fields("TRANS_ROUTING")
        Else
            Text1(15).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_CLAVE_SWIFT")) Then
            Text1(17).Text = tRs.Fields("TRANS_CLAVE_SWIFT")
        Else
            Text1(17).Text = ""
        End If
        If Not IsNull(tRs.Fields("TRANS_CUENTA")) Then
            Text1(16).Text = tRs.Fields("TRANS_CUENTA")
        Else
            Text1(16).Text = ""
        End If
        If Not IsNull(tRs.Fields("NOTAS")) Then
            Text1(18).Text = tRs.Fields("NOTAS")
        Else
            Text1(18).Text = ""
        End If
        If Not IsNull(tRs.Fields("EMAIL")) Then
            Text1(0).Text = tRs.Fields("EMAIL")
        Else
            Text1(0).Text = ""
        End If
        If tRs.Fields("ALMACEN1") = "S" Then
            Check1.value = 1
        Else
            Check1.value = 0
        End If
        If tRs.Fields("ALMACEN2") = "S" Then
            Check2.value = 1
        Else
            Check2.value = 0
        End If
        If tRs.Fields("ALMACEN3") = "S" Then
            Check3.value = 1
        Else
            Check3.value = 0
        End If
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT NOMBRE, RFC, ID_PROVEEDOR, DIRECCION FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text5.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR") & "")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(2) = tRs.Fields("RFC")
            If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(3) = tRs.Fields("DIRECCION")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 1 Then
        Text1(Index).BackColor = &HFFE1E1
    End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    If Index = 7 Or Index = 9 Or Index = 10 Or Index = 11 Then
        Valido = "1234567890-()"
    Else
        If Index = 18 Then
            Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz -()/&%@!?*+"
        Else
            If Index = 0 Then
                Valido = "1234567890.abcdefghijklmnñopqrstuvwxyz@-_"
            Else
                Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
            End If
        End If
    End If
    If Index = 18 Or Index = 0 Then
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
    If Index <> 1 Then
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
