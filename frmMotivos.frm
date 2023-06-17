VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMotivos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA O MODIFICACION DE MOTIVOS DE MATERIAL DAÑADO"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   13
      Top             =   1800
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
         MouseIcon       =   "frmMotivos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmMotivos.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   975
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmMotivos.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "frmMotivos.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "frmMotivos.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Informacion"
      TabPicture(1)   =   "frmMotivos.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "ID"
      Tab(1).Control(3)=   "Label3"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton Command1 
         Caption         =   "Insertar"
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
         Picture         =   "frmMotivos.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   855
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
         Left            =   5880
         Picture         =   "frmMotivos.frx":6AC2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73800
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73800
         TabIndex        =   9
         Top             =   1060
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
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
      Begin VB.Label ID 
         Caption         =   "ID:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMotivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As adodb.Connection
Private Sub Command1_Click()
Dim sBuscar As String
    If Text1.Text <> "" Then
        If Option2.Value = True Then
            sBuscar = "INSERT INTO MOTIVOS_SCRAP(MOTIVO) VALUES('" & Text1.Text & "');"
            cnn.Execute (sBuscar)
            Text1.Text = ""
        End If
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM MOTIVOS_SCRAP WHERE MOTIVO LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    With tRs
        If Not (.EOF And .BOF) Then
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID"))
            tLi.SubItems(1) = .Fields("MOTIVO")
        End If
    End With
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New adodb.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "Descripción", 5500
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
    frmScrap.Show vbModal
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.Count > 0 Then
        Text2.Text = Item.SubItems(1)
        Text3.Text = Item.SubItems(1)
        Text4.Text = Item
    End If
End Sub
Private Sub Option1_Click()
    Command2.Visible = True
    Command1.Visible = False
End Sub
Private Sub Option2_Click()
    Command1.Visible = True
    Command2.Visible = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    ListView1.ListItems.Clear
    If KeyAscii = 13 Then
        If Option1.Value = True Then
            Command2.Value = True
        End If
        If Option2.Value = True Then
            Command1.Value = True
        End If
    End If
End Sub
