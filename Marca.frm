VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Marca 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MARCAS COMERCIALES"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4800
      TabIndex        =   9
      Top             =   1320
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Marca.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Marca.frx":030A
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Marca.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "Marca.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar Marcas"
      TabPicture(0)   =   "Marca.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Nueva Marca"
      TabPicture(1)   =   "Marca.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtID"
      Tab(1).Control(1)=   "Text2"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74520
         TabIndex        =   2
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre Marca"
         Height          =   195
         Left            =   -74400
         TabIndex        =   5
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Marca"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1050
      End
   End
End
Attribute VB_Name = "Marca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "NOMBRE DE LA MARCA", 4000
    End With
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    If txtID.Text = "" Then
        sBuscar = "INSERT INTO MARCA(MARCA) VALUES('" & Text2.Text & "');"
    Else
        sBuscar = "UPDATE MARCA SET MARCA = '" & Text2.Text & "' WHERE ID_MARCA = " & txtID.Text
    End If
    cnn.Execute (sBuscar)
    Text2.Text = ""
    txtID.Text = ""
    ListView1.ListItems.Clear
    Text1.Text = ""
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.Count > 0 Then
        txtID.Text = Item
        Text2.Text = Item.SubItems(1)
        Text1.Text = Item.SubItems(1)
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        sBus = "SELECT * FROM MARCA WHERE MARCA LIKE '%" & Replace(Text1.Text, "*", "%") & "%'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_MARCA") & "")
                    If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(1) = .Fields("MARCA") & ""
                .MoveNext
            Loop
        End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. *"
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
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
