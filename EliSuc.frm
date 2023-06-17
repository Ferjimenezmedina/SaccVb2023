VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form EliSuc 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Eliminar Sucursal"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   33
      Top             =   360
      Width           =   975
      Begin VB.Label Label11 
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
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "EliSuc.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "EliSuc.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   18
      Top             =   1560
      Width           =   975
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   23
         Top             =   0
         Width           =   975
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "EliSuc.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "EliSuc.frx":1FD6
            Top             =   240
            Width           =   705
         End
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
            TabIndex        =   24
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   21
         Top             =   1320
         Width           =   975
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "EliSuc.frx":3A88
            MousePointer    =   99  'Custom
            Picture         =   "EliSuc.frx":3D92
            Top             =   240
            Width           =   675
         End
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
            TabIndex        =   22
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   19
         Top             =   1320
         Width           =   975
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
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "EliSuc.frx":55BC
            MousePointer    =   99  'Custom
            Picture         =   "EliSuc.frx":58C6
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "EliSuc.frx":7288
         MousePointer    =   99  'Custom
         Picture         =   "EliSuc.frx":7592
         Top             =   240
         Width           =   735
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Sucursal"
      TabPicture(0)   =   "EliSuc.frx":92BC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "EliSuc.frx":92D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(7)=   "Text3"
      Tab(1).Control(8)=   "Text4"
      Tab(1).Control(9)=   "Text6"
      Tab(1).Control(10)=   "Text7"
      Tab(1).Control(11)=   "Text8"
      Tab(1).Control(12)=   "Text9"
      Tab(1).Control(13)=   "Text10"
      Tab(1).ControlCount=   14
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -74160
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -69840
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74040
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -69960
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -74040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -71520
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -74040
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
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
         Left            =   6960
         Picture         =   "EliSuc.frx":92F4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         DataField       =   "NOMBRE"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3360
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataField       =   "ID_PROVEEDOR"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3360
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label10 
         Caption         =   "Folio :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Telefono : "
         Height          =   255
         Left            =   -70680
         TabIndex        =   31
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Ciudad :"
         Height          =   255
         Left            =   -70680
         TabIndex        =   29
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Colonia :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Calle :"
         Height          =   255
         Left            =   -72000
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Clave :"
         DataSource      =   "Adodc2"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   615
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   10
      Top             =   2760
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "EliSuc.frx":BCC6
         MousePointer    =   99  'Custom
         Picture         =   "EliSuc.frx":BFD0
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "EliSuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ItmIdSuc As String
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
        .ColumnHeaders.Add , , "ID Sucursal", 1000
        .ColumnHeaders.Add , , "Nombre", 2000, lvwColumnCenter
        .ColumnHeaders.Add , , "Calle", 2500, lvwColumnCenter
        .ColumnHeaders.Add , , "Colonia", 1500, lvwColumnCenter
        .ColumnHeaders.Add , , "Telefono", 1500, lvwColumnCenter
        .ColumnHeaders.Add , , "Ciudad", 0, lvwColumnCenter
        .ColumnHeaders.Add , , "Estado", 0, lvwColumnCenter
    End With
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    sBuscar = "DELETE FROM SUCURSALES WHERE ID_SUCURSAL = " & Text1.Text & " AND ELIMINADO = 'N'"
    cnn.Execute (sBuscar)
    MsgBox "SUCURSAL ELIMINADA!", vbInformation, "SACC"
    Text1.Text = ""
    Text2.Text = ""
    Text5.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    If ItmIdSuc <> "" Then
        Dim sBuscar As String
        sBuscar = "UPDATE SUCURSALES SET NOMBRE = '" & Text3.Text & "', CALLE = '" & Text4.Text & "', COLONIA = '" & Text6.Text & "', CIUDAD = '" & Text7.Text & "', ESTADO = '" & Text8.Text & "', TELEFONO = '" & Text9.Text & "' WHERE ID_SUCURSAL = " & ItmIdSuc
        cnn.Execute (sBuscar)
        Text3.Text = ""
        Text4.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(1)
    Text4.Text = Item.SubItems(2)
    Text6.Text = Item.SubItems(3)
    Text7.Text = Item.SubItems(5)
    Text8.Text = Item.SubItems(6)
    Text9.Text = Item.SubItems(4)
    ItmIdSuc = Item
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT NOMBRE, CALLE, ID_SUCURSAL, COLONIA, TELEFONO, CIUDAD, ESTADO FROM SUCURSALES WHERE NOMBRE LIKE '%" & Text5.Text & "%' AND ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_SUCURSAL") & "")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CALLE")) Then tLi.SubItems(2) = tRs.Fields("CALLE")
            If Not IsNull(tRs.Fields("COLONIA")) Then tLi.SubItems(3) = tRs.Fields("COLONIA")
            If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(4) = tRs.Fields("TELEFONO")
            If Not IsNull(tRs.Fields("CIUDAD")) Then tLi.SubItems(5) = tRs.Fields("CIUDAD")
            If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(6) = tRs.Fields("ESTADO")
            tRs.MoveNext
        Loop
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
