VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRastreaPadido 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rastrear Pedidos"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   6
      Top             =   4200
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRastreaPadido.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRastreaPadido.frx":030A
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRastreaPadido.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
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
         Left            =   4800
         Picture         =   "FrmRastreaPadido.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar"
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   6015
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3120
            TabIndex        =   2
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50724865
            CurrentDate     =   39464
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50724865
            CurrentDate     =   39464
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label Label3 
            Caption         =   "y el "
            Height          =   255
            Left            =   2760
            TabIndex        =   11
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Pedido entre el"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Producto :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmRastreaPadido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM VsRastreaPedido WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("COMENTARIO")) Then tLi.SubItems(3) = tRs.Fields("COMENTARIO")
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(4) = tRs.Fields("SUCURSAL")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(6) = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.DTPicker1 = Format(Date - 15, "dd/mm/yyyy")
    Me.DTPicker2 = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "Id Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Notas", 5500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Pidio", 1500
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
