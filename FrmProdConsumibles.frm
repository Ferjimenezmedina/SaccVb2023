VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmProdConsumibles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos Consumibles"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   19
      Top             =   3480
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProdConsumibles.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdConsumibles.frx":030A
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   17
      Top             =   2280
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmProdConsumibles.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdConsumibles.frx":26F6
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmProdConsumibles.frx":40B8
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
      TabPicture(1)   =   "FrmProdConsumibles.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1"
      Tab(1).Control(1)=   "Text7"
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(3)=   "Text4"
      Tab(1).Control(4)=   "Text5"
      Tab(1).Control(5)=   "Text6"
      Tab(1).Control(6)=   "cmdEnviar"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label3"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(12)=   "Label6"
      Tab(1).ControlCount=   13
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmProdConsumibles.frx":40F0
         Left            =   -71880
         List            =   "FrmProdConsumibles.frx":40FD
         TabIndex        =   23
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -74400
         MaxLength       =   100
         TabIndex        =   21
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5040
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5040
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -74400
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -74400
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -70800
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   765
         Left            =   -74400
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2880
         Width           =   5775
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
         Left            =   -72000
         Picture         =   "FrmProdConsumibles.frx":4152
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
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
      Begin VB.Label Label8 
         Caption         =   "* Tipo :"
         Height          =   255
         Left            =   -71880
         TabIndex        =   24
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "* Cuenta contable :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "* Clave del Producto :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "* Descripción :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "* Precio :"
         Height          =   255
         Left            =   -70800
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Notas :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   11
         Top             =   2640
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmProdConsumibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProv As String
Private Sub cmdEnviar_Click()
    IdProv = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    IdProv = ""
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
        .ColumnHeaders.Add , , "Id Producto", 0
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Precio", 1500
        .ColumnHeaders.Add , , "Notas", 1500
        .ColumnHeaders.Add , , "Controla Existencias", 1500
        .ColumnHeaders.Add , , "Cuenta Contable", 1500
    End With
End Sub
Private Sub Image8_Click()
On Error GoTo ERRO
    If Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text7.Text <> "" Then
        Dim ContExis As String
        Text1.Text = Replace(Text1.Text, ",", "")
        Text2.Text = Replace(Text2.Text, ",", "")
        Text3.Text = Replace(Text3.Text, ",", "")
        Text4.Text = Replace(Text4.Text, ",", "")
        Text5.Text = Replace(Text5.Text, ",", "")
        Text6.Text = Replace(Text6.Text, ",", "")
        Dim sBuscar As String
        If Combo1.Text = "Controla existencias" Then
            ContExis = "S"
        Else
            If Combo1.Text = "No controla existencias" Then
                ContExis = "N"
            Else
                ContExis = "E"
            End If
        End If
        If IdProv <> "" Then
            sBuscar = "UPDATE PRODUCTOS_CONSUMIBLES SET ID_PRODUCTO = '" & Text3.Text & "', Descripcion = '" & Text4.Text & "', PRECIO = " & Text5.Text & ", NOTAS = '" & Text6.Text & "', CONTROLA_EXISTENCIA = '" & ContExis & "', CUENTA_CONTABLE = '" & Text7.Text & "' WHERE ID_PRODUCTO = '" & IdProv & "'"
            cnn.Execute (sBuscar)
        Else
            sBuscar = "INSERT INTO PRODUCTOS_CONSUMIBLES (ID_PRODUCTO, DESCRIPCION, PRECIO, NOTAS, CONTROLA_EXISTENCIA, CUENTA_CONTABLE) VALUES ('" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & ContExis & "', '" & Text7.Text & "');"
            cnn.Execute (sBuscar)
        End If
        IdProv = ""
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
    Else
        MsgBox "Falta información necesaria para el registro", vbExclamation, "SACC"
    End If
    Exit Sub
ERRO:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    IdProv = ""
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Text2.Text = Item
    Text3.Text = Item
    Text4.Text = Item.SubItems(1)
    Text5.Text = Item.SubItems(2)
    Text6.Text = Item.SubItems(3)
    If Item.SubItems(4) = "S" Then
        Combo1.Text = "Controla existencias"
    Else
        If Item.SubItems(4) = "N" Then
            Combo1.Text = "No controla existencias"
        Else
            Combo1.Text = "Entrada sin controlar existencia"
        End If
    End If
    Text7.Text = Item.SubItems(5)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT * FROM PRODUCTOS_CONSUMIBLES WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT * FROM PRODUCTOS_CONSUMIBLES WHERE Descripcion LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
        End If
        ListView1.ListItems.Clear
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                    If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(2) = tRs.Fields("PRECIO")
                    If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(3) = tRs.Fields("NOTAS")
                    If Not IsNull(tRs.Fields("CONTROLA_EXISTENCIA")) Then tLi.SubItems(4) = tRs.Fields("CONTROLA_EXISTENCIA")
                    If Not IsNull(tRs.Fields("CUENTA_CONTABLE")) Then tLi.SubItems(5) = tRs.Fields("CUENTA_CONTABLE")
                tRs.MoveNext
            Loop
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
