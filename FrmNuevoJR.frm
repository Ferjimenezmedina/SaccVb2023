VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmNuevoJR 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear Juego de Reparación"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame25 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   33
      Top             =   3960
      Width           =   975
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pegar"
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
      Begin VB.Image Image23 
         Height          =   630
         Left            =   120
         MouseIcon       =   "FrmNuevoJR.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmNuevoJR.frx":030A
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame24 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   31
      Top             =   3960
      Width           =   975
      Begin VB.Image Image22 
         Height          =   825
         Left            =   120
         MouseIcon       =   "FrmNuevoJR.frx":1D8C
         MousePointer    =   99  'Custom
         Picture         =   "FrmNuevoJR.frx":2096
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copiar"
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
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView LVCopia 
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   25
      Top             =   6360
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
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmNuevoJR.frx":425C
         MousePointer    =   99  'Custom
         Picture         =   "FrmNuevoJR.frx":4566
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   23
      Top             =   5160
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
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmNuevoJR.frx":6648
         MousePointer    =   99  'Custom
         Picture         =   "FrmNuevoJR.frx":6952
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Productos"
      TabPicture(0)   =   "FrmNuevoJR.frx":8314
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Agregados"
      TabPicture(1)   =   "FrmNuevoJR.frx":8330
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView4"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "ListView3"
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(4)=   "Label4"
      Tab(1).ControlCount=   5
      Begin MSComctlLib.ListView ListView4 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
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
         Left            =   -69960
         Picture         =   "FrmNuevoJR.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   20
         Top             =   3960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4683
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
      Begin VB.Frame Frame4 
         Caption         =   "Producto Seleccionado como Elemento del Juego de Reparación"
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   6120
         Width           =   6255
         Begin VB.CommandButton Command2 
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
            Left            =   4920
            Picture         =   "FrmNuevoJR.frx":AD1E
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   960
            TabIndex        =   19
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Producto Seleccionado como Juego de Reparación"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   6255
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4560
         TabIndex        =   10
         Top             =   3480
         Width           =   1575
         Begin VB.OptionButton Option4 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         Begin VB.OptionButton Option2 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   0
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3201
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2990
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
      Begin VB.Label Label6 
         Caption         =   "Nuevos Insumos Agregados..."
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
         TabIndex        =   29
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Insumos Capturados Anteriormente..."
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
         TabIndex        =   28
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmNuevoJR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim EliIt As Integer
Private Sub Command1_Click()
    If EliIt <> 0 Then
        ListView3.ListItems.Remove (EliIt)
        EliIt = 0
    End If
End Sub
Private Sub Command2_Click()
    If Text3.Text <> "" And Text4.Text <> "" And Text6.Text <> "" Then
        Dim tLi As ListItem
        Set tLi = ListView3.ListItems.Add(, , Text3.Text)
        tLi.SubItems(1) = Text4.Text
        tLi.SubItems(2) = Text6.Text
        Text4.Text = ""
        Text6.Text = ""
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Frame25.Visible = False
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
        .ColumnHeaders.Add , , "Clave del Producto", 2350
        .ColumnHeaders.Add , , "Descripcion", 5950
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2350
        .ColumnHeaders.Add , , "Descripcion", 5950
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 2350
        .ColumnHeaders.Add , , "Clave del Producto", 2350
        .ColumnHeaders.Add , , "Cantidad", 1450
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 2350
        .ColumnHeaders.Add , , "Clave del Producto", 2350
        .ColumnHeaders.Add , , "Cantidad", 1450
    End With
    With LVCopia
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 2350
        .ColumnHeaders.Add , , "Clave del Producto", 2350
        .ColumnHeaders.Add , , "Cantidad", 1450
    End With
End Sub
Private Sub Image22_Click()
    Dim Cont As Integer
    Dim tLi As ListItem
    LVCopia.ListItems.Clear
    For Cont = 1 To ListView4.ListItems.Count
        Set tLi = LVCopia.ListItems.Add(, , ListView4.ListItems(Cont))
        tLi.SubItems(1) = ListView4.ListItems(Cont).SubItems(1)
        tLi.SubItems(2) = ListView4.ListItems(Cont).SubItems(2)
        Frame25.Visible = True
        Frame24.Visible = False
    Next Cont
End Sub
Private Sub Image23_Click()
    If LVCopia.ListItems.Count > 0 Then
        Dim Cont As Integer
        Dim tLi As ListItem
        For Cont = 1 To LVCopia.ListItems.Count
            Set tLi = ListView3.ListItems.Add(, , Text3.Text)
                tLi.SubItems(1) = LVCopia.ListItems(Cont).SubItems(1)
                tLi.SubItems(2) = LVCopia.ListItems(Cont).SubItems(2)
                Frame24.Visible = True
                Frame25.Visible = False
        Next Cont
        LVCopia.ListItems.Clear
    Else
        MsgBox "NO HA COPIADO UN JUEGO DE REPARACIÓN!", vbInformation, "SACC"
    End If
End Sub
Private Sub Image8_Click()
    Dim NReg As Integer
    Dim Cont As Integer
    Dim sBuscar As String
    NReg = ListView3.ListItems.Count
    For Cont = 1 To NReg
        sBuscar = "INSERT INTO JUEGO_REPARACION (ID_REPARACION, ID_PRODUCTO, CANTIDAD) VALUES ('" & ListView3.ListItems(Cont) & "', '" & ListView3.ListItems(Cont).SubItems(1) & "', " & ListView3.ListItems(Cont).SubItems(2) & ");"
        cnn.Execute (sBuscar)
    Next Cont
    ListView3.ListItems.Clear
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Text3.Text = Item
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Item & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_REPARACION"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
    Text2.SetFocus
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text4.Text = Item
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EliIt = Item.Index
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tLi As ListItem
        Dim tRs As ADODB.Recordset
        ListView1.ListItems.Clear
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND TIPO IN ('COMPUESTO', 'EQUIVALE')"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1.Text & "%' AND TIPO IN ('COMPUESTO', 'EQUIVALE')"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
            ListView1.SetFocus
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tLi As ListItem
        Dim tRs As ADODB.Recordset
        ListView2.ListItems.Clear
        If Option3.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text2.Text & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
        If Option3.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text1.Text & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
        ListView2.SetFocus
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HFFE1E1
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &H80000005
End Sub
