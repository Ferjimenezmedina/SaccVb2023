VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAcepAlmacen1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisi�n de Material para Compra (Almacen 1)"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Folio"
      Height          =   1335
      Left            =   9240
      TabIndex        =   15
      Top             =   360
      Width           =   975
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
         Left            =   120
         Picture         =   "FrmAcepAlmacen1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   12
      Top             =   4920
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAcepAlmacen1.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmAcepAlmacen1.frx":2CDC
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   10
      Top             =   3720
      Width           =   975
      Begin VB.Label Label7 
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmAcepAlmacen1.frx":4DBE
         MousePointer    =   99  'Custom
         Picture         =   "FrmAcepAlmacen1.frx":50C8
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "FrmAcepAlmacen1.frx":6A8A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Aprobados"
      TabPicture(1)   =   "FrmAcepAlmacen1.frx":6AA6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command4"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Rechazados"
      TabPicture(2)   =   "FrmAcepAlmacen1.frx":6AC2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command6"
      Tab(2).Control(1)=   "ListView3"
      Tab(2).ControlCount=   2
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5280
         TabIndex        =   0
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5400
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
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
         Left            =   -67560
         Picture         =   "FrmAcepAlmacen1.frx":6ADE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8070
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
      Begin VB.CommandButton Command4 
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
         Left            =   -67440
         Picture         =   "FrmAcepAlmacen1.frx":94B0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5400
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
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
      Begin VB.CommandButton Command1 
         Caption         =   "Aprobar"
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
         Left            =   7320
         Picture         =   "FrmAcepAlmacen1.frx":BE82
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5400
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4575
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8070
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
      Begin VB.Label Label2 
         Caption         =   "Cantidad que funciono"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Id Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   5400
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAcepAlmacen1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim Itm As String
Dim SuItm1 As String
Dim SuItm2 As String
Dim SuItm3 As String
Dim SuItm4 As String
Dim SuItm5 As String
Dim SuItm6 As String
Dim SuItm7 As String
Dim SuItm8 As String
Dim SuItm9 As String
Dim Itm1 As String
Dim SuItm11 As String
Dim SuItm21 As String
Dim SuItm31 As String
Dim SuItm41 As String
Dim SuItm51 As String
Dim SuItm61 As String
Dim SuItm71 As String
Dim SuItm81 As String
Dim SuItm91 As String
Dim Itm2 As String
Dim SuItm12 As String
Dim SuItm22 As String
Dim SuItm32 As String
Dim SuItm42 As String
Dim SuItm52 As String
Dim SuItm62 As String
Dim SuItm72 As String
Dim SuItm82 As String
Dim SuItm92 As String
Dim ind As Integer
Dim InDQ1 As Integer
Dim InDQ2 As Integer
Private Sub Command1_Click()
Dim NReg As Integer
Dim Cont As Integer
Dim SUMADO As String
If Text2.Text <> "" Then
    'If CDbl(Text2.Text) >= CDbl(SuItm4) Then
        NReg = ListView2.ListItems.Count
        SUMADO = "N"
        For Cont = 1 To NReg
            If ListView2.ListItems.Item(Cont) = Itm Then
                ListView2.ListItems.Item(Cont).SubItems(4) = CDbl(Text2.Text) + CDbl(ListView2.ListItems.Item(Cont).SubItems(4))
                SUMADO = "S"
            End If
        Next Cont
        If ind > 0 Then
            If SUMADO = "N" Then
                Set tLi = ListView2.ListItems.Add(, , Itm & "")
                If Not IsNull(SuItm9) Then tLi.SubItems(1) = SuItm9
                If Not IsNull(SuItm1) Then tLi.SubItems(2) = SuItm1
                If Not IsNull(SuItm2) Then tLi.SubItems(3) = SuItm2
                If Not IsNull(SuItm3) Then tLi.SubItems(4) = Text2.Text
                If Not IsNull(SuItm4) Then tLi.SubItems(5) = SuItm4
                If Not IsNull(SuItm5) Then tLi.SubItems(6) = SuItm5
                If Not IsNull(SuItm6) Then tLi.SubItems(7) = SuItm6
                If Not IsNull(SuItm7) Then tLi.SubItems(8) = SuItm7
                If Not IsNull(SuItm8) Then tLi.SubItems(9) = SuItm8
            End If
            ListView1.ListItems.Remove (ind)
            ind = 0
            If (SuItm3 - Text2.Text) <> 0 Then
                Set tLi = ListView3.ListItems.Add(, , Itm & "")
                If Not IsNull(SuItm9) Then tLi.SubItems(1) = SuItm9
                If Not IsNull(SuItm1) Then tLi.SubItems(2) = SuItm1
                If Not IsNull(SuItm2) Then tLi.SubItems(3) = SuItm2
                If Not IsNull(SuItm3) Then tLi.SubItems(4) = (SuItm3 - Text2.Text)
                If Not IsNull(SuItm4) Then tLi.SubItems(5) = SuItm4
                If Not IsNull(SuItm5) Then tLi.SubItems(6) = SuItm5
                If Not IsNull(SuItm6) Then tLi.SubItems(7) = SuItm6
                If Not IsNull(SuItm7) Then tLi.SubItems(8) = SuItm7
                If Not IsNull(SuItm8) Then tLi.SubItems(9) = SuItm8
            End If
        End If
        Text1.Text = ""
        Text2.Text = ""
    'Else
    '    MsgBox "LA CANTIDAD ES MAYOR A LA CANTIDAD VALIDA!", vbInformation, "SACC"
    'End If
Else
    MsgBox "ES NECESARIO QUE DE UNA CANTIDAD!", vbInformation, "SACC"
End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM VsComprasAlm1reporte WHERE  GRUPO='" & Text3.Text & "' AND APROVADO = 'P' AND CAMPOALMACEN IN ('A1', 'A2') ORDER BY GRUPO ASC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_REVISION") & "")
                If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(1) = tRs.Fields("GRUPO")
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(3) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(5) = tRs.Fields("EXISTENCIA")
                If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(6) = tRs.Fields("C_MINIMA")
                If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(7) = tRs.Fields("C_MAXIMA")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(8) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then tLi.SubItems(9) = tRs.Fields("PRECIO_COMPRA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command4_Click()
    If InDQ1 > 0 Then
        Set tLi = ListView1.ListItems.Add(, , Itm & "")
        If Not IsNull(SuItm91) Then tLi.SubItems(1) = SuItm91
        If Not IsNull(SuItm11) Then tLi.SubItems(2) = SuItm11
        If Not IsNull(SuItm21) Then tLi.SubItems(3) = SuItm21
        If Not IsNull(SuItm31) Then tLi.SubItems(4) = SuItm31
        If Not IsNull(SuItm41) Then tLi.SubItems(5) = SuItm41
        If Not IsNull(SuItm51) Then tLi.SubItems(6) = SuItm51
        If Not IsNull(SuItm61) Then tLi.SubItems(7) = SuItm61
        If Not IsNull(SuItm71) Then tLi.SubItems(8) = SuItm71
        If Not IsNull(SuItm81) Then tLi.SubItems(9) = SuItm81
        ListView2.ListItems.Remove (InDQ1)
        InDQ1 = 0
    End If
End Sub
Private Sub Command6_Click()
    If InDQ2 > 0 Then
        Set tLi = ListView1.ListItems.Add(, , Itm & "")
        If Not IsNull(SuItm92) Then tLi.SubItems(1) = SuItm92
        If Not IsNull(SuItm12) Then tLi.SubItems(2) = SuItm12
        If Not IsNull(SuItm22) Then tLi.SubItems(3) = SuItm22
        If Not IsNull(SuItm32) Then tLi.SubItems(4) = SuItm32
        If Not IsNull(SuItm42) Then tLi.SubItems(5) = SuItm42
        If Not IsNull(SuItm52) Then tLi.SubItems(6) = SuItm52
        If Not IsNull(SuItm62) Then tLi.SubItems(7) = SuItm62
        If Not IsNull(SuItm72) Then tLi.SubItems(8) = SuItm72
        If Not IsNull(SuItm82) Then tLi.SubItems(9) = SuItm82
        ListView3.ListItems.Remove (InDQ2)
        InDQ2 = 0
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
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
        .ColumnHeaders.Add , , "Id Revision", 0
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Id Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Cant. Minima", 1000
        .ColumnHeaders.Add , , "Cant. Maxima", 1000
        .ColumnHeaders.Add , , "Proveedor", 3000
        .ColumnHeaders.Add , , "Precio Ofertado", 1000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Revision", 0
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Id Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Cant. Minima", 1000
        .ColumnHeaders.Add , , "Cant. Maxima", 1000
        .ColumnHeaders.Add , , "Proveedor", 3000
        .ColumnHeaders.Add , , "Precio Ofertado", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Revision", 0
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Id Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Cant. Minima", 1000
        .ColumnHeaders.Add , , "Cant. Maxima", 1000
        .ColumnHeaders.Add , , "Proveedor", 3000
        .ColumnHeaders.Add , , "Precio Ofertado", 1000
    End With
    sBuscar = "SELECT * FROM VsComprasAlm1reporte WHERE APROVADO = 'P' AND CAMPOALMACEN IN ('A1', 'A2') ORDER BY GRUPO ASC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_REVISION") & "")
                If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(1) = tRs.Fields("GRUPO")
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(3) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(5) = tRs.Fields("EXISTENCIA")
                If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(6) = tRs.Fields("C_MINIMA")
                If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(7) = tRs.Fields("C_MAXIMA")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(8) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then tLi.SubItems(9) = tRs.Fields("PRECIO_COMPRA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image5_Click()
    Dim sBuscar As String
    Dim Cont As Integer
    Dim NReg As Integer
    Dim tRs As ADODB.Recordset
    NReg = ListView2.ListItems.Count
    For Cont = 1 To NReg
        sBuscar = "UPDATE REV_COMPRA_ALMACEN1 SET CANTIDAD_APROVADA = " & ListView2.ListItems(Cont).SubItems(4) & ", APROVADO = 'R', FECHA_APROVADO = '" & Format(Date, "dd/mm/yyyy") & "' WHERE ID_REVISION = " & ListView2.ListItems(Cont)
        Set tRs = cnn.Execute(sBuscar)
    Next Cont
   ListView2.ListItems.Clear
   ListView3.ListItems.Clear
End Sub
Private Sub Image9_Click()
     Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Itm = Item
    SuItm9 = Item.SubItems(1)
    SuItm1 = Item.SubItems(2)
    SuItm2 = Item.SubItems(3)
    SuItm3 = Item.SubItems(4)
    SuItm4 = Item.SubItems(5)
    SuItm5 = Item.SubItems(6)
    SuItm6 = Item.SubItems(7)
    SuItm7 = Item.SubItems(8)
    SuItm8 = Item.SubItems(9)
    ind = Item.Index
    Text1.Text = Item.SubItems(2)
    Text2.Text = Item.SubItems(4)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Itm1 = Item
    SuItm91 = Item.SubItems(1)
    SuItm11 = Item.SubItems(2)
    SuItm21 = Item.SubItems(3)
    SuItm31 = Item.SubItems(4)
    SuItm41 = Item.SubItems(5)
    SuItm51 = Item.SubItems(6)
    SuItm61 = Item.SubItems(7)
    SuItm71 = Item.SubItems(8)
    SuItm81 = Item.SubItems(9)
    InDQ1 = Item.Index
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Itm2 = Item
    SuItm92 = Item.SubItems(1)
    SuItm12 = Item.SubItems(2)
    SuItm22 = Item.SubItems(3)
    SuItm32 = Item.SubItems(4)
    SuItm42 = Item.SubItems(5)
    SuItm52 = Item.SubItems(6)
    SuItm62 = Item.SubItems(7)
    SuItm72 = Item.SubItems(8)
    SuItm82 = Item.SubItems(9)
    InDQ2 = Item.Index
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Text3.Text <> "" Then
            Command2.Value = True
        End If
    End If
End Sub
