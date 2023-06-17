VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmConsultaVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Ventas"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7680
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   10
      Top             =   3240
      Width           =   975
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmConsultaVentas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmConsultaVentas.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   8
      Top             =   4440
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmConsultaVentas.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmConsultaVentas.frx":2156
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmConsultaVentas.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
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
         Left            =   3840
         Picture         =   "FrmConsultaVentas.frx":4254
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por Cliente"
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Orden de compra"
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Factura"
         Height          =   255
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por  Nota de Venta"
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
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
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
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
Attribute VB_Name = "FrmConsultaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Venta", 1200
        .ColumnHeaders.Add , , "Factura", 1200
        .ColumnHeaders.Add , , "Orden", 1200
        .ColumnHeaders.Add , , "Nombre", 5700
        .ColumnHeaders.Add , , "Total", 1200
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Venta", 1200
        .ColumnHeaders.Add , , "Producto", 3200
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio", 1200
    End With
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
            Print #foo, StrCopi
            Close #foo
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Text1.Text
    Me.ListView2.ListItems.Clear
    sBuscar = "SELECT ID_VENTA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA FROM VENTAS_DETALLE WHERE ID_VENTA = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option1_Click()
    Text1.Text = ""
End Sub
Private Sub Option2_Click()
    Text1.Text = ""
End Sub
Private Sub Option3_Click()
    Text1.Text = ""
End Sub
Private Sub Option4_Click()
    Text1.Text = ""
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    If Option1.Value Then
        Valido = "1234567890"
    Else
        Valido = "1234567890. ABCDEFGHIJKLMNÑOPQRSTUVWXYZ#$%&/*"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Text1.Text
    Me.ListView1.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, NOOC, NOMBRE, TOTAL FROM VENTAS WHERE ID_VENTA LIKE '%" & sBuscar & "%' AND FACTURADO <> 2 ORDER BY ID_VENTA"
    End If
    If Option2.Value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, NOOC, NOMBRE, TOTAL  FROM VENTAS WHERE FOLIO LIKE '%" & sBuscar & "%' AND FACTURADO <> 2 ORDER BY ID_VENTA"
    End If
    If Option3.Value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, NOOC, NOMBRE, TOTAL  FROM VENTAS WHERE NOOC LIKE '%" & sBuscar & "%' AND FACTURADO <> 2 ORDER BY ID_VENTA"
    End If
    If Option4.Value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, NOOC, NOMBRE, TOTAL  FROM VENTAS WHERE NOMBRE LIKE '%" & sBuscar & "%' AND FACTURADO <> 2 ORDER BY ID_VENTA"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("NOOC")) Then tLi.SubItems(2) = tRs.Fields("NOOC")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
