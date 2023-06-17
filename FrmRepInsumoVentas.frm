VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepInsumoVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ventas por Insumo"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   16
      Top             =   4920
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepInsumoVentas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepInsumoVentas.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   3
      Top             =   6120
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepInsumoVentas.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepInsumoVentas.frx":2156
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   1
      Top             =   7320
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepInsumoVentas.frx":26E5
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepInsumoVentas.frx":29EF
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepInsumoVentas.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton Command1 
         Caption         =   "Busca"
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
         Picture         =   "FrmRepInsumoVentas.frx":4AED
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5055
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8916
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4048
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
      Begin VB.Frame Frame1 
         Caption         =   "Fechas"
         Height          =   1455
         Left            =   6960
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   600
            TabIndex        =   12
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   40135
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   11
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   40135
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Busca"
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
         Left            =   7800
         Picture         =   "FrmRepInsumoVentas.frx":74BF
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Insumo :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   9960
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmRepInsumoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sIdProd As String
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    If sIdProd <> "" Then
        sBuscar = "SELECT ID_COMANDA, ID_PRODUCTO, INSUMO, CANTIDADVTA FROM VsRepJR WHERE INSUMO = '" & sIdProd & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    Else
        sBuscar = "SELECT ID_COMANDA, ID_PRODUCTO, INSUMO, CANTIDADVTA FROM VsRepJR WHERE FECHA_INICIO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("INSUMO")) Then tLi.SubItems(2) = tRs.Fields("INSUMO")
            If Not IsNull(tRs.Fields("CANTIDADVTA")) Then tLi.SubItems(3) = tRs.Fields("CANTIDADVTA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
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
        .ColumnHeaders.Add , , "ID", 1800
        .ColumnHeaders.Add , , "Descripcion", 5000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "COMANDA", 2200
        .ColumnHeaders.Add , , "PRODUCTO", 2200
        .ColumnHeaders.Add , , "INSUMO", 2200
        .ColumnHeaders.Add , , "CANTIDAD", 1440
    End With
End Sub
Private Sub Image26_Click()
    If ListView2.ListItems.Count > 0 Then
        FunImprime
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    sIdProd = Item
End Sub
Private Sub FunImprime()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\RepVentasInsumo.pdf") Then
        Exit Sub
    End If
    Posi = 145
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Courier_Bold, MacRomanEncoding
    ' Encabezado del reporte
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 85, 100, 450, "Insumos en Comandas", "F2", 14, hCenter
    oDoc.WTextBox 80, 400, 100, 200, "Periodo:", "F2", 10, hCenter
    oDoc.WTextBox 90, 400, 100, 200, DTPicker1.Value, "F2", 10, hCenter
    oDoc.WTextBox 100, 400, 100, 200, DTPicker2.Value, "F2", 10, hCenter
    oDoc.WTextBox 130, 85, 100, 175, "Comanda", "F2", 10, hLeft
    oDoc.WTextBox 130, 205, 100, 175, "Producto", "F2", 10, hLeft
    oDoc.WTextBox 130, 324, 100, 175, "Insumo", "F2", 10, hLeft
    oDoc.WTextBox 130, 428, 100, 175, "Cantidad", "F2", 10, hLeft
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    Posi = Posi + 6
    For Cont = 1 To ListView2.ListItems.Count
        oDoc.WTextBox Posi, 85, 20, 175, ListView2.ListItems(Cont), "F2", 8, hLeft
        oDoc.WTextBox Posi, 205, 20, 175, ListView2.ListItems(Cont).SubItems(1), "F2", 8, hLeft
        oDoc.WTextBox Posi, 324, 20, 175, ListView2.ListItems(Cont).SubItems(2), "F2", 8, hLeft
        oDoc.WTextBox Posi, 428, 20, 175, ListView2.ListItems(Cont).SubItems(3), "F2", 7, hLeft
        Posi = Posi + 12
        If Posi >= 760 Then
            oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
            ConPag = ConPag + 1
            Posi = 145
            oDoc.NewPage A4_Vertical
            oDoc.WImage 70, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 85, 100, 450, "Insumos en Comandas", "F2", 14, hCenter
            oDoc.WTextBox 80, 400, 100, 200, "Periodo:", "F2", 10, hCenter
            oDoc.WTextBox 90, 400, 100, 200, DTPicker1.Value, "F2", 10, hCenter
            oDoc.WTextBox 100, 400, 100, 200, DTPicker2.Value, "F2", 10, hCenter
            oDoc.WTextBox 130, 85, 100, 175, "Comanda", "F2", 8, hLeft
            oDoc.WTextBox 130, 205, 100, 175, "Producto", "F2", 8, hLeft
            oDoc.WTextBox 130, 324, 100, 175, "Insumo", "F2", 8, hLeft
            oDoc.WTextBox 130, 428, 100, 175, "Cantidad", "F2", 8, hLeft
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, Posi
            oDoc.WLineTo 580, Posi
            oDoc.LineStroke
            Posi = Posi + 6
        End If
    Next
    ' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 600
    oDoc.WLineTo 780, 600
    oDoc.LineStroke
    Posi = Posi + 6
    'cierre del reporte
    oDoc.PDFClose
    oDoc.Show
End Sub
