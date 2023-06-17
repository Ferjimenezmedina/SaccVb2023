VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepCartVac 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte Compra Proveedores Varios"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
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
      Left            =   7800
      Picture         =   "FrmRepCartVac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   12
      Top             =   4200
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepCartVac.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCartVac.frx":2CDC
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   9
      Top             =   5400
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepCartVac.frx":481E
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCartVac.frx":4B28
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepCartVac.frx":6C0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   4575
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Por Producto"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Por Proveedor"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16252929
         CurrentDate     =   42573
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16252929
         CurrentDate     =   42573
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Al :"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Del :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmRepCartVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private cnn As ADODB.Connection
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value Then
        sBuscar = "SELECT REV_COMPRA_ALMACEN1.ID_PRODUCTO, ALMACEN1.Descripcion, REV_COMPRA_ALMACEN1.CANTIDAD, PROVEEDOR_ALMACEN1.NOMBRE, REV_COMPRA_ALMACEN1.FECHA, REV_COMPRA_ALMACEN1.PRECIO_COMPRA, REV_COMPRA_ALMACEN1.Sucursal , REV_COMPRA_ALMACEN1.GRUPO FROM REV_COMPRA_ALMACEN1 INNER JOIN ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PRODUCTO = ALMACEN1.ID_PRODUCTO INNER JOIN PROVEEDOR_ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PROVEEDOR = PROVEEDOR_ALMACEN1.ID_PROVEEDOR WHERE REV_COMPRA_ALMACEN1.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND PROVEEDOR_ALMACEN1.NOMBRE LIKE '%" & Text1.Text & "%' UNION"
        sBuscar = sBuscar & " SELECT REV_COMPRA_ALMACEN1.ID_PRODUCTO, ALMACEN3.Descripcion, REV_COMPRA_ALMACEN1.CANTIDAD, PROVEEDOR_ALMACEN1.NOMBRE, REV_COMPRA_ALMACEN1.FECHA, REV_COMPRA_ALMACEN1.PRECIO_COMPRA, REV_COMPRA_ALMACEN1.Sucursal , REV_COMPRA_ALMACEN1.GRUPO FROM REV_COMPRA_ALMACEN1 INNER JOIN ALMACEN3 ON REV_COMPRA_ALMACEN1.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN PROVEEDOR_ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PROVEEDOR = PROVEEDOR_ALMACEN1.ID_PROVEEDOR WHERE REV_COMPRA_ALMACEN1.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND PROVEEDOR_ALMACEN1.NOMBRE LIKE '%" & Text1.Text & "%' "
    Else
        sBuscar = "SELECT REV_COMPRA_ALMACEN1.ID_PRODUCTO, ALMACEN1.Descripcion, REV_COMPRA_ALMACEN1.CANTIDAD, PROVEEDOR_ALMACEN1.NOMBRE, REV_COMPRA_ALMACEN1.FECHA, REV_COMPRA_ALMACEN1.PRECIO_COMPRA, REV_COMPRA_ALMACEN1.Sucursal , REV_COMPRA_ALMACEN1.GRUPO FROM REV_COMPRA_ALMACEN1 INNER JOIN ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PRODUCTO = ALMACEN1.ID_PRODUCTO INNER JOIN PROVEEDOR_ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PROVEEDOR = PROVEEDOR_ALMACEN1.ID_PROVEEDOR WHERE REV_COMPRA_ALMACEN1.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND REV_COMPRA_ALMACEN1.ID_PRODUCTO LIKE '%" & Text1.Text & "%' UNION"
        sBuscar = sBuscar & " SELECT REV_COMPRA_ALMACEN1.ID_PRODUCTO, ALMACEN3.Descripcion, REV_COMPRA_ALMACEN1.CANTIDAD, PROVEEDOR_ALMACEN1.NOMBRE, REV_COMPRA_ALMACEN1.FECHA, REV_COMPRA_ALMACEN1.PRECIO_COMPRA, REV_COMPRA_ALMACEN1.Sucursal , REV_COMPRA_ALMACEN1.GRUPO FROM REV_COMPRA_ALMACEN1 INNER JOIN ALMACEN3 ON REV_COMPRA_ALMACEN1.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN PROVEEDOR_ALMACEN1 ON REV_COMPRA_ALMACEN1.ID_PROVEEDOR = PROVEEDOR_ALMACEN1.ID_PROVEEDOR WHERE REV_COMPRA_ALMACEN1.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND REV_COMPRA_ALMACEN1.ID_PRODUCTO LIKE '%" & Text1.Text & "%' "
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then tLi.SubItems(5) = Format(tRs.Fields("PRECIO_COMPRA"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("Sucursal")) Then tLi.SubItems(6) = tRs.Fields("Sucursal")
            If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(7) = tRs.Fields("GRUPO")
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
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Precio de Compra", 1500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Folio", 1500
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
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
