VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEstadoProveedor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Estado con Proveedores"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   18
      Top             =   4680
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmEstadoProveedor.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmEstadoProveedor.frx":030A
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Busqueda"
      TabPicture(0)   =   "FrmEstadoProveedor.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmEstadoProveedor.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9128
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1200
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Numero de orden"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         Height          =   1215
         Left            =   5640
         TabIndex        =   8
         Top             =   480
         Width           =   1695
         Begin VB.OptionButton Option3 
            Caption         =   "Nacionales"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Internacionales"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Indirectas"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Todas"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   7440
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         Begin VB.CheckBox Check1 
            Caption         =   "Rando de fechas"
            Height          =   195
            Left            =   360
            TabIndex        =   3
            Top             =   960
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39955
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   720
            TabIndex        =   5
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39955
         End
         Begin VB.Label Label2 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBuscar2 
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
         Left            =   4440
         Picture         =   "FrmEstadoProveedor.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   9375
         _ExtentX        =   16536
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
      Begin VB.Label Label1 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmEstadoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub cmdBuscar2_Click()
    Busqueda
End Sub
Private Sub Form_Load()
' EN ORDEN_COMPRA EL CAMPO CONFIRMADA MARCA CON Y SI ESTA PAGADA Y CON X SI AUN NO SE PAGA
' EN LA TABLA ORDEN_COMPRA_DETALLE TIENE LOS CAMPOS DE CANTIDAD Y CANTIDAD_RECIBIDA
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Orden", 1000
        .ColumnHeaders.Add , , "Tipo", 1200
        .ColumnHeaders.Add , , "Proveedor", 4200
        .ColumnHeaders.Add , , "Total", 1850
        .ColumnHeaders.Add , , "Moneda", 1850
        .ColumnHeaders.Add , , "Fecha", 1850
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id. Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4200
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Surtido", 1200
        .ColumnHeaders.Add , , "Pendiente", 1200
        .ColumnHeaders.Add , , "Factura", 1200
        .ColumnHeaders.Add , , "No. Envio", 1200
    End With
    Busqueda
End Sub
Private Sub Busqueda()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value Then
        sBuscar = " WHERE NOMBRE LIKE '%" & Text1.Text & "%' "
    Else
       sBuscar = " WHERE NUM_ORDEN IN (" & Text1.Text & ") "
    End If
    If Option3.Value Then
        sBuscar = sBuscar & "AND TIPO = 'N' "
    Else
        If Option4.Value Then
            sBuscar = sBuscar & "AND TIPO = 'I' "
        Else
            If Option5.Value Then
                sBuscar = sBuscar & "AND TIPO = 'X' "
            End If
        End If
    End If
    If Check1.Value = 1 Then
        sBuscar = sBuscar & "AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' "
    End If
    sBuscar = "SELECT NUM_ORDEN, TIPO, NOMBRE, TOTAL, MONEDA, FECHA FROM VsRepComprasPagos " & sBuscar & " GROUP BY NUM_ORDEN, TIPO, NOMBRE, TOTAL, MONEDA, FECHA ORDER BY NUM_ORDEN"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(1) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(4) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT * FROM VsRepComprasPagos WHERE NUM_ORDEN = " & Item & " AND TIPO = '" & Item.SubItems(1) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("SURTIDO")) Then
                tLi.SubItems(3) = tRs.Fields("SURTIDO")
                tLi.SubItems(4) = CDbl(tRs.Fields("CANTIDAD")) - CDbl(tRs.Fields("SURTIDO"))
            Else
                tLi.SubItems(3) = "0"
                tLi.SubItems(4) = tRs.Fields("CANTIDAD")
            End If
            If Not IsNull(tRs.Fields("FACT_PROVE")) Then tLi.SubItems(5) = tRs.Fields("FACT_PROVE")
            If Not IsNull(tRs.Fields("NO_ENVIO")) Then tLi.SubItems(6) = tRs.Fields("NO_ENVIO")
            tRs.MoveNext
        Loop
    End If
End Sub
