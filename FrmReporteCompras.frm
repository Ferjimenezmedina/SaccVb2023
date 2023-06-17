VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteCompras 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta de compras pendientes de entrada"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      OLEDropMode     =   1
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmReporteCompras.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
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
         Left            =   7680
         Picture         =   "FrmReporteCompras.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Proveedor"
         Height          =   195
         Left            =   6240
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Producto"
         Height          =   195
         Left            =   6240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   5295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5530
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   6
      Top             =   6120
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmReporteCompras.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmReporteCompras.frx":2CF8
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmReporteCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Actualiza
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id", 0
        .ColumnHeaders.Add , , "Num. Orden", 1500
        .ColumnHeaders.Add , , "Tipo", 500
        .ColumnHeaders.Add , , "Proveedor", 3500
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1400
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Producto", 1500
        .ColumnHeaders.Add , , "Descripción", 500
        .ColumnHeaders.Add , , "Cantidad", 3500
        .ColumnHeaders.Add , , "C. Surtda", 1200
    End With
    Actualiza
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.fecha , ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA.ID_ORDEN_COMPRA IN (SELECT ID_ORDEN_COMPRA FROM ORDEN_COMPRA_DETALLE AS ORDEN_COMPRA_DETALLE_1 Where (CANTIDAD > Surtido) GROUP BY ID_ORDEN_COMPRA)) AND (ORDEN_COMPRA_DETALLE.ID_PRODUCTO LIKE '%" & Text1.Text & "%') GROUP BY ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, dbo.ORDEN_COMPRA.TOTAL"
    Else
        sBuscar = "SELECT ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA.ID_ORDEN_COMPRA IN (SELECT ID_ORDEN_COMPRA From ORDEN_COMPRA_DETALLE WHERE (CANTIDAD > Surtido) GROUP BY ID_ORDEN_COMPRA)) AND PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_COMPRA"))
            If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(2) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("Total")) Then tLi.SubItems(5) = tRs.Fields("Total")
            tRs.MoveNext
        Loop
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
    ListView2.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, SURTIDO From ORDEN_COMPRA_DETALLE WHERE (CANTIDAD > SURTIDO) AND ID_ORDEN_COMPRA = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion") & ""
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
                If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(3) = tRs.Fields("SURTIDO") & ""
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Actualiza
    End If
End Sub
