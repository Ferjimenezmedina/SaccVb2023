VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerRastreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pedidos que componen la requisiscion "
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7200
   Icon            =   "frmVerRastreo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6480
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ListView lvDet 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmVerRastreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvDet
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "SOLICITO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FECHA", 1000
        .ColumnHeaders.Add , , "COMENTARIO", 4500
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLIENTE", 5000
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PROVEEDOR", 4000
        .ColumnHeaders.Add , , "ULTIMO PRECIO", 1500
        .ColumnHeaders.Add , , "FECHA", 1500
    End With
End Sub
Private Sub Timer1_Timer()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Timer1.Enabled = False
    sBuscar = "SELECT * FROM RASTREOREQUI WHERE ID_REQUI IN (" & Text1.Text & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = lvDet.ListItems.Add(, , tRs.Fields("SOLICITO") & "")
                tLi.SubItems(1) = tRs.Fields("CANTIDAD") & ""
                tLi.SubItems(3) = tRs.Fields("COMENTARIO") & ""
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE, ID_PRODUCTO, PRECIO_VENTA FROM VsLicitaCliProd WHERE ID_PRODUCTO = '" & Text2.Text & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NOMBRE") & "")
                If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(1) = tRs.Fields("PRECIO_VENTA") & ""
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Text2.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , "PRECIO DE VENTA PUBLICO (ALMACEN 3)")
                tLi.SubItems(1) = CDbl(tRs.Fields("PRECIO_COSTO")) * (CDbl(tRs.Fields("GANANCIA")) + 1)
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Text2.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , "PRECIO DE VENTA PUBLICO (ALMACEN 2)")
                tLi.SubItems(1) = CDbl(tRs.Fields("PRECIO_COSTO")) * (CDbl(tRs.Fields("GANANCIA")) + 1)
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Text2.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , "PRECIO DE VENTA PUBLICO (ALMACEN 1)")
                tLi.SubItems(1) = CDbl(tRs.Fields("PRECIO_COSTO")) * (CDbl(tRs.Fields("GANANCIA")) + 1)
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT TOP 5 NOMBRE, PRECIO, FECHA FROM VsComprasProveedor WHERE ID_PRODUCTO = '" & Text2.Text & "' ORDER BY FECHA DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("NOMBRE"))
                If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(1) = tRs.Fields("PRECIO")
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
End Sub
