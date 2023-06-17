VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form RASTREAR 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rastreo de Pedidos"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   6
      Top             =   7320
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "RASTREAR.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "RASTREAR.frx":030A
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "PEDIDOS"
      TabPicture(0)   =   "RASTREAR.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
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
         Left            =   2640
         Picture         =   "RASTREAR.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   10080
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   4680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   8280
         TabIndex        =   16
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango"
         Height          =   1215
         Left            =   7800
         TabIndex        =   8
         Top             =   120
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39623
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39653
         End
         Begin VB.Label Label4 
            Caption         =   "Al"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "De:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   5040
         Width           =   9975
         _ExtentX        =   17595
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   9975
         _ExtentX        =   17595
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "SUCURSAL"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "RASTREAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.Clear
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT *  FROM VSRASTREO WHERE ENTREGADO <> 'I' AND CANTIDAD > 0 AND SUCURSAL='" & Combo1.Text & "'AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' ORDER BY FECHA DESC"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(6) = tRs.Fields("SUCURSAL")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_PEDIDO")) Then tLi.SubItems(7) = tRs.Fields("ID_PEDIDO")
            If Not IsNull(tRs.Fields("ENTREGADO")) Then
                If tRs.Fields("ENTREGADO") = "0" Then
                    tLi.SubItems(3) = "EN ESPERA"
                Else
                    If tRs.Fields("ENTREGADO") = "S" Then
                        tLi.SubItems(3) = "SURTIDO / EN CAMINO"
                    Else
                        tLi.SubItems(3) = "EN REQUISICION"
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
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
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Producto", 1500
        .ColumnHeaders.Add , , "Descripcion", 4500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Estado", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Uso", 1500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Id Pedido", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Estado", 4500
        .ColumnHeaders.Add , , "Fecha", 4500
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Label2.Caption = Item
    Text1.Text = Item
    Text2.Text = ListView1.SelectedItem.SubItems(2)
    Text4.Text = ListView1.SelectedItem.SubItems(4)
    ListView2.ListItems.Clear
    'SUM(CANTIDAD) AS
    sBuscar = "SELECT CANTIDAD, ACTIVO FROM REQUISICION WHERE ID_PRODUCTO ='" & Text1.Text & "' AND CANTIDAD='" & Text2.Text & "'AND FECHA='" & Text4.Text & "' GROUP BY ACTIVO, CANTIDAD"
    Set tRs = cnn.Execute(sBuscar)
     If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("CANTIDAD"))
            tLi.SubItems(2) = ListView1.SelectedItem.SubItems(4)
            If Not IsNull(tRs.Fields("ACTIVO")) Then
                If tRs.Fields("ACTIVO") = "0" Then
                    tLi.SubItems(1) = "EN ESPERA DE COTIZACIÓN"
                    Text3.Text = "EN ESPERA DE COTIZACIÓN"
                Else
                    If tRs.Fields("ACTIVO") = "1" Then
                        tLi.SubItems(1) = "EN ESPERA DE APROVACION DE COTIZACIÓN"
                         Text3.Text = "EN ESPERA DE APROVACION DE COTIZACIÓN"
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT  CANTIDAD, NUM_ORDEN, CONFIRMADA FROM VsOrdenCompra WHERE ID_PRODUCTO='" & Text1.Text & "' AND CANTIDAD='" & Text2.Text & "' AND FECHA='" & Text4.Text & "'  ORDER BY NUM_ORDEN"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("CANTIDAD"))
            If Not IsNull(tRs.Fields("CONFIRMADA")) Then
                If tRs.Fields("CONFIRMADA") = "N" Then
                    tLi.SubItems(1) = "EN PREORDEN DE COMPRA"
                    Text3.Text = "EN PREORDEN DE COMPRA"
                Else
                    If tRs.Fields("CONFIRMADA") = "P" Then
                        tLi.SubItems(1) = "PENDIENTE DE APROVACIÓN"
                        Text3.Text = "PENDIENTE DE APROVACIÓN"
                    Else
                        If tRs.Fields("CONFIRMADA") = "S" Then
                            tLi.SubItems(1) = "EN ESPERA DE PAGO A PROVEEDOR"
                            Text3.Text = "EN ESPERA DE PAGO A PROVEEDOR"
                        Else
                            If tRs.Fields("CONFIRMADA") = "X" Then
                                tLi.SubItems(1) = "PAGADO Y EN CAMINO"
                                Text3.Text = "PAGADO Y EN CAMINO"
                            End If
                        End If
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub BuscaPedidos()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, ENTREGADO, ID_PEDIDO FROM DETALLE_PEDIDO WHERE ENTREGADO <> 'I' AND CANTIDAD > 0"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("ENTREGADO")) Then
                If tRs.Fields("ENTREGADO") = "0" Then
                    tLi.SubItems(3) = "EN ESPERA"
                Else
                    If tRs.Fields("ENTREGADO") = "S" Then
                        tLi.SubItems(3) = "SURTIDO / EN CAMINO"
                    Else
                        tLi.SubItems(3) = "EN REQUISICION"
                    End If
                End If
            End If
            sBuscar = "SELECT FECHA, TIPO, SUCURSAL FROM PEDIDO WHERE ID_PEDIDO = " & tRs.Fields("ID_PEDIDO")
            Set tRs1 = cnn.Execute(sBuscar)
            If Not IsNull(tRs1.Fields("FECHA")) Then tLi.SubItems(4) = tRs1.Fields("FECHA")
             If Not IsNull(tRs1.Fields("TIPO")) Then
                If tRs1.Fields("TIPO") = "D" Then
                    tLi.SubItems(5) = "DE VENTA"
                Else
                    tLi.SubItems(5) = "DE SUCURSAL"
                End If
            End If
            If Not IsNull(tRs1.Fields("SUCURSAL")) Then tLi.SubItems(6) = tRs1.Fields("SUCURSAL")
            tRs.MoveNext
        Loop
    End If
End Sub
