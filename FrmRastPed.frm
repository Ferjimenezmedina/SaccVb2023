VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRastPed 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rastrear Pedidos en Proceso de Compras o Almacen"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   8
      Top             =   5880
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRastPed.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRastPed.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Rastrear"
      TabPicture(0)   =   "FrmRastPed.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6840
         TabIndex        =   0
         Text            =   "<TODAS>"
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   9015
         _ExtentX        =   15901
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
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   4440
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   1080
         TabIndex        =   6
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal  :"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Productos Pedidos"
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmRastPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Combo1_Click()
    Dim NoReg As Integer
    Dim Con As Integer
    BuscaPedidos
    Con = 1
    NoReg = ListView1.ListItems.Count
    If Combo1.Text <> "<TODAS>" Then
        Do While Con <= NoReg
            If ListView1.ListItems(Con).SubItems(6) <> Combo1.Text Then
                ListView1.ListItems.Remove (Con)
                NoReg = ListView1.ListItems.Count
            Else
                Con = Con + 1
            End If
        Loop
    End If
End Sub
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
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    End With
    BuscaPedidos
End Sub
Private Sub BuscaPedidos()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, Descripcion, CANTIDAD, ENTREGADO, ID_PEDIDO FROM DETALLE_PEDIDO WHERE ENTREGADO <> 'I' AND CANTIDAD > 0 AND  ID_PEDIDO >=1881 "
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
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    ListView1.SortOrder = 1 Xor ListView1.SortOrder
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    ListView2.SortOrder = 1 Xor ListView2.SortOrder
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Label4.Caption = Item
    ListView2.ListItems.Clear
    sBuscar = "SELECT SUM(CANTIDAD) AS CANTIDAD, ACTIVO FROM REQUISICION WHERE ID_PRODUCTO = '" & Item & "' GROUP BY ACTIVO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("CANTIDAD"))
            If Not IsNull(tRs.Fields("ACTIVO")) Then
                If tRs.Fields("ACTIVO") = "0" Then
                    tLi.SubItems(1) = "EN ESPERA DE COTIZACIÓN"
                Else
                    If tRs.Fields("ACTIVO") = "1" Then
                        tLi.SubItems(1) = "EN ESPERA DE APROVACION DE COTIZACIÓN"
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT SUM(CANTIDAD) AS CANTIDAD, NUM_ORDEN, CONFIRMADA FROM VsOrdenCompra WHERE ID_PRODUCTO = '" & Item & "' GROUP BY NUM_ORDEN, CONFIRMADA ORDER BY NUM_ORDEN"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("CANTIDAD"))
            If Not IsNull(tRs.Fields("CONFIRMADA")) Then
                If tRs.Fields("CONFIRMADA") = "N" Then
                    tLi.SubItems(1) = "EN PREORDEN DE COMPRA"
                Else
                    If tRs.Fields("CONFIRMADA") = "P" Then
                        tLi.SubItems(1) = "PENDIENTE DE APROVACIÓN"
                    Else
                        If tRs.Fields("CONFIRMADA") = "S" Then
                            tLi.SubItems(1) = "EN ESPERA DE PAGO A PROVEEDOR"
                        Else
                            If tRs.Fields("CONFIRMADA") = "X" Then
                                tLi.SubItems(1) = "PAGADO Y EN CAMINO"
                            End If
                        End If
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
