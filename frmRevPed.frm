VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRevPed 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Revisar Pedidos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   4
      Top             =   5640
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmRevPed.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmRevPed.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmRevPed.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBorrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
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
         Left            =   8040
         Picture         =   "frmRevPed.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   5640
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   10186
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmRevPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numRem As Integer
Dim tLi As ListItem
Private cnn As ADODB.Connection
Private Sub cmdBorrar_Click()
    On Error GoTo ManejaError
    If numRem <> "0" Then
        ListView1.ListItems.Remove (numRem)
        numRem = 0
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim ID As Integer
        sBuscar = "SELECT ID FROM DETALLE_PEDIDO WHERE ID ='" & txtId.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                sBuscar2 = "UPDATE DETALLE_PEDIDO SET ENTREGADO ='1' WHERE ID ='" & tRs.Fields("ID") & "'"
                Set tRs2 = cnn.Execute(sBuscar2)
            End If
        End With
    Else
        MsgBox "NO SELECCIONO NINGUN PEDIDO", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
        .ColumnHeaders.Add , , "PIDIO", 2000
        .ColumnHeaders.Add , , "FECHA", 1440
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "Descripcion", 5000
        .ColumnHeaders.Add , , "ID_DETALLE", 0
    End With
    LLENA_LISTA_PEDIDOS
End Sub
Public Sub LLENA_LISTA_PEDIDOS()
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    ListView1.ListItems.Clear
    sBuscar = "SELECT  P.ID_PEDIDO, P.SUCURSAL, P.PIDIO, P.FECHA, D.ID, D.ID_PRODUCTO, D.CANTIDAD, A.Descripcion FROM PEDIDO AS P JOIN DETALLE_PEDIDO AS D ON D.ID_PEDIDO = P.ID_PEDIDO JOIN ALMACEN3 AS A ON A.ID_PRODUCTO = D.ID_PRODUCTO WHERE D.ENTREGADO = '0' ORDER BY P.SUCURSAL"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PEDIDO"))
            If Not IsNull(.Fields("PIDIO")) Then tLi.SubItems(1) = Trim(.Fields("Sucursal"))
            If Not IsNull(.Fields("PIDIO")) Then tLi.SubItems(2) = Trim(.Fields("PIDIO"))
            If Not IsNull(.Fields("fecha")) Then tLi.SubItems(3) = Trim(.Fields("fecha"))
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(6) = Trim(.Fields("Descripcion"))
            If Not IsNull(.Fields("ID")) Then tLi.SubItems(7) = Trim(.Fields("ID"))
            .MoveNext
        Wend
        .Close
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    ListView1.SortOrder = 1 Xor ListView1.SortOrder
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ManejaError:
    txtId = Item.SubItems(7)
    numRem = Item.Index
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
