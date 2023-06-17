VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEntradaExistenciasTemporales 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Entrada de existencias temporales"
   ClientHeight    =   5535
   ClientLeft      =   4425
   ClientTop       =   1965
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10350
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   10
      Top             =   3000
      Width           =   975
      Begin VB.Label Label8 
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
      Begin VB.Image Image8 
         Enabled         =   0   'False
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmEntradaExistenciasTemporales.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmEntradaExistenciasTemporales.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   8
      Top             =   4200
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmEntradaExistenciasTemporales.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "frmEntradaExistenciasTemporales.frx":1FD6
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label9 
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmEntradaExistenciasTemporales.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCliente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DTFecha"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Producto y Movimiento"
      TabPicture(1)   =   "frmEntradaExistenciasTemporales.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(3)=   "Text3"
      Tab(1).Control(4)=   "ListView2"
      Tab(1).Control(5)=   "Option1"
      Tab(1).Control(6)=   "Option2"
      Tab(1).Control(7)=   "txtProducto"
      Tab(1).Control(8)=   "btnAgregar"
      Tab(1).Control(9)=   "ListViewArticulosCesta"
      Tab(1).Control(10)=   "btnQuitar"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton btnQuitar 
         Caption         =   "Quitar"
         Enabled         =   0   'False
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
         Picture         =   "frmEntradaExistenciasTemporales.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4800
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListViewArticulosCesta 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   24
         Top             =   3120
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton btnAgregar 
         Caption         =   "Agregar"
         Enabled         =   0   'False
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
         Picture         =   "frmEntradaExistenciasTemporales.frx":6AC2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2640
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   375
         Left            =   7320
         TabIndex        =   22
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52559873
         CurrentDate     =   39246
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   4560
         Width           =   7695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCliente 
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtProducto 
         Height          =   495
         Left            =   -74640
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Descripción"
         Height          =   195
         Left            =   -68040
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clave"
         Height          =   195
         Left            =   -68040
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   4
         Top             =   960
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68520
         TabIndex        =   2
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73920
         TabIndex        =   1
         Top             =   600
         Width           =   5775
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3015
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5318
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
         Caption         =   "Fecha"
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Motivo"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   -69360
         TabIndex        =   3
         Top             =   2760
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmEntradaExistenciasTemporales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub btnAgregar_Click()
Dim Selected As Integer
Dim tLi As ListItem
    Selected = ListView2.SelectedItem.Index
    Set tLi = ListViewArticulosCesta.ListItems.Add(, , ListView2.ListItems.Item(Selected) & "")
    tLi.SubItems(1) = ListView2.ListItems.Item(Selected).SubItems(1)
    tLi.SubItems(2) = Text3.Text
    Text3.Text = ""
    btnQuitar.Enabled = True
    If Combo1.Text <> "" And txtCliente.Text <> "" And ListViewArticulosCesta.ListItems.Count <> 0 Then
        Image8.Enabled = True
    End If
    Text2.Text = ""
End Sub
Private Sub btnQuitar_Click()
    Dim Selected As Integer
    Dim tLi As ListItem
    Selected = ListViewArticulosCesta.SelectedItem.Index
    ListViewArticulosCesta.ListItems.Remove (Selected)
    If ListViewArticulosCesta.ListItems.Count = 0 Then
        Image8.Enabled = False
        btnQuitar.Enabled = False
        Text3.Enabled = False
        ListView2.ListItems.Clear
        Text2.SetFocus
    End If
End Sub
Private Sub Combo1_DropDown()
    Dim sBuscarSucursal As String
    Dim rstSucursales As ADODB.Recordset
    sBuscarSucursal = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set rstSucursales = cnn.Execute(sBuscarSucursal)
    With rstSucursales
        Combo1.Clear
        If (.EOF And .BOF) Then
            MsgBox ("NO EXISTEN SUCURSALES")
        Else
            Do While Not .EOF
                If Not IsNull(.Fields("NOMBRE")) Then
                    Combo1.AddItem (.Fields("NOMBRE"))
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Combo1_LostFocus()
    Combo1.Enabled = False
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTFecha.Value = Format(Date, "dd/mm/yyyy")
    DTFecha.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 6500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE DEL CLIENTE", 2000
        .ColumnHeaders.Add , , "NOMBRE", 6500
     End With
    With ListViewArticulosCesta
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5450
        .ColumnHeaders.Add , , "CANTIDAD", 1000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim sBuscar3 As String
    Dim sBuscar4 As String
    Dim sBuscar5 As String
    Dim sBuscar6 As String
    Dim nIDExistenciaTemporal
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim Cont As Integer
    Dim cant As Double
    Dim operacionCambioExistencias As Double
    sBuscar = "INSERT INTO EXISTENCIAS_TEMPORAL (ID_USUARIO,FECHA,MOTIVO,ID_CLIENTE)VALUES ('" & VarMen.Text1(0).Text & "', '" & DTFecha.Value & "','" & Text4.Text & "', '" & txtCliente.Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    sBuscar2 = "SELECT TOP 1 ID_MOVEXISTENCIA FROM EXISTENCIAS_TEMPORAL ORDER BY ID_MOVEXISTENCIA DESC"
    Set tRs2 = cnn.Execute(sBuscar2)
    nIDExistenciaTemporal = tRs2.Fields("ID_MOVEXISTENCIA")
    For Cont = 1 To ListViewArticulosCesta.ListItems.Count
        operacionCambioExistencias = 0
        cant = 0
        sBuscar3 = "INSERT INTO EXISTENCIAS_TEMPORAL_DETALLES (ID_MOVEXISTENCIA,ID_PRODUCTO,CANTIDAD,SUCURSAL)  VALUES"
        sBuscar3 = sBuscar3 & "('" & nIDExistenciaTemporal
        sBuscar3 = sBuscar3 & "', '" & ListViewArticulosCesta.ListItems(Cont)
        sBuscar3 = sBuscar3 & "', '" & ListViewArticulosCesta.ListItems.Item(Cont).SubItems(2)
        sBuscar3 = sBuscar3 & "', '" & Combo1.Text & "' );"
        Set tRs3 = cnn.Execute(sBuscar3)
        cant = ListViewArticulosCesta.ListItems.Item(Cont).SubItems(2)
        sBuscar4 = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListViewArticulosCesta.ListItems(Cont) & "' AND SUCURSAL = '" & Combo1.Text & "'"
        Set tRs4 = cnn.Execute(sBuscar4)
        If Not (tRs4.EOF And tRs4.BOF) Then
            operacionCambioExistencias = CDbl(tRs4.Fields("CANTIDAD")) + CDbl(cant)
            sBuscar5 = "UPDATE EXISTENCIAS SET CANTIDAD = " & operacionCambioExistencias & " WHERE SUCURSAL = '" & Combo1.Text & " ' AND ID_PRODUCTO = '" & ListViewArticulosCesta.ListItems(Cont) & "'"
            Set tRs5 = cnn.Execute(sBuscar5)
        Else
            sBuscar5 = "INSERT INTO EXISTENCIAS(ID_PRODUCTO,CANTIDAD,SUCURSAL) VALUES ('" & ListViewArticulosCesta.ListItems(Cont) & "', '" & ListViewArticulosCesta.ListItems.Item(Cont).SubItems(2) & "','" & Combo1.Text & "')"
            Set tRs5 = cnn.Execute(sBuscar5)
        End If
    Next Cont
    Limpiar
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView2.ListItems.Count > 0 Then
        txtProducto = Item
        Text2.Text = Item.SubItems(1)
        Text3.Enabled = True
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
      If ListView3.ListItems.Count > 0 Then
        txtCliente = Item
        Text5.Text = Item.SubItems(1)
        Text4.Enabled = True
      End If
End Sub
Private Sub Option1_Click()
    If Option1.Value Then
        Text2.SetFocus
    End If
End Sub
Private Sub Option2_Click()
    If Option2.Value Then
        Text2.SetFocus
    End If
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text2.Text & "%' ORDER BY Descripcion"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
            ListView2.SetFocus
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text3_Change()
    If (Text3.Text <> "") Then
       btnAgregar.Enabled = True
   Else
       btnAgregar.Enabled = False
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&#%@!?*+"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Combo1.Text <> "" Then
       If KeyAscii = 13 Then
           Dim tRs As ADODB.Recordset
           Dim tLi As ListItem
           Dim sBus As String
           sBus = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text5.Text & "%'"
           Set tRs = cnn.Execute(sBus)
           With tRs
               ListView3.ListItems.Clear
               Do While Not .EOF
                   Set tLi = ListView3.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                   If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                   .MoveNext
               Loop
               ListView3.SetFocus
           End With
       End If
       Dim Valido As String
       Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890.% "
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       If KeyAscii > 26 Then
           If InStr(Valido, Chr(KeyAscii)) = 0 Then
               KeyAscii = 0
           End If
       End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Public Sub Limpiar()
    Combo1.Clear
    Text5.Text = ""
    ListView2.ListItems.Clear
    Text4.Text = ""
    Text4.Enabled = False
    txtCliente.Text = ""
    Text2.Text = ""
    ListView3.ListItems.Clear
    txtProducto.Text = ""
    Text3.Text = ""
    Text3.Enabled = False
    ListViewArticulosCesta.ListItems.Clear
    Combo1.Enabled = True
    Image8.Enabled = False
    btnQuitar.Enabled = False
End Sub
