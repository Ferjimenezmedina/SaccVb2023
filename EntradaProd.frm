VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form EntradaProd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Entrada de productos al Almacen"
   ClientHeight    =   4950
   ClientLeft      =   3165
   ClientTop       =   3270
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   10575
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   37
      Top             =   2400
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "EntradaProd.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "EntradaProd.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label11 
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame26 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   35
      Top             =   1200
      Width           =   975
      Begin VB.Image Image24 
         Height          =   765
         Left            =   240
         MouseIcon       =   "EntradaProd.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "EntradaProd.frx":1FD6
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ver"
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   31
      Top             =   3600
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "EntradaProd.frx":3A64
         MousePointer    =   99  'Custom
         Picture         =   "EntradaProd.frx":3D6E
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
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "Proveedor"
      TabPicture(0)   =   "EntradaProd.frx":5E50
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListProv"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProveedor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Guardar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Producto"
      TabPicture(1)   =   "EntradaProd.frx":5E6C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Option2"
      Tab(1).Control(1)=   "Option1"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "ListView1"
      Tab(1).Control(4)=   "txtClaveProducto"
      Tab(1).Control(5)=   "Text5"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label6"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton Guardar 
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
         Height          =   375
         Left            =   7800
         Picture         =   "EntradaProd.frx":5E88
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4080
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Descripcion"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -70440
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clave"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -70440
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Entrada"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   25
         Top             =   2760
         Width           =   9015
         Begin VB.TextBox txtSucursal 
            Height          =   285
            Left            =   8280
            TabIndex        =   33
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "EntradaProd.frx":885A
            Left            =   6480
            List            =   "EntradaProd.frx":8864
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1320
            Width           =   6975
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   12
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            MaxLength       =   8
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "Codigo de barras"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Sucursal"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   5760
            TabIndex        =   28
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Precio"
            Height          =   255
            Left            =   3600
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.TextBox txtClaveProducto 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67680
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74040
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   3600
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7560
         TabIndex        =   20
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51511297
         CurrentDate     =   39260
      End
      Begin MSComctlLib.ListView ListProv 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4048
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
         Left            =   960
         TabIndex        =   0
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label7 
         Caption         =   "Clave del producto"
         Height          =   255
         Left            =   -69120
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Producto"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   6960
         TabIndex        =   19
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Total"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Numero de entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de factura"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "EntradaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Combo1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_SUCURSAL,NOMBRE FROM SUCURSALES WHERE NOMBRE ='" & Combo1.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    txtSucursal.Text = tRs.Fields("ID_SUCURSAL")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then
                Combo1.AddItem tRs.Fields("NOMBRE")
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HFFE1E1
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &H80000005
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListProv
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Proveedor", 1500
        .ColumnHeaders.Add , , "Nombre", 6100
        .ColumnHeaders.Add , , "Ciudad", 2300
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2400
        .ColumnHeaders.Add , , "Descripcion", 7000
    End With
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Guardar_Click()
On Error GoTo ManejaError
    If txtProveedor.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And DTPicker1.Value <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim tRs3 As ADODB.Recordset
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim sBuscar3 As String
        sBuscar = "INSERT INTO ENTRADAS(ID_PROVEEDOR, FECHA, TOTAL, FACTURA, ID_USUARIO) VALUES('" & txtProveedor & "','" & DTPicker1.Value & "','" & Text3.Text & "','" & Text2.Text & "','" & VarMen.Text1(0) & "')"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar2 = "SELECT MAX(ID_ENTRADA) AS N FROM ENTRADAS ORDER BY N"
        Set tRs2 = cnn.Execute(sBuscar2)
        Text4.Text = tRs2.Fields("N")
        Text1.Enabled = False
        ListProv.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text5.Enabled = True
        ListView1.Enabled = True
        Option1.Enabled = True
        Option2.Enabled = True
        Guardar.Enabled = False
    Else
        MsgBox ("Falta Informacion necesaria")
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Image24_Click()
On Error GoTo ManejaError
    VERENTRADA.Show vbModal
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Image8_Click()
    If txtClaveProducto.Text <> "" And txtSucursal.Text <> "" And Text7.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim tRs3 As ADODB.Recordset
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim sBuscar3 As String
        Dim cantExistencia As Double
        Dim cant As Double
        
        cant = Text7.Text
        sBuscar = "INSERT INTO ENTRADA_PRODUCTO(ID_ENTRADA, ID_PRODUCTO, CANTIDAD, PRECIO,MONEDA, FECHA, ID_SUCURSAL, CODIGO_BARAS) VALUES('" & Text4.Text & "','" & txtClaveProducto.Text & "', '" & Text7.Text & "','" & Text8.Text & "','" & Combo2.Text & "','" & DTPicker1.Value & "','" & txtSucursal.Text & "', '" & Text10.Text & "')"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar2 = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & txtClaveProducto.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        Set tRs2 = cnn.Execute(sBuscar2)
        If Not (tRs2.BOF And tRs2.EOF) Then
            cantExistencia = tRs2.Fields("CANTIDAD") + cant
            sBuscar3 = "UPDATE EXISTENCIAS SET CANTIDAD =" & cantExistencia & " WHERE SUCURSAL = '" & Combo1.Text & " ' AND ID_PRODUCTO = '" & txtClaveProducto.Text & "'"
            Set tRs3 = cnn.Execute(sBuscar3)
        Else
            sBuscar3 = "INSERT INTO EXISTENCIAS(ID_PRODUCTO,CANTIDAD,SUCURSAL) VALUES('" & txtClaveProducto.Text & "','" & cant & "', '" & Combo1.Text & "')"
            Set tRs3 = cnn.Execute(sBuscar3)
        End If
        Text5.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Text10.Text = ""
    Else
        MsgBox "Falta informacion necesaria para el traspaso o el inventario no existe"
    End If
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListProv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtProveedor = Item
    Text2.Enabled = True
    Text3.Enabled = True
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text8.Enabled = True
    Text10.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    'cmdAdd.Enabled = True
    txtClaveProducto.Text = Item
    Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear

End Sub

Private Sub Option1_Click()
    Text5.SetFocus
End Sub
Private Sub Option2_Click()
    Text5.SetFocus
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar
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
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Public Sub Buscar()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR,NOMBRE,CIUDAD FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    ListProv.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
                Set tLi = ListProv.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("CIUDAD")
                tRs.MoveNext
        Loop
    ListProv.SetFocus
    End If
End Sub
Private Sub Text10_Change()
On Error GoTo ManejaError
        Text10.Text = Replace(Text10.Text, ",", "")
        Text10.Text = Replace(Text10.Text, "-", "")
        Text10.Text = Replace(Text10.Text, "_", "")
        Text10.Text = Replace(Text10.Text, ".", "")
        Text10.Text = Replace(Text10.Text, "*", "")
        Text10.Text = Replace(Text10.Text, "%", "")
        Text10.Text = Replace(Text10.Text, "&", "")
        Text10.Text = Replace(Text10.Text, "/", "")
        Text10.Text = Replace(Text10.Text, "'", "")
        Text10.Text = Replace(Text10.Text, "$", "")
        Text10.Text = Replace(Text10.Text, "=", "")
        Text10.Text = Replace(Text10.Text, "@", "")
        Text10.Text = Replace(Text10.Text, "!", "")
        Text10.Text = Replace(Text10.Text, "?", "")
        Text10.Text = Replace(Text10.Text, "^", "")
        Text10.Text = Replace(Text10.Text, "#", "")
        Text10.Text = Replace(Text10.Text, " ", "")
        Text10.Text = Replace(Text10.Text, "+", "")
        Text10.Text = Replace(Text10.Text, ";", "")
        Text10.Text = Replace(Text10.Text, ":", "")
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text10_GotFocus()
    Text10.BackColor = &HFFE1E1
End Sub

Private Sub Text10_LostFocus()
    Text10.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
        On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar2
    End If
    Dim Valido As String
    Valido = "1234567890"
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

Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar2
    End If
    Dim Valido As String
    Valido = "1234567890. "
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
Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    Text7.Enabled = True
    If KeyAscii = 13 Then
        Buscar2
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
Public Sub Buscar2()
   On Error GoTo ManejaError
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM vsALMACENES_123 WHERE ID_PRODUCTO LIKE '%" & Text5.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM vsALMACENES_123 WHERE Descripcion LIKE '%" & Text5.Text & "%' ORDER BY Descripcion"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = tRs.Fields("Descripcion")
                    tRs.MoveNext
            Loop
        ListView1.SetFocus
        End If
  Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HFFE1E1
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
        Valido = "1234567890"
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
Private Sub Text7_LostFocus()
    Text7.BackColor = &H80000005
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    Dim Valido As String
        Valido = "1234567890."
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
Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub

