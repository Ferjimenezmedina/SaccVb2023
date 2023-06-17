VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form VentasCostos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas-Costos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   32
      Top             =   4800
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "VentasCostos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "VentasCostos.frx":030A
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12515
      _Version        =   393216
      TabOrientation  =   3
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "VentasCostos.frx":1F0C
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
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DTPicker1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ListView1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Combo1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ListView2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Check1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      Begin VB.CheckBox Check1 
         Caption         =   "Por fecha"
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5745
         TabIndex        =   30
         Top             =   6120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   6120
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   6120
         Width           =   4815
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   6480
         TabIndex        =   23
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton Option2 
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   5400
         TabIndex        =   20
         Top             =   2760
         Width           =   1575
         Begin VB.OptionButton Option4 
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Clave"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton Command3 
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
         Left            =   8040
         Picture         =   "VentasCostos.frx":1F28
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command2 
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
         Left            =   8040
         Picture         =   "VentasCostos.frx":48FA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   16
         Top             =   6600
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar"
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
         Left            =   6600
         Picture         =   "VentasCostos.frx":72CC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         Top             =   6600
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   5520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51838977
         CurrentDate     =   39428
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   5520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51838977
         CurrentDate     =   39428
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   3000
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label8 
         Caption         =   "Producto"
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   6120
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   6120
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Utilidad total"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Al"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Del"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   0
      Top             =   6000
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "VentasCostos.frx":9C9E
         MousePointer    =   99  'Custom
         Picture         =   "VentasCostos.frx":9FA8
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label34 
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "VentasCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim Venta As String
Dim costo As String
Dim utilidad As String
Dim Sucursal As String
Dim Cliente As String
Dim producto As String
Private Sub Check1_Click()
    If Check1.value = 1 Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Dim sBuscarSucursal As String
     Dim rstSucursales As ADODB.Recordset
        sBuscarSucursal = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
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
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Check1.value = False
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Combo1.Text = ""
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
  BuscarCliente
End Sub
Private Sub Command3_Click()
    BuscarProducto
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "CLAVE DEL CLIENTE", 1000
        .ColumnHeaders.Add , , "NOMBRE", 5500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 1000
        .ColumnHeaders.Add , , "Descripcion", 5500
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image14_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    If Text6.Text <> "" Then
        Cliente = Text6.Text
    End If
    If Combo1.Text <> "" Then
      Sucursal = Combo1.Text
    End If
    If Text5.Text <> "" Then
      producto = Text5.Text
    End If
    sBuscar = "SELECT SUM (PRECIO_VENTA) AS VENTA, SUM(PRECIO_COSTO) AS COSTO FROM VsVentasCostos WHERE PRECIO_COSTO IS NOT NULL "
    If Cliente <> "" Then
        sBuscar = sBuscar & " AND ID_CLIENTE ='" & Cliente & "'"
    End If
    If Sucursal <> "" Then
        sBuscar = sBuscar & " AND SUCURSAL ='" & Sucursal & "'"
    End If
    If producto <> "" Then
        sBuscar = sBuscar & " AND ID_PRODUCTO ='" & producto & "'"
    End If
    If Check1.value = 1 Then
        sBuscar = sBuscar & " AND FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & "'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not IsNull(.Fields("VENTA")) Then
            Venta = .Fields("VENTA")
        Else
            Venta = 0
        End If
        If Not IsNull(.Fields("COSTO")) Then
            costo = .Fields("COSTO")
        Else
            costo = 0
        End If
    End With
    utilidad = Venta - costo
    Text3.Text = utilidad
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.COUNT > 0 Then
        Text4.Text = Item.SubItems(1)
        Text6.Text = Item
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView2.ListItems.COUNT > 0 Then
        Text5.Text = Item
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim BuscarCliente As String
        Dim tLi As ListItem
        If Option2.value = True Then
            BuscarCliente = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        ElseIf Option1.value = True Then
            BuscarCliente = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE ID_CLIENTE = " & Text1.Text
        End If
        Set tRs = cnn.Execute(BuscarCliente)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                .MoveNext
            Loop
        End With
    End If
    Dim Valido As String
    If Option1.value = True Then
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
    If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscarProducto As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option3.value = True Then
            sBuscarProducto = "SELECT ID_PRODUCTO, Descripcion FROM VsAlmacenes123 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' ORDER BY ID_PRODUCTO"
        ElseIf Option4.value = True Then
            sBuscarProducto = "SELECT ID_PRODUCTO, Descripcion FROM VsAlmacenes123 WHERE Descripcion LIKE '%" & Text2.Text & "%' ORDER BY Descripcion"
        End If
        Set tRs = cnn.Execute(sBuscarProducto)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
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
Public Sub BuscarCliente()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim BuscarCliente As String
    Dim tLi As ListItem
    If Option2.value = True Then
        BuscarCliente = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
    ElseIf Option1.value = True Then
        BuscarCliente = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE ID_CLIENTE LIKE '%" & Text1.Text & "%'"
    End If
    Set tRs = cnn.Execute(BuscarCliente)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
            If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub BuscarProducto()
        On Error GoTo ManejaError
        Dim sBuscarProducto As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option3.value = True Then
            sBuscarProducto = "SELECT ID_PRODUCTO, Descripcion FROM VsAlmacenes123 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' ORDER BY ID_PRODUCTO"
        ElseIf Option4.value = True Then
            sBuscarProducto = "SELECT ID_PRODUCTO, Descripcion FROM VsAlmacenes123 WHERE Descripcion LIKE '%" & Text2.Text & "%' ORDER BY Descripcion"
        End If
        Set tRs = cnn.Execute(sBuscarProducto)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
