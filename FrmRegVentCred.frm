VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegVentCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Venta a Credito"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   9720
      ScaleHeight     =   6195
      ScaleWidth      =   1635
      TabIndex        =   18
      Top             =   0
      Width           =   1695
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   4680
         Width           =   975
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmRegVentCred.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmRegVentCred.frx":030A
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Facturar"
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
      Left            =   6720
      Picture         =   "FrmRegVentCred.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
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
      Left            =   8280
      Picture         =   "FrmRegVentCred.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar"
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
      Left            =   6360
      Picture         =   "FrmRegVentCred.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Descripcion"
      Height          =   195
      Left            =   6840
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   195
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
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
      Left            =   8400
      Picture         =   "FrmRegVentCred.frx":9B32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label5 
      Caption         =   "Subtotal :"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Total :"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad :"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   " Clave del Producto :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmRegVentCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim ClvVenta As Integer
Dim ClvCuenta As Integer
Dim DesProducto As String
Dim VarPrecioVenta As String
Dim EliProd As String
Dim EliCant As Double
Dim EliPrecio As Double
Private Sub Command1_Click()
    Buscar
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim sum As Double
    Dim P_COSTO As String
    Dim ganan As String
    Text3.Text = Replace(Text3.Text, ".", ",")
    Text4.Text = Replace(Text4.Text, ".", ",")
    sum = Format(CDbl(Text4.Text) + (CDbl(Text3.Text) * CDbl(VarPrecioVenta)), "0.00")
    Text3.Text = Replace(Text3.Text, ",", ".")
    Text4.Text = Replace(Text4.Text, ",", ".")
    If FrmVentaCredito.Text4.Text >= sum Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tRs1 As Recordset
        Dim tLi As ListItem
        Text3.Text = Replace(Text3.Text, ",", ".")
        Text4.Text = Replace(Text4.Text, ",", ".")
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text2.Text & "' AND SUCURSAL = '" & Menu.Text4(0).Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If CDbl(tRs.Fields("CANTIDAD")) >= CDbl(Text3.Text) Then
            Dim NuevaExistencia As Double
            NuevaExistencia = Format(tRs.Fields("CANTIDAD") - CDbl(Text3.Text), "0.00")
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NuevaExistencia & " WHERE ID_PRODUCTO = '" & Text2.Text & "' AND SUCURSAL = '" & Menu.Text4(0).Text & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            VarPrecioVenta = Replace(VarPrecioVenta, ".", "")
            Text3.Text = Replace(Text3.Text, ".", "")
            Text4.Text = Replace(Text4.Text, ".", ".")
            Text3.Text = Replace(Text3.Text, ".", ",")
            Text4.Text = Replace(Text4.Text, ".", ",")
            Text5.Text = Format((CDbl(Text5.Text) + (CDbl(Text3.Text) * CDbl(VarPrecioVenta))), "0.00")
            Text4.Text = Format(CDbl(Text5.Text) * 1.15, "0.00")
            VarPrecioVenta = Replace(VarPrecioVenta, ",", ".")
            Text3.Text = Replace(Text3.Text, ",", ".")
            Text4.Text = Replace(Text4.Text, ",", ".")
            VarPrecioVenta = Replace(VarPrecioVenta, ",", ".")
            Text3.Text = Replace(Text3.Text, ",", ".")
            Text4.Text = Replace(Text4.Text, ",", ".")
            sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, CANTIDAD, ID_PRODUCTO, PRECIO_VENTA) VALUES (" & ClvCuenta & ", " & Text3.Text & ", '" & Text2.Text & "', " & VarPrecioVenta & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Text2.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            P_COSTO = tRs.Fields("PRECIO_COSTO")
            ganan = tRs.Fields("GANANCIA")
            P_COSTO = Replace(P_COSTO, ",", ".")
            ganan = Replace(ganan, ",", ".")
            sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, CANTIDAD, ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA, PRECIO_VENTA) VALUES (" & ClvVenta & ", " & Text3.Text & ", '" & Text2.Text & "', '" & DesProducto & "', " & P_COSTO & ", " & ganan & ", " & VarPrecioVenta & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT CANTIDAD, ID_PRODUCTO, PRECIO_VENTA FROM CUENTA_DETALLE WHERE ID_CUENTA = " & ClvCuenta
            Set tRs = cnn.Execute(sBuscar)
            'Me.Command2.Enabled = False
            With tRs
                If (.BOF And .EOF) Then
                    Text1.Text = ""
                    MsgBox "   No se encontro ningun producto   "
                Else
                    ListView2.ListItems.Clear
                    .MoveFirst
                    Do While Not .EOF
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi.SubItems(1) = .Fields("PRECIO_VENTA") & ""
                        tLi.SubItems(2) = .Fields("CANTIDAD")
                        .MoveNext
                    Loop
                End If
            End With
        Else
            MsgBox "    La existencia no cuanta con cantidad suficiente para surtit el pedido    "
        End If
    Else
        MsgBox "  AL AGREGAR EL ARTICULO EL CLIENTE PASA SU LIMITE DE CREDITO, ES NECESARIA UNA AUTORIZACION.   "
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub Command3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command5.SetFocus
    End If
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    Dim tLi As ListItem
    Dim NueCant As String
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & EliProd & "' AND SUCURSAL = '" & Menu.Text4(0).Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    NueCant = EliCant + CDbl(tRs.Fields("CANTIDAD"))
    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NueCant & " WHERE ID_PRODUCTO = '" & EliProd & "' AND SUCURSAL = '" & Menu.Text4(0).Text & "'"
    Set tRs1 = cnn.Execute(sBuscar)
    sBuscar = "DELETE FROM CUENTA_DETALLE WHERE ID_CUENTA = " & ClvCuenta & " AND ID_PRODUCTO = '" & EliProd & "'"
    cnn.Execute (sBuscar)
    sBuscar = "DELETE FROM VENTAS_DETALLE WHERE ID_VENTA = " & ClvVenta & " AND ID_PRODUCTO = '" & EliProd & "'"
    cnn.Execute (sBuscar)
    Text5.Text = Replace(Text5.Text, ".", ",")
    Text5.Text = Format(CDbl(Text5.Text) - (CDbl(EliPrecio) * CDbl(EliCant)), "0.00")
    Text5.Text = Replace(Text5.Text, ".", ",")
    Text4.Text = Format(CDbl(Text5.Text) * 1.15, "0.00")
    Text5.Text = Replace(Text5.Text, ",", ".")
    sBuscar = "SELECT CANTIDAD, ID_PRODUCTO, PRECIO_VENTA FROM CUENTA_DETALLE WHERE ID_CUENTA = " & ClvCuenta
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
        Else
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("PRECIO_VENTA") & ""
                tLi.SubItems(2) = .Fields("CANTIDAD")
                .MoveNext
            Loop
            
        End If
    End With
    Me.Command4.Enabled = False
    If Me.ListView2.ListItems.Count = 0 Then
        'Me.Command2.Enabled = True
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Private Sub Command5_Click()
    Text5.Text = Replace(Text5.Text, ",", ".")
    Text4.Text = Replace(Text4.Text, ",", ".")
    frmFacturaCredito.Show vbModal
    
End Sub

Private Sub Form_Load()
On Error GoTo ManejaError
    ClvCuenta = FrmVentaCredito.Text5.Text
    ClvVenta = FrmVentaCredito.Text6.Text
    Me.Command1.Enabled = False
    Me.Command3.Enabled = False
    Me.Command4.Enabled = False
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Clave del Producto", 1800
        .ColumnHeaders.Add , , "Descripcion", 7450
        .ColumnHeaders.Add , , "Precio de venta", 2450
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Clave del Producto", 1800
        .ColumnHeaders.Add , , "Precio de venta", 2450
        .ColumnHeaders.Add , , "Cantidad", 1500
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    If Option1.Value Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & Text1.Text & "%'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
            MsgBox "   No se encontro ningun producto producto   "
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then tLi.SubItems(2) = Format((.Fields("PRECIO_COSTO") * (.Fields("GANANCIA") + 1)), "0.00")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = Item
    DesProducto = Item.SubItems(1)
    VarPrecioVenta = FormatNumber(Item.SubItems(2), 2)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And DesProducto <> "" Then
        Text3.SetFocus
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EliProd = Item
    EliCant = Item.SubItems(2)
    EliPrecio = Item.SubItems(1)
    Command4.Enabled = True
End Sub
Private Sub Text1_Change()
    If Text1.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        Buscar
        ListView1.SetFocus
    End If
End Sub
Private Sub Text2_Change()
    If Text2.Text = "" Or Text3.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
End Sub
Private Sub Text3_Change()
    If Text2.Text = "" Or Text3.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
End Sub
Private Sub Text3_GotFocus()
    'Text3.SetFocus
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text3.Text <> "" Then
        Me.Command3.Enabled = True
        Me.Command3.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub

