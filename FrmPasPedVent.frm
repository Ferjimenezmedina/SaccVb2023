VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPasPedVent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pasar Pedido a Venta ( Peidos de Clientes )"
   ClientHeight    =   6150
   ClientLeft      =   585
   ClientTop       =   1395
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   10320
      ScaleHeight     =   6075
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   0
      Width           =   1455
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   4680
         Width           =   975
         Begin VB.Label Label10 
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
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmPasPedVent.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmPasPedVent.frx":030A
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.CommandButton ELLI 
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
      Left            =   8880
      Picture         =   "FrmPasPedVent.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   4560
      Picture         =   "FrmPasPedVent.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox TxtNoPed 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.ListView Lvw2 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
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
   Begin MSComctlLib.ListView Lvw1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
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
   Begin VB.Label Label1 
      Caption         =   "No. de Pedido Seleccionado :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
End
Attribute VB_Name = "FrmPasPedVent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim guia As String
Dim elind As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdCerrar_Click()
On Error GoTo ManejaError
    Dim ClvCliente As Integer
    Dim NomCliente As String
    Dim TotVenta As String
    Dim DesClente As String
    Dim fecha As String
    Dim ClvUsuario As String
    Dim ClvVenta As Integer
    Dim ClvProducto As String
    Dim DesProducto As String
    Dim CantProducto As String
    Dim PreVenta As String
    Dim PreCosto As String
    Dim GanProducto As String
    Dim sBuscar As String
    DesClente = "0"
    TotVenta = "0"
    PreVenta = "0"
    PreCosto = "0"
    GanProducto = "0"
    CantProducto = "0"
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    sBuscar = "SELECT ID_CLIENTE, USUARIO FROM PED_CLIEN WHERE NO_PEDIDO = " & TxtNoPed.Text
    Set tRs = cnn.Execute(sBuscar)
    ClvCliente = tRs.Fields("ID_CLIENTE")
    ClvUsuario = tRs.Fields("USUARIO")
    sBuscar = "SELECT NOMBRE, DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & ClvCliente
    Set tRs = cnn.Execute(sBuscar)
    If tRs.Fields("DESCUENTO") <> "" Then
        DesClente = tRs.Fields("DESCUENTO")
    End If
    NomCliente = tRs.Fields("NOMBRE")
    ClvUsuario = Menu.Text1(1).Text
    fecha = Date
    sBuscar = "SELECT ID_PRODUCTO FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & TxtNoPed.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
            sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            TotVenta = CDbl(TotVenta) + (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA"))))
            tRs.MoveNext
        Loop
    End If
    Dim DES As Double
    DES = CDbl(TotVenta) * CDbl(DesClente)
    TotVenta = CDbl(TotVenta) - DES
    TotVenta = Replace(TotVenta, ",", ".")
    DesClente = Replace(DesClente, ",", ".")
    sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, TOTAL, DESCUENTO, ID_USUARIO, FECHA, SUCURSAL) VALUES (" & ClvCliente & ", '" & NomCliente & "', '" & TotVenta & "', '" & DesClente & "', '" & ClvUsuario & "', '" & fecha & "', 'BODEGA');"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT ID_VENTA FROM VENTAS ORDER BY ID_VENTA DESC"
    Set tRs = cnn.Execute(sBuscar)
    ClvVenta = tRs.Fields("ID_VENTA")
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD_PEDIDA FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & TxtNoPed.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
            ClvProducto = tRs.Fields("ID_PRODUCTO")
            CantProducto = tRs.Fields("CANTIDAD_PEDIDA")
            sBuscar = "SELECT DESCRIPCION, GANANCIA, PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProducto & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            DesProducto = tRs1.Fields("DESCRIPCION")
            PreCosto = tRs1.Fields("PRECIO_COSTO")
            GanProducto = tRs1.Fields("GANANCIA")
            PreVenta = CDbl(PreCosto) * (1 + CDbl(GanProducto))
            DesProducto = Replace(DesProducto, ",", ".")
            CantProducto = Replace(CantProducto, ",", ".")
            PreVenta = Replace(PreVenta, ",", ".")
            PreCosto = Replace(PreCosto, ",", ".")
            CantProducto = Replace(CantProducto, ",", ".")
            GanProducto = Replace(GanProducto, ",", ".")
            sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA) VALUES (" & ClvVenta & ", '" & ClvProducto & "', '" & DesProducto & "', '" & CantProducto & "', '" & PreVenta & "', '" & PreCosto & "', '" & GanProducto & "');"
            cnn.Execute (sBuscar)
            tRs.MoveNext
            'FALTA IMPRIMA TICKER DE VENTA... CON DATOS
        Loop
    End If
    sBuscar = "INSERT INTO PEDIDO_VENTA (ID_VENTA, NO_PEDIDO) VALUES (" & ClvVenta & ", " & TxtNoPed.Text & ");"
    cnn.Execute (sBuscar)
    sBuscar = "DELETE FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = '" & TxtNoPed.Text & "'"
    cnn.Execute (sBuscar)
    sBuscar = "DELETE FROM PED_CLIEN WHERE NO_PEDIDO = '" & TxtNoPed.Text & "'"
    cnn.Execute (sBuscar)
    Actualizar
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
Private Sub ELLI_Click()
On Error GoTo ManejaError
    Dim sEliminar As String
    sEliminar = "DELETE FROM PED_CLIEN WHERE NO_PEDIDO = " & guia
    cnn.Execute (sEliminar)
    Me.ELLI.Enabled = False
    Lvw1.ListItems.Clear
    Actualizar
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
'Private Sub Form_Load()
'    Me.ELLI.Enabled = False
'    Const sPathBase As String = "LINUX"
'    Set cnn = New ADODB.Connection
'    Set rst = New ADODB.Recordset
'    With cnn
'        .ConnectionString = _
'            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
'            "Data Source=" & sPathBase & ";"
'        .Open
'    End With
'    With Lvw1
'        .View = lvwReport
'        .Gridlines = True
'        .LabelEdit = lvwManual
 '       .ColumnHeaders.Add , , "No. Pedido", 1000
 '       .ColumnHeaders.Add , , "Capturo", 1500
 '       .ColumnHeaders.Add , , "Cliente", 6500
 '       .ColumnHeaders.Add , , "Fecha", 1500
 '   End With
 '   With Lvw2
 '       .View = lvwReport
 '       .Gridlines = True
 '       .LabelEdit = lvwManual
 '       .ColumnHeaders.Add , , "Clave del Producto", 4500
 '       .ColumnHeaders.Add , , "Cantidad Pedida", 2000
 '       .ColumnHeaders.Add , , "Cantidad en Existencia", 2000
 '       .ColumnHeaders.Add , , "Cantidad Pendiente", 2000
'    End With
'    Actualizar
'End Sub
Private Sub Actualizar()
    CmdCerrar.Enabled = False
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As Recordset
    Dim tRs2 As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM PED_CLIEN WHERE ESTADO LIKE 'C'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Lvw1.ListItems.Clear
            Do While Not .EOF
                Set tLi = Lvw1.ListItems.Add(, , .Fields("NO_PEDIDO") & "")
                tLi.SubItems(1) = .Fields("USUARIO") & ""
                BusClie = "SELECT * FROM CLIENTE WHERE ID_CLIENTE =" & .Fields("ID_CLIENTE")
                Set tRs2 = cnn.Execute(BusClie)
                tLi.SubItems(2) = tRs2.Fields("NOMBRE") & ""
                tLi.SubItems(3) = .Fields("FECHA") & ""
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

Private Sub Form_Load()

End Sub

Private Sub Lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & CDbl(Item)
    Set tRs = cnn.Execute(sBuscar)
    If (tRs.BOF And tRs.EOF) Then
        Lvw2.ListItems.Clear
        TxtNoPed.Text = ""
        MsgBox "                   Este pedido esta vacio                   "
        guia = Item
        Me.ELLI.Enabled = True
        elind = Item.Index
    Else
        Lvw2.ListItems.Clear
        TxtNoPed.Text = Item
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = Lvw2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
            tLi.SubItems(1) = tRs.Fields("CANTIDAD_PEDIDA") & ""
            tLi.SubItems(2) = tRs.Fields("CANTIDAD_EXISTENCIA") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD_PENDIENTE") & ""
            tRs.MoveNext
        Loop
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
Private Sub TxtNoPed_Change()
On Error GoTo ManejaError
    If TxtNoPed.Text = "" Then
        CmdCerrar.Enabled = False
    Else
        CmdCerrar.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtNoPed_GotFocus()
    TxtNoPed.BackColor = &HFFE1E1
End Sub
Private Sub TxtNoPed_LostFocus()
      TxtNoPed.BackColor = &H80000005
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


