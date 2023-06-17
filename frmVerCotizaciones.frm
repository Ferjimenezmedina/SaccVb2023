VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmVerCotizaciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Cotizaciónes"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   17
      Top             =   3840
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmVerCotizaciones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmVerCotizaciones.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmVerCotizaciones.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwProveedores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwCotizaciones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   5175
         Begin VB.TextBox txtDias 
            Height          =   285
            Left            =   360
            TabIndex        =   1
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtPrecio 
            CausesValidation=   0   'False
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
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
            Left            =   3720
            Picture         =   "frmVerCotizaciones.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtCotizacion 
            Height          =   285
            Left            =   3720
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtProveedor 
            Height          =   285
            Left            =   3720
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtRequisicion 
            Height          =   285
            Left            =   3720
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtIndex 
            Height          =   285
            Left            =   5040
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   840
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmVerCotizaciones.frx":4DDA
            Left            =   2280
            List            =   "frmVerCotizaciones.frx":4DE4
            TabIndex        =   3
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dias de entrega"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Precio"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblProveedor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   2040
            TabIndex        =   14
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblProducto 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2040
            TabIndex        =   13
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   2280
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lvwCotizaciones 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwProveedores 
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmVerCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim sDescripcion As String
Dim sCantidad As String
Dim tRs As ADODB.Recordset
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    If Puede_Agregar Then
        Dim tRs1 As ADODB.Recordset
        Dim iAfectados As Long
        sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'A',  DIAS_ENTREGA = " & Me.txtDias.Text & ", PRECIO = " & Me.txtPrecio.Text & ", MONEDA = '" & Combo1.Text & "' WHERE ID_COTIZACION IN (" & Me.txtCotizacion.Text & ") AND ID_PRODUCTO = '" & lblProducto.Caption & "'"
        cnn.Execute (sqlQuery)
        sqlQuery = "UPDATE REQUISICION SET COTIZADA = 1, ACTIVO = 0 WHERE ID_REQUISICION IN (" & Me.txtRequisicion.Text & ") AND ID_PRODUCTO = '" & lblProducto.Caption & "'"
        Set tRs1 = cnn.Execute(sqlQuery, iAfectados, adCmdText)
        If iAfectados = 0 Then
            sqlQuery = "INSERT INTO REQUISICION (ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, COTIZADA, ACTIVO, CONTADOR, ALMACEN, URGENTE, MARCA, COMENTARIO) VALUES ('" & Me.lblProducto.Caption & "', '" & sDescripcion & "', " & sCantidad & ", '" & Date & "', 1, 0, 1, 'A3', 'N', 'N/E', 'REINSERTADO POR EL SISTEMA POR PERDIDA DE DATOS')"
            cnn.Execute (sqlQuery)
            'sqlQuery = "SELECT TOP 1 ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
            'Set tRs1 = cnn.Execute(sqlQuery)
            'sqlQuery = "UPDATE COTIZA_REQUI SET ID_REQUISICION = " & tRs1.Fields("ID_REQUISICION") & " WHERE ID_COTIZACION = " & Me.txtCotizacion.Text
            'cnn.Execute (sqlQuery)
        End If
        If VarMen.TxtEmp(12).Text = "SIMPLE" Then
            sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'X' WHERE ID_COTIZACION IN (" & Me.txtCotizacion.Text & ") AND ID_PRODUCTO = '" & lblProducto.Caption & "'"
            cnn.Execute (sqlQuery)
            sqlQuery = "UPDATE REQUISICION SET ACTIVO = '1' WHERE ID_REQUISICION IN (" & Me.txtRequisicion.Text & ") AND ID_PRODUCTO = '" & lblProducto.Caption & "'"
            cnn.Execute (sqlQuery)
            sqlQuery = "DELETE FROM COTIZA_REQUI WHERE ESTADO_ACTUAL <> 'X' AND ID_COTIZACION IN (" & Me.txtCotizacion.Text & ") AND ID_PRODUCTO = '" & lblProducto.Caption & "'"
            cnn.Execute (sqlQuery)
        End If
        Llenar_Lista_Cotizaciones (Me.txtProveedor.Text)
        Me.txtCotizacion.Text = ""
        Me.txtDias.Text = ""
        Me.txtPrecio.Text = ""
        Me.txtRequisicion.Text = ""
        'lvwCotizaciones.SetFocus
        'lvwCotizaciones.ListItems.Item(Val(txtIndex.Text)).Selected = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo2_Click()
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    End If
End Sub
Private Sub Form_Activate()
On Error GoTo ManejaError
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
    With Me.lvwCotizaciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 2200
        .ColumnHeaders.Add , , "PRECIO", 2200
        .ColumnHeaders.Add , , "FECHA", 2000
        .ColumnHeaders.Add , , "MONEDA", 2200
    End With
    With Me.lvwProveedores
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "PROVEEDOR", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Cotizaciones(nId_Proveedor As Integer)
On Error GoTo ManejaError
    'sqlQuery = "SELECT ID_COTIZACION, ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, MONEDA FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'A' AND ID_PROVEEDOR = " & nId_Proveedor
    sqlQuery = "SELECT STUFF((SELECT ',' + CONVERT(VARCHAR(20), ID_COTIZACION) FROM COTIZA_REQUI CR WHERE CR.FOLIO = RE.FOLIO  AND (ESTADO_ACTUAL = 'A') AND (ID_PROVEEDOR = " & nId_Proveedor & ") FOR XML PATH('')), 1, 1, '') AS ID_COTIZACION, STUFF((SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM COTIZA_REQUI CR WHERE CR.FOLIO  = RE.FOLIO AND (ESTADO_ACTUAL = 'A') AND (ID_PROVEEDOR = " & nId_Proveedor & ") FOR XML PATH('')), 1, 1, '') AS ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, Descripcion, SUM(CANTIDAD) AS CANTIDAD, DIAS_ENTREGA, Precio, fecha, Moneda FROM COTIZA_REQUI RE WHERE (ESTADO_ACTUAL = 'A') AND (ID_PROVEEDOR = " & nId_Proveedor & ") GROUP BY FOLIO, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, DIAS_ENTREGA, PRECIO, FECHA, MONEDA"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwCotizaciones.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = .Fields("ID_REQUISICION")
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(4) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(6) = Trim(.Fields("DIAS_ENTREGA"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(7) = Trim(.Fields("PRECIO"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(8) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("MONEDA")) Then tLi.SubItems(9) = Trim(.Fields("MONEDA"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Cotizaciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_COTIZACION)ID_COTIZACION FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'A'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_COTIZACION") <> 0 Then
            Hay_Cotizaciones = True
        Else
            Hay_Cotizaciones = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    Dim nId_Proveedor As Integer
    sqlQuery = "SELECT P.ID_PROVEEDOR, P.NOMBRE, P.DIRECCION, P.COLONIA, P.CIUDAD, P.CP, P.RFC, P.TELEFONO1, P.TELEFONO2, P.TELEFONO3, P.NOTAS, P.ESTADO, P.PAIS FROM PROVEEDOR AS P JOIN COTIZA_REQUI AS C ON C.ID_PROVEEDOR = P.ID_PROVEEDOR WHERE P.ELIMINADO = 'N' AND ESTADO_ACTUAL = 'A' ORDER BY P.NOMBRE"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Me.lvwProveedores.ListItems.Clear
            Do While Not .EOF
                If nId_Proveedor <> .Fields("ID_PROVEEDOR") Then
                    Set tLi = lvwProveedores.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                    If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(.Fields("DIRECCION"))
                    If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(.Fields("COLONIA"))
                    If Not IsNull(.Fields("CP")) Then tLi.SubItems(4) = Trim(.Fields("CP"))
                    If Not IsNull(.Fields("RFC")) Then tLi.SubItems(5) = Trim(.Fields("RFC"))
                    If Not IsNull(.Fields("TELEFONO1")) Then tLi.SubItems(6) = Trim(.Fields("TELEFONO1"))
                    If Not IsNull(.Fields("TELEFONO2")) Then tLi.SubItems(7) = Trim(.Fields("TELEFONO2"))
                    If Not IsNull(.Fields("TELEFONO3")) Then tLi.SubItems(8) = Trim(.Fields("TELEFONO3"))
                    If Not IsNull(.Fields("NOTAS")) Then tLi.SubItems(9) = Trim(.Fields("NOTAS"))
                    If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = Trim(.Fields("ESTADO"))
                    If Not IsNull(.Fields("PAIS")) Then tLi.SubItems(11) = Trim(.Fields("PAIS"))
                    'INICIO PARA NO REPETIR PROVEEDORES
                    nId_Proveedor = .Fields("ID_PROVEEDOR")
                    ' FIN
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
Function Hay_Proveedores() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_PROVEEDOR)ID_PROVEEDOR FROM PROVEEDOR WHERE ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_PROVEEDOR") <> 0 Then
            Hay_Proveedores = True
        Else
            Hay_Proveedores = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwCotizaciones_Click()
    Me.txtDias.SetFocus
End Sub
Private Sub lvwCotizaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.lblProducto.Caption = Item.SubItems(3)
    Me.txtCotizacion.Text = Item
    Me.txtRequisicion.Text = Item.SubItems(1)
    txtIndex.Text = Item.Index
    txtDias.Text = Item.SubItems(6)
    txtPrecio.Text = Item.SubItems(7)
    sDescripcion = Item.SubItems(4)
    sCantidad = Item.SubItems(5)
    If txtPrecio.Text = "" Or txtPrecio.Text = "0" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT TOP 1 PRECIO FROM VsComprasProveedor WHERE ID_PROVEEDOR = " & txtProveedor & " AND ID_PRODUCTO = '" & lvwCotizaciones.SelectedItem.SubItems(3) & "' ORDER BY FECHA DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("PRECIO")) Then txtPrecio = tRs.Fields("PRECIO")
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwCotizaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.txtDias.SetFocus
    End If
End Sub
Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If Hay_Cotizaciones Then
        Llenar_Lista_Cotizaciones (Item)
    End If
    Me.lblProveedor.Caption = Item.SubItems(1)
    Me.txtProveedor.Text = Item
    Me.lblProducto.Caption = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCotizacion_GotFocus()
    Me.txtCotizacion.BackColor = &HFFE1E1
End Sub
Private Sub txtCotizacion_LostFocus()
    txtCotizacion.BackColor = &H80000005
End Sub
Private Sub txtDias_GotFocus()
    Me.txtDias.BackColor = &HFFE1E1
End Sub
Private Sub txtDias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPrecio.SetFocus
    End If
End Sub
Private Sub txtDias_LostFocus()
    txtDias.BackColor = &H80000005
End Sub
Function Puede_Agregar() As Boolean
On Error GoTo ManejaError
    If Trim(Val(Me.txtCotizacion.Text)) = 0 Then
        MsgBox "SELECCIONE EL ARTICULO", vbInformation, "SACC"
        Me.lvwCotizaciones.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    If Val(Me.txtDias.Text) = 0 Then
        MsgBox "ESCRIBA LOS DIAS DE ENTREGA", vbInformation, "SACC"
        Me.txtDias.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    If Val(Me.txtPrecio.Text) = 0 Then
        MsgBox "ESCRIBA EL PRECIO", vbInformation, "SACC"
        Me.txtPrecio.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    If Trim(Val(Me.txtCotizacion.Text)) = 0 Then
        MsgBox "SELECCIONE EL ARTICULO", vbInformation, "SACC"
        Puede_Agregar = False
        Exit Function
    End If
    If Combo1.Text = "" Then
        MsgBox "SELECCIONE LA MONEDA EN LA QUE ESTA EXPRESADA LA COTIZACION", vbInformation, "SACC"
        Puede_Agregar = False
        Exit Function
    End If
    Puede_Agregar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub txtPrecio_GotFocus()
    Me.txtPrecio.BackColor = &HFFE1E1
End Sub
Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdAgregar.Value = True
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
Private Sub txtPrecio_LostFocus()
    txtPrecio.BackColor = &H80000005
End Sub
Private Sub txtProveedor_GotFocus()
    Me.txtProveedor.BackColor = &HFFE1E1
End Sub
Private Sub txtProveedor_LostFocus()
    txtProveedor.BackColor = &H80000005
End Sub
Private Sub txtRequisicion_GotFocus()
    Me.txtRequisicion.BackColor = &HFFE1E1
End Sub
Private Sub txtRequisicion_LostFocus()
    txtRequisicion.BackColor = &H80000005
End Sub
