VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAutorizarCotizaciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELECCIONAR COTIZACIONES"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   18
      Top             =   5760
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmAutorizarCotizaciones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAutorizarCotizaciones.frx":030A
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cotizaciones"
      TabPicture(0)   =   "frmAutorizarCotizaciones.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtRequisicion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCotizacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   7095
         Begin MSComctlLib.ListView lvwRequisiciones 
            Height          =   1455
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   2566
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   7095
         Begin MSComctlLib.ListView lvwCotizaciones 
            Height          =   2055
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   5160
         Width           =   7095
         Begin VB.CommandButton cdmAutorizar 
            Caption         =   "Autorización"
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
            Picture         =   "frmAutorizarCotizaciones.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Rechazar"
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
            Left            =   3120
            Picture         =   "frmAutorizarCotizaciones.frx":4DDA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Devolver"
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
            Left            =   1680
            Picture         =   "frmAutorizarCotizaciones.frx":77AC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCant 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4440
            Picture         =   "frmAutorizarCotizaciones.frx":A17E
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   585
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtInd 
            Height          =   285
            Left            =   5760
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDolar 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label lblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   6735
         End
         Begin VB.Label Label1 
            Caption         =   "Cantidad del producto:"
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   6240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRequisicion 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   6240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblProd 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmAutorizarCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim Cont As Integer
Dim NoRe As Integer
Dim IdProd As String
Private Sub cdmAutorizar_Click()
On Error GoTo ManejaError
    If Me.txtCotizacion.Text = "" Then
        MsgBox "SELECCIONE LA COTIZACION AUTORIZADA", vbInformation, "SACC"
    Else
        sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'X', CANTIDAD = " & txtCant.Text & " WHERE ID_COTIZACION IN (" & Me.txtCotizacion.Text & ") AND ID_PRODUCTO = '" & IdProd & "'"
        cnn.Execute (sqlQuery)
        sqlQuery = "UPDATE REQUISICION SET ACTIVO = '1' WHERE ID_REQUISICION IN (" & Me.txtRequisicion.Text & ") AND ID_PRODUCTO = '" & IdProd & "'"
        cnn.Execute (sqlQuery)
        NoRe = Me.lvwCotizaciones.ListItems.Count
        For Cont = 1 To NoRe
            If Me.lvwCotizaciones.ListItems.Item(Cont) <> Me.txtCotizacion.Text Then
                sqlQuery = "DELETE FROM COTIZA_REQUI WHERE ESTADO_ACTUAL <> 'X' AND ID_COTIZACION IN (" & lvwCotizaciones.ListItems.Item(Cont) & ") AND ID_PRODUCTO = '" & IdProd & "'"
                cnn.Execute (sqlQuery)
            End If
        Next Cont
        txtCotizacion.Text = ""
        lvwCotizaciones.ListItems.Clear
        Llenar_Lista_Requisiciones
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
    On Error GoTo ManejaError
    If Me.txtRequisicion.Text = "" Then
        MsgBox "SELECCIONE LA REQUISICION CANCELADA", vbInformation, "SACC"
    Else
        NoRe = Me.lvwCotizaciones.ListItems.Count
        For Cont = 1 To NoRe
            sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'Z' WHERE ID_COTIZACION IN (" & lvwCotizaciones.ListItems.Item(NoRe) & ")"
            cnn.Execute (sqlQuery)
        Next Cont
        sqlQuery = "UPDATE REQUISICION SET ACTIVO = '1' WHERE ID_REQUISICION IN (" & Me.txtRequisicion.Text & ")"
        Set tRs = cnn.Execute(sqlQuery)
        Me.lvwCotizaciones.ListItems.Clear
        Llenar_Lista_Requisiciones
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    If Command3.Caption = "Cambiar" Then
        txtCant.Enabled = True
        txtCant.SetFocus
        Command3.Caption = "Guardar"
        Command2.Enabled = False
        Command1.Enabled = False
        cdmAutorizar.Enabled = False
    Else
        If txtCant.Text <> "" Then
            lvwRequisiciones.ListItems.Item(CDbl(txtInd.Text)).SubItems(3) = txtCant.Text
            Command3.Caption = "Cambiar"
            Command2.Enabled = True
            Command1.Enabled = True
            cdmAutorizar.Enabled = True
            txtCant.Enabled = False
        End If
    End If
End Sub
Private Sub Form_Activate()
On Error GoTo ManejaError
    If Hay_Requisiciones Then
        Llenar_Lista_Requisiciones
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
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With Me.lvwRequisiciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
        .ColumnHeaders.Add , , "CONTADOR", 0, 2
    End With
    With Me.lvwCotizaciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 2300
        .ColumnHeaders.Add , , "Clave del Producto", 0
        .ColumnHeaders.Add , , "Descripcion", 0
        .ColumnHeaders.Add , , "CANTIDAD", 0, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 1200, 2
        .ColumnHeaders.Add , , "PRECIO", 1200, 2
        .ColumnHeaders.Add , , "FECHA", 2000, 2
    End With
    sBusca = "SELECT TOP 1 isnull(VENTA, 0) AS VENTA FROM DOLAR ORDER BY FECHA"
    Set tRs = cnn.Execute(sBusca)
    If Not (tRs.EOF And tRs.BOF) Then
        If tRs.Fields("VENTA") = 0 Then
            txtDolar.Text = 1
            MsgBox "NO EXISTEN REGISTROS DEL TIPO DE CAMBIO DEL DOLAR", vbInformation, "SACC"
        Else
            txtDolar.Text = tRs.Fields("VENTA")
        End If
    Else
        txtDolar.Text = 1
        MsgBox "NO EXISTEN REGISTROS DEL TIPO DE CAMBIO DEL DOLAR", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Requisiciones()
On Error GoTo ManejaError
    sqlQuery = "SELECT STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM COTIZA_REQUI RQ WHERE RQ.ID_PRODUCTO = R.ID_PRODUCTO AND (ESTADO_ACTUAL = 'A')  FOR XML PATH('')), 1, 1, '') AS ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD, FECHA, 1 AS CONTADOR FROM REQUISICION AS R WHERE (ACTIVO = 0) AND (COTIZADA = 1) AND (ID_PRODUCTO IN (SELECT ID_PRODUCTO From COTIZA_REQUI WHERE (ID_PRODUCTO = R.ID_PRODUCTO) AND (ID_REQUISICION = R.ID_REQUISICION) AND (ESTADO_ACTUAL = 'A'))) GROUP BY ID_PRODUCTO, DESCRIPCION, FECHA"
    'sqlQuery = "SELECT ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, CONTADOR FROM REQUISICION AS R WHERE (ACTIVO = 0) AND (COTIZADA = 1) AND (ID_PRODUCTO IN (SELECT ID_PRODUCTO From COTIZA_REQUI WHERE (ID_PRODUCTO = R.ID_PRODUCTO) AND (ID_REQUISICION = R.ID_REQUISICION) AND (ESTADO_ACTUAL = 'A')))"
    Set tRs = cnn.Execute(sqlQuery)
    Me.lvwRequisiciones.ListItems.Clear
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwRequisiciones.ListItems.Add(, , .Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(4) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("CONTADOR")) Then
                    tLi.SubItems(5) = Trim(.Fields("CONTADOR"))
                Else
                    tLi.SubItems(5) = "0"
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
Function Hay_Requisiciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_REQUISICION)ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_REQUISICION") <> 0 Then
            Hay_Requisiciones = True
        Else
            Hay_Requisiciones = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image9_Click()
    On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwCotizaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtCotizacion.Text = Item
    Me.lblProveedor.Caption = Item.SubItems(3)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwRequisiciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If Hay_Cotizaciones Then
        txtInd.Text = Item.Index
        txtCant.Text = Item.SubItems(3)
        txtRequisicion.Text = Item
        lblProd.Caption = Item.SubItems(1) & " " & Item.SubItems(2)
        IdProd = Item.SubItems(1)
        Llenar_Lista_Cotizaciones Item.SubItems(1)
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Cotizaciones() As Boolean
On Error GoTo ManejaError
    Hay_Cotizaciones = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Cotizaciones(cId_Producto As String)
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim CONT2 As Integer
    Dim Precio As Integer
    Dim IDC As String
    Dim IDR As String
    Dim IDP As String
    Dim De As String
    Dim Nombre As String
    Dim fecha As String
    sqlQuery = "SELECT STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_COTIZACION) FROM COTIZA_REQUI CR WHERE CR.FOLIO = C.FOLIO  AND (C.ESTADO_ACTUAL = 'A') AND (C.ID_PRODUCTO = '" & cId_Producto & "') AND (C.PRECIO <> 0) FOR XML PATH('')), 1, 1, '') AS ID_COTIZACION, ('" & txtRequisicion.Text & "') AS ID_REQUISICION, P.NOMBRE, C.ID_PROVEEDOR, C.ID_PRODUCTO, C.DESCRIPCION, C.CANTIDAD, C.DIAS_ENTREGA, C.PRECIO, C.FECHA, ISNULL(C.MONEDA, '0') AS MONEDA FROM COTIZA_REQUI AS C INNER JOIN PROVEEDOR AS P ON P.ID_PROVEEDOR = C.ID_PROVEEDOR WHERE (C.ESTADO_ACTUAL = 'A') AND (C.ID_PRODUCTO = '" & cId_Producto & "') AND (C.PRECIO <> 0) GROUP BY P.NOMBRE, C.ID_PROVEEDOR, C.ID_PRODUCTO, C.DESCRIPCION, C.CANTIDAD, C.DIAS_ENTREGA, C.PRECIO,  C.FECHA, MONEDA, FOLIO, ESTADO_ACTUAL ORDER BY C.PRECIO, C.DIAS_ENTREGA"
    'sqlQuery = "SELECT C.ID_COTIZACION, C.ID_REQUISICION, P.NOMBRE, C.ID_PROVEEDOR, C.ID_PRODUCTO, C.Descripcion, C.CANTIDAD, C.DIAS_ENTREGA, C.PRECIO, C.FECHA, ISNULL(C.MONEDA, '0') AS MONEDA FROM COTIZA_REQUI AS C JOIN PROVEEDOR AS P ON P.ID_PROVEEDOR= C.ID_PROVEEDOR WHERE C.ESTADO_ACTUAL = 'A' AND C.ID_PRODUCTO = '" & cId_Producto & "' AND C.PRECIO <> 0 AND ID_REQUISICION IN (" & txtRequisicion.Text & ") ORDER BY C.PRECIO, C.DIAS_ENTREGA"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwCotizaciones.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = Trim(.Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(5) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(6) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(7) = Trim(.Fields("DIAS_ENTREGA"))
                If .Fields("MONEDA") = "PESOS" Or .Fields("MONEDA") = "0" Then
                    If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(8) = Trim(.Fields("PRECIO"))
                Else
                    If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(8) = Val(.Fields("PRECIO")) * Val(txtDolar.Text)
                End If
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(9) = Trim(.Fields("FECHA"))
                .MoveNext
            Loop
        End If
    End With
    If lvwCotizaciones.ListItems.Count > 0 Then
        For Cont = 1 To lvwCotizaciones.ListItems.Count
            For CONT2 = Cont + 1 To lvwCotizaciones.ListItems.Count
                If CDbl(lvwCotizaciones.ListItems.Item(Cont).SubItems(8)) > CDbl(lvwCotizaciones.ListItems.Item(CONT2).SubItems(8)) Then
                    IDC = lvwCotizaciones.ListItems.Item(Cont)  'lvwCotizaciones.ListItems.Add(, , .Fields("ID_COTIZACION"))
                    IDR = lvwCotizaciones.ListItems.Item(Cont).SubItems(1) ' = Trim(.Fields("ID_REQUISICION"))
                    IDP = lvwCotizaciones.ListItems.Item(Cont).SubItems(2) '= Trim(.Fields("ID_PROVEEDOR"))
                    Nombre = lvwCotizaciones.ListItems.Item(Cont).SubItems(3) '= Trim(.Fields("NOMBRE"))
                    De = lvwCotizaciones.ListItems.Item(Cont).SubItems(7) '= Trim(.Fields("DIAS_ENTREGA"))
                    Precio = lvwCotizaciones.ListItems.Item(Cont).SubItems(8) '= Trim(.Fields("PRECIO"))
                    fecha = lvwCotizaciones.ListItems.Item(Cont).SubItems(9) '= Trim(.Fields("FECHA"))
                    lvwCotizaciones.ListItems.Item(Cont) = lvwCotizaciones.ListItems.Item(CONT2)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(1) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(1)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(2) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(2)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(3) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(3)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(7) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(7)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(8) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(8)
                    lvwCotizaciones.ListItems.Item(Cont).SubItems(9) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(9)
                    lvwCotizaciones.ListItems.Item(CONT2) = IDC
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(1) = IDR
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(2) = IDP
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(3) = Nombre
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(7) = De
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(8) = Precio
                    lvwCotizaciones.ListItems.Item(CONT2).SubItems(9) = fecha
                End If
            Next CONT2
        Next Cont
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCant_GotFocus()
    txtCant.BackColor = &HFFE1E1
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCant.Text <> "" Then
        Command3.Value = True
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
Private Sub txtCant_LostFocus()
    txtCant.BackColor = &H80000005
End Sub
