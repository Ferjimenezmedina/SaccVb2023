VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOrdenesProduccion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Producción"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   16
      Top             =   6000
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmOrdenesProduccion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmOrdenesProduccion.frx":030A
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   14
      Top             =   4680
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmOrdenesProduccion.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "frmOrdenesProduccion.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Productos"
      TabPicture(0)   =   "frmOrdenesProduccion.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEstado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvwProductosComanda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "opnClaveComanda"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "opnDescripcionComanda"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProductoComanda"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCantidadComanda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdBuscarComanda"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdQuitarComanda"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAgregarComanda"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "SSTab2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   6720
         Width           =   6015
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Productos"
         TabPicture(0)   =   "frmOrdenesProduccion.frx":40D4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lvwNuevaComanda"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Faltantes"
         TabPicture(1)   =   "frmOrdenesProduccion.frx":40F0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lfalt"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView lvwNuevaComanda 
            Height          =   1695
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lfalt 
            Height          =   1695
            Left            =   -74760
            TabIndex        =   20
            Top             =   360
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdAgregarComanda 
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
         Left            =   7320
         Picture         =   "frmOrdenesProduccion.frx":410C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitarComanda 
         Caption         =   "Quitar"
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
         Left            =   7320
         Picture         =   "frmOrdenesProduccion.frx":6ADE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarComanda 
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
         Left            =   7320
         Picture         =   "frmOrdenesProduccion.frx":94B0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCantidadComanda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtProductoComanda 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   5175
      End
      Begin VB.OptionButton opnDescripcionComanda 
         Caption         =   "Descripción"
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
         Left            =   5640
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton opnClaveComanda 
         Caption         =   "Clave"
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
         Left            =   5640
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwProductosComanda 
         Height          =   2055
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario:"
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
         TabIndex        =   22
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Left            =   5040
         TabIndex        =   10
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda"
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
         TabIndex        =   7
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblEstado 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   6120
         Width           =   9135
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmOrdenesProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim juegore As String
Dim nCantidad As Double
Dim cId_Producto As String
Private Sub Image8_Click()
On Error GoTo ManejaError
    If Puede_Guardar Then
        If Text1.Text = "" Then
            MsgBox "INGRESE UN COMENTARIO PARA PODER CERRAR  COMANDA!", vbInformation, "SACC"
        Else
            Dim NoRe As Integer
            Dim Cont As Integer
            Dim nComanda As Integer
            Dim cTipo As String
            sqlQuery = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_AGENTE, ID_SUCURSAL, TIPO, COMENTARIO, SUCURSAL) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & VarMen.Text1(0).Text & ", " & VarMen.Text1(5).Text & ", 'P','" & Text1.Text & "', '" & VarMen.Text4(0).Text & "')"
            cnn.Execute (sqlQuery)
            Me.lblEstado.Caption = "Enviando"
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            sqlQuery = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
            Set tRs = cnn.Execute(sqlQuery)
            nComanda = tRs.Fields("ID_COMANDA")
            Me.lblEstado.Caption = Me.lblEstado.Caption & " comanda " & nComanda
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            NoRe = Me.lvwNuevaComanda.ListItems.Count
            For Cont = 1 To NoRe
                If Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "T" Then
                    cTipo = "T" 'Toner
                Else
                    If Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "I" Then
                        cTipo = "I" 'Tinta
                    Else
                        cTipo = "X" 'Error
                    End If
                End If
                sqlQuery = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO,CLASIFICACION) VALUES (" & nComanda & ", " & Cont & ", '" & Me.lvwNuevaComanda.ListItems.Item(Cont) & "', " & Me.lvwNuevaComanda.ListItems.Item(Cont).SubItems(2) & ", '" & cTipo & "','P');"
                cnn.Execute (sqlQuery)
                sqlQuery = "INSERT INTO PRODPEND (ID_COMANDA, ARTICULO) VALUES (" & nComanda & ", " & Cont & ");"
                cnn.Execute (sqlQuery)
                Me.lblEstado.Caption = Me.lblEstado.Caption & ", producto " & Cont & " de " & NoRe
                Me.lblEstado.ForeColor = vbBlack
                DoEvents
            Next Cont
            Imprimir_Ticket (nComanda)
            Imprimir_Ticket (nComanda)
            Borrar_Campos
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdAgregarComanda_Click()
On Error GoTo ManejaError
    If Label4.Caption <> "" Then
        If txtCantidadComanda.Text <> "" Then
            If txtCantidadComanda <> 0 Then
                Set tLi = Me.lvwNuevaComanda.ListItems.Add(, , Me.lvwProductosComanda.SelectedItem)
                tLi.SubItems(1) = Me.lvwProductosComanda.SelectedItem.SubItems(1)
                tLi.SubItems(2) = Me.txtCantidadComanda.Text
                Me.lblEstado.Caption = ""
                Me.txtProductoComanda.SetFocus
                'faltantes
                nCantidad = txtCantidadComanda.Text
                Hay_existencias
            Else
                MsgBox "NO PUEDE HACER PEDIDOS EN CEROS!", vbInformation, "SACC"
            End If
        Else
            MsgBox "NO HA DADO UNA CANTIDAD A PEDIR!", vbInformation, "SACC"
        End If
    Else
        MsgBox "NO HA SELECCIONADO EL ARTICULO A PEDIR!", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Hay_existencias()
On Error GoTo ManejaError
    Dim tRs3 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim cMin As Double
    Dim sqlQuery As String
    Dim Pedido As Boolean
    bBandExis = False
    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Label4.Caption & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
                Pedido = True
                sqlQuery = "SELECT ID_EXISTENCIA, ID_PRODUCTO, CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = 'BODEGA'"
                Set tRs2 = cnn.Execute(sqlQuery)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    If Not (tRs2.Fields("CANTIDAD") >= (tRs.Fields("CANTIDAD") * nCantidad)) Then
                        'NO HAY SUFICIENTE EXISTENCIA
                        Set tLi = Me.lfalt.ListItems.Add(, , cId_Producto)
                            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            tLi.SubItems(2) = ((tRs.Fields("CANTIDAD") * (nCantidad)) - (CDbl(tRs2.Fields("CANTIDAD"))))
                        bBandExis = True
                    End If
                Else
                    'NO HAY REGISTRO EN LA TABLA
                    Set tLi = Me.lfalt.ListItems.Add(, , cId_Producto)
                    tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                    tLi.SubItems(2) = tRs.Fields("CANTIDAD") * nCantidad
                    bBandExis = True
                End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBuscarComanda_Click()
    If Puede_Buscar_Producto Then
        Me.lblEstado.Caption = "Buscando"
        Me.lblEstado.ForeColor = vbBlack
        DoEvents
        Llenar_Lista_Productos Trim(Me.txtProductoComanda.Text)
    End If
End Sub
Private Sub cmdQuitarComanda_Click()
    If Me.lvwNuevaComanda.ListItems.Count <> 0 Then
        If Me.lvwNuevaComanda.SelectedItem.Selected Then
            Me.lvwNuevaComanda.ListItems.Remove (Me.lvwNuevaComanda.SelectedItem.Index)
        End If
    End If
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
    With lvwProductosComanda
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "Descripcion", 4600
        .ColumnHeaders.Add , , "GANACIA", 0
        .ColumnHeaders.Add , , "PRECIO COSTO", 0
        .ColumnHeaders.Add , , "PRECIO", 1000
    End With
    With lvwNuevaComanda
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "Descripcion", 4600
        .ColumnHeaders.Add , , "CANTIDAD", 1000
    End With
    With lfalt
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "JUEGO DE REPARACION", 2000
        .ColumnHeaders.Add , , "ID_PRODUCTO", 4600
         .ColumnHeaders.Add , , "cantidad", 4600
          .ColumnHeaders.Add , , "existencia", 4600
        .ColumnHeaders.Add , , "FALTANTE", 1000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwProductosComanda_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwProductosComanda.ListItems.Count > 0 Then
        Label4.Caption = Item
    End If
End Sub
Private Sub lvwProductosComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.lvwProductosComanda.SelectedItem.Selected Then
            Me.txtProductoComanda.Text = lvwProductosComanda.SelectedItem
            Label4.Caption = lvwProductosComanda.SelectedItem
            Me.txtCantidadComanda.SetFocus
        End If
    End If
End Sub
Private Sub txtCantidadComanda_GotFocus()
    Me.txtCantidadComanda.SelStart = 0
    Me.txtCantidadComanda.SelLength = Len(Me.txtCantidadComanda.Text)
End Sub
Private Sub txtCantidadComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.cmdAgregarComanda.Value = True
    Else
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub txtProductoComanda_GotFocus()
    Me.txtProductoComanda.SelStart = 0
    Me.txtProductoComanda.SelLength = Len(Me.txtProductoComanda.Text)
End Sub
Private Sub txtProductoComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lvwProductosComanda.SetFocus
        Me.cmdBuscarComanda.Value = True
    End If
End Sub
Function Llenar_Lista_Productos(cProducto As String)
On Error GoTo ManejaError
    If Me.opnClaveComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    ElseIf Me.opnDescripcionComanda.Value = True Then
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE A.Descripcion LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    Else
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.Descripcion, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION = '" & cProducto & "' ORDER BY J.ID_REPARACION"
    End If
        Set tRs = cnn.Execute(sqlQuery)
        Me.lvwProductosComanda.ListItems.Clear
        With tRs
            If Not (.EOF And .BOF) Then
                Me.lblEstado.Caption = ""
                Do While Not .EOF
                    Set tLi = lvwProductosComanda.ListItems.Add(, , .Fields("ID_REPARACION"))
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                    If Not IsNull(.Fields("GANANCIA")) Then tLi.SubItems(2) = .Fields("GANANCIA")
                    If Not IsNull(.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                    If Not IsNull(.Fields("PRECIO_COSTO")) And Not IsNull(.Fields("GANANCIA")) Then
                        tLi.SubItems(4) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "###,###,##0.00")
                    Else
                        MsgBox "BASE DE DATOS CORRUPTA", vbCritical, "ERROR GRAVE"
                    End If
                    .MoveNext
                Loop
            Else
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtProductoComanda.SetFocus
            End If
        End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Puede_Guardar() As Boolean
    If Me.lvwNuevaComanda.ListItems.Count = 0 Then
        Puede_Guardar = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    Puede_Guardar = True
End Function
Function Imprimir_Ticket(cNoCom As Integer)
On Error GoTo ManejaError
    Printer.Print "        " & VarMen.Text5(0).Text
    Printer.Print "           ORDEN DE PRODUCCIÓN"
    Printer.Print "FECHA : " & Now
    Printer.Print "No. DE ORDEN DE PRODUCCCION : " & cNoCom
    Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "COMENTARIO : " & Text1.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TINTA"
    Dim NRegistros As Integer
    NRegistros = Me.lvwNuevaComanda.ListItems.Count
    Dim Con As Integer
    Dim POSY As Integer
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For Con = 1 To NRegistros
        If Mid(Me.lvwNuevaComanda.ListItems.Item(Con), 3, 1) = "I" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(Con).SubItems(2)
        End If
    Next Con
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TONER"
    POSY = POSY + 600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For Con = 1 To NRegistros
        If Mid(Me.lvwNuevaComanda.ListItems.Item(Con), 3, 1) = "T" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(Con).SubItems(2)
        End If
    Next Con
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.EndDoc
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Borrar_Campos()
    Me.txtCantidadComanda.Text = "1"
    Me.txtProductoComanda.Text = ""
    Me.lvwProductosComanda.ListItems.Clear
    Me.lvwNuevaComanda.ListItems.Clear
    Me.lblEstado.Caption = ""
End Function
Function Puede_Buscar_Producto() As Boolean
    If Trim(Me.txtProductoComanda.Text) = "" Then
        Puede_Buscar_Producto = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    Puede_Buscar_Producto = True
End Function
