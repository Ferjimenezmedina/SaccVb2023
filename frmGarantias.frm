VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGarantias 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GARANTIAS"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   26
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
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmGarantias.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmGarantias.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmGarantias.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvwVentaDetalle"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvwGarantia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTipo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtComentario"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtIDDET"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdGarantia"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCantidad2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCantidad"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDescripcion"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPrecio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtProducto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdAgregar"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdOk"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtVenta"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      Begin VB.TextBox txtVenta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   6975
         Begin VB.Label lblNombre 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   17
            Top             =   240
            Width           =   5295
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdOk 
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
         Left            =   4440
         Picture         =   "frmGarantias.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
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
         Left            =   5880
         Picture         =   "frmGarantias.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtProducto 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtCantidad2 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGarantia 
         Caption         =   "Garantia"
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
         Left            =   5880
         Picture         =   "frmGarantias.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtIDDET 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtComentario 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   3840
         Width           =   3255
      End
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   3720
         TabIndex        =   4
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Venta"
         Height          =   195
         Left            =   2880
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Asistencia"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Comanda"
         Height          =   195
         Left            =   2880
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwGarantia 
         Height          =   1935
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwVentaDetalle 
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2778
         View            =   3
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
         Alignment       =   1  'Right Justify
         Caption         =   "Venta :"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción :"
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
         TabIndex        =   23
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "No hay producto seleccionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   3360
         Width           =   5535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   3840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItMx As ListItem, itmX2 As ListItem
Private cnn As ADODB.Connection
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    Me.lvwVentaDetalle.SetFocus
    If Val(Me.txtCantidad.Text) <> 0 Then
        If Me.lvwVentaDetalle.SelectedItem.Selected = True Then
            If (Val(Me.txtCantidad.Text)) <= Val(Me.txtCantidad2.Text) Then
                Me.lvwVentaDetalle.SelectedItem.SubItems(5) = Val(Me.lvwVentaDetalle.SelectedItem.SubItems(5)) + Val(Me.txtCantidad.Text)
                Set itmX2 = Me.lvwGarantia.ListItems.Add(, , Trim(Me.txtVenta.Text))
                itmX2.SubItems(1) = Trim(Me.txtProducto.Text)
                itmX2.SubItems(2) = Trim(Me.txtDescripcion.Text)
                itmX2.SubItems(3) = Val(Me.txtCantidad.Text)
                itmX2.SubItems(4) = (Val(Me.txtPrecio.Text) / Val(Me.txtCantidad2.Text))
                itmX2.SubItems(5) = CDbl(Me.txtCantidad.Text) * CDbl((Val(Me.txtPrecio.Text) / Val(Me.txtCantidad2.Text)))
                itmX2.SubItems(6) = Me.txtIDDET.Text
                itmX2.SubItems(7) = Me.txtComentario.Text
                itmX2.SubItems(8) = txttipo.Text
                Me.cmdGarantia.Enabled = True
            Else
                MsgBox "LA CANTIDAD DEBE SER MENOR", vbInformation, "SACC"
            End If
        Else
            MsgBox "DE DOBLE CLICK EN EN ARTICULO SELECCIONADO", vbInformation, "SACC"
        End If
    Else
        MsgBox "ESCRIBA LA CANTIDAD", vbInformation, "SACC"
        Me.txtCantidad.SetFocus
    End If
    
    Exit Sub
ManejaError:
    MsgBox "NO SE PUEDE AGREGAR", vbCritical, "MENSAJE "
    Err.Clear
End Sub
Private Sub cmdGarantia_Click()
    Dim sBuscar  As String
    Dim tRs As ADODB.Recordset
    Dim nID_GARANTIA As Integer
    Guardar_Garantia
    Imprimir_Ticket
    sBuscar = "SELECT TOP 1 ID_GARANTIA From GARANTIAS ORDER BY ID_GARANTIA DESC"
    Set tRs = cnn.Execute(sBuscar)
    nID_GARANTIA = tRs.Fields("ID_GARANTIA")
    sBuscar = "UPDATE VENTAS_DETALLE SET ID_GARANTIA = " & nID_GARANTIA & " WHERE ID_VENTA = " & Trim(Me.txtVenta.Text)
    cnn.Execute (sBuscar)
    cmdGarantia.Enabled = False
    lvwVentaDetalle.ListItems.Clear
    lvwGarantia.ListItems.Clear
    txtCantidad.Text = ""
    Label5.Caption = "No hay producto seleccionado"
    txtVenta.Text = ""
    txtComentario.Text = ""
    lblNombre.Caption = "---"
    txttipo.Text = ""
End Sub
Private Sub cmdOk_Click()
    If Trim(Me.txtVenta.Text) <> "" Then
        Llenar_Lista_Ventas_Detalles Trim(CDbl(Me.txtVenta.Text))
    Else
        If MsgBox("ESCRIBA EL NUMERO DE LA VENTA", vbInformation, "SACC") = vbOK Then
            Me.txtVenta.SetFocus
        End If
    End If
    Me.lvwVentaDetalle.SetFocus
End Sub
Sub Llenar_Lista_Ventas_Detalles(Clave As Double)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Option1.Value = True Then
        Me.lvwVentaDetalle.ListItems.Clear
        sBuscar = "SELECT V.ID_VENTA, D.ID_PRODUCTO, D.Descripcion, D.CANTIDAD, D.PRECIO_VENTA, D.IDDET, C.Nombre FROM VENTAS AS V JOIN VENTAS_DETALLE AS D ON V.ID_VENTA = D.ID_VENTA JOIN CLIENTE AS C ON V.ID_CLIENTE = C.ID_CLIENTE WHERE V.ID_VENTA = " & Clave & " and FACTURADO <> 2"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            Set ItMx = Me.lvwVentaDetalle.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then ItMx.SubItems(1) = Trim(tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then ItMx.SubItems(2) = Trim(tRs.Fields("Descripcion"))
            If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(3) = Trim(tRs.Fields("CANTIDAD"))
            If Not IsNull(tRs.Fields("Precio_Venta")) Then ItMx.SubItems(4) = Trim(tRs.Fields("Precio_Venta"))
            If Not IsNull(tRs.Fields("Nombre")) Then Me.lblNombre.Caption = Trim(tRs.Fields("Nombre"))
            If Not IsNull(tRs.Fields("IDDET")) Then ItMx.SubItems(6) = Trim(tRs.Fields("IDDET"))
            tRs.MoveNext
        Loop
    Else
        If Option2.Value = True Then
            sBuscar = "SELECT ID_VENTA FROM VENTAS_DETALLE WHERE NO_COM_AT = 'A" & txtVenta.Text & "'"
        Else
            sBuscar = "SELECT ID_VENTA FROM VENTAS_DETALLE WHERE NO_COM_AT = 'C" & txtVenta.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "SELECT ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, NOMBRE, IDDET FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA")
            If Not (tRs.EOF And tRs.BOF) Then
                txtVenta.Text = tRs.Fields("ID_VENTA")
                Do While tRs.EOF
                    Set ItMx = Me.lvwVentaDetalle.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                    If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then ItMx.SubItems(1) = Trim(tRs.Fields("ID_PRODUCTO"))
                    If Not IsNull(tRs.Fields("Descripcion")) Then ItMx.SubItems(2) = Trim(tRs.Fields("Descripcion"))
                    If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(3) = Trim(tRs.Fields("CANTIDAD"))
                    If Not IsNull(tRs.Fields("Precio_Venta")) Then ItMx.SubItems(4) = Trim(tRs.Fields("Precio_Venta"))
                    If Not IsNull(tRs.Fields("Nombre")) Then Me.lblNombre.Caption = Trim(tRs.Fields("Nombre"))
                    If Not IsNull(tRs.Fields("IDDET")) Then ItMx.SubItems(6) = Trim(tRs.Fields("IDDET"))
                    tRs.MoveNext
                Loop
            End If
        Else
            If Option2.Value = True Then
                MsgBox "NO SE ENCONTRO EL NUMERO DE ASISTENCIA!", vbExclamation, "SACC"
            Else
                MsgBox "NO SE ENCONTRO EL NUMERO DE COMANDA!", vbExclamation, "SACC"
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwVentaDetalle
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Ventas", 0
        .ColumnHeaders.Add , , "Producto", 1500
        .ColumnHeaders.Add , , "Descripcion", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio", 1500
        .ColumnHeaders.Add , , "Agregado", 0
        .ColumnHeaders.Add , , "IDDET", 0
    End With
    With lvwGarantia
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Ventas", 0
        .ColumnHeaders.Add , , "Producto", 1500
        .ColumnHeaders.Add , , "Descripcion", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio", 1000
        .ColumnHeaders.Add , , "Importe", 1500
        .ColumnHeaders.Add , , "IDDET", 1
        .ColumnHeaders.Add , , "Comentario", 1500
        .ColumnHeaders.Add , , "Tipo", 0
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwVentaDetalle_Click()
On Error GoTo ManejaEerror:
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Me.txtProducto.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(1)
    Me.txtDescripcion.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(2)
    Me.txtCantidad2.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(3)
    Me.txtPrecio.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(4)
    Me.txtIDDET.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(6)
    Label5.Caption = Me.lvwVentaDetalle.SelectedItem.SubItems(1) '"No hay producto seleccionado"
    Me.cmdAgregar.Enabled = False
    sBuscar = "SELECT TIPO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & lvwVentaDetalle.SelectedItem.SubItems(1) & "'"
    Set tRs = cnn.Execute(sBuscar)
    txttipo.Text = tRs.Fields("TIPO")
ManejaEerror:
    Err.Clear
End Sub
Private Sub lvwVentaDetalle_DblClick()
    Me.cmdAgregar.Enabled = True
    Me.txtProducto.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(1)
    Me.txtDescripcion.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(2)
    Me.txtCantidad2.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(3)
    Me.txtPrecio.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(4)
    Me.txtIDDET.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(6)
    Label5.Caption = Me.lvwVentaDetalle.SelectedItem.SubItems(2)
    Me.txtCantidad.SetFocus
End Sub
Private Sub lvwVentaDetalle_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError:
    If KeyAscii = 13 Then
        Me.txtProducto.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(1)
        Me.txtDescripcion.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(2)
        Me.txtCantidad2.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(3)
        Me.txtPrecio.Text = Me.lvwVentaDetalle.SelectedItem.SubItems(4)
        Me.cmdAgregar.Enabled = True
        Me.txtCantidad.SetFocus
    Else
        KeyAscii = 0
    End If
ManejaError:
    Err.Clear
End Sub
Private Sub TxtCantidad_Change()
    If (txtComentario.Text = "") Or (txtCantidad.Text = "") Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
    End If
End Sub
Private Sub txtCantidad_GotFocus()
    txtCantidad.BackColor = &HFFE1E1
End Sub
Private Sub TxtCantidad_LostFocus()
    txtCantidad.BackColor = &H80000005
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And (txtCantidad.Text <> "") Then
        txtComentario.SetFocus
    Else
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub txtComentario_Change()
    If (txtComentario.Text = "") Or (txtCantidad.Text = "") Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
    End If
End Sub
Private Sub txtComentario_GotFocus()
    txtComentario.BackColor = &HFFE1E1
End Sub
Private Sub txtComentario_LostFocus()
    txtComentario.BackColor = &H80000005
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And Me.cmdAgregar.Enabled Then
        Me.cmdAgregar.Value = True
    End If
End Sub
Private Sub txtVenta_GotFocus()
    txtVenta.BackColor = &HFFE1E1
    txtVenta.SelStart = 0
    txtVenta.SelLength = Len(txtVenta.Text)
End Sub
Private Sub txtVenta_LostFocus()
    txtVenta.BackColor = &H80000005
End Sub
Private Sub txtVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOk.Value = True
        Me.cmdOk.SetFocus
    Else
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Sub Guardar_Garantia()
    Dim CONTADOR As Integer
    Dim NUMERO_REGISTROS As Integer
    Dim Venta As Double
    Dim producto As String
    Dim CANTIDAD As Double
    Dim Precio As Double
    Dim importe As Double
    Dim Descripcion As String
    Dim COMENTARIO As String
    Dim Tipo As String
    Dim sBuscar As String
    NUMERO_REGISTROS = Me.lvwGarantia.ListItems.Count
    For CONTADOR = 1 To NUMERO_REGISTROS
        Venta = Me.lvwGarantia.ListItems.Item(CONTADOR)
        producto = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(1)
        CANTIDAD = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(3)
        Precio = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(4)
        importe = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(5)
        Descripcion = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(2)
        IDDET = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(6)
        COMENTARIO = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(7)
        Tipo = Me.lvwGarantia.ListItems.Item(CONTADOR).SubItems(8)
         sBuscar = "INSERT INTO GARANTIAS (ID_VENTA,FECHA,ID_PRODUCTO,CANTIDAD,PRECIO,IMPORTE,Descripcion,ESTADO,IDDET,COMENTARIO,TIPO) Values (" & Venta & ", '" _
        & Format(Date, "dd/mm/yyyy") & "', '" & producto & "', " & CANTIDAD & ", " & _
        Precio & ", " & importe & ", '" & Descripcion & "', 'P', " & IDDET & ", '" & COMENTARIO & "', '" & Tipo & "');"
        cnn.Execute (sBuscar)
    Next CONTADOR
End Sub
Sub Imprimir_Ticket()
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "No. DE VENTA : " & Me.txtVenta.Text
    Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text
    Printer.Print "CLIENTE : " & lblNombre.Caption
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                          TICKET DE GARANTIA"
    Printer.Print "--------------------------------------------------------------------------------"
    Dim NRegistros As Integer
    NRegistros = Me.lvwGarantia.ListItems.Count
    Dim Con As Integer
    Dim POSY As Integer
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 1300
    Printer.Print "Precio unitario"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For Con = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print Me.lvwGarantia.ListItems(Con).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 1900
        Printer.Print Me.lvwGarantia.ListItems(Con).SubItems(4)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2900
        Printer.Print Me.lvwGarantia.ListItems(Con).SubItems(3)
    Next Con
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print "                APLICA RESTRICCIONES"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
End Sub
