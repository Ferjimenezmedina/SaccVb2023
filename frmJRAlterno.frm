VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmJRAlterno 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JUEGO DE REPARACIÓN ALTERNO"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   22
      Top             =   6120
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmJRAlterno.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmJRAlterno.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmJRAlterno.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAlmacen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvwRemplasos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvwAlmacen"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtIdProducto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdBuscar2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBorrar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.Frame Frame4 
         Caption         =   "Cantidad Parcial"
         Height          =   1215
         Left            =   6840
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
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
            Left            =   120
            Picture         =   "frmJRAlterno.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Omitir"
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
         Left            =   6960
         Picture         =   "frmJRAlterno.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   8175
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   120
            Width           =   6495
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "REEMPLAZO:"
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
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8175
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "PRODUCTO:"
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
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "PIEZA:"
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
            Left            =   480
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   11
            Top             =   120
            Width           =   6615
         End
         Begin VB.Label lblPieza 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   480
            Width           =   6615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   6840
         TabIndex        =   8
         Top             =   5400
         Width           =   1455
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
            Left            =   120
            Picture         =   "frmJRAlterno.frx":77AC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
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
         Left            =   6960
         Picture         =   "frmJRAlterno.frx":A17E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar2 
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
         Left            =   6960
         Picture         =   "frmJRAlterno.frx":CB50
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtIdProducto 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Width           =   6615
      End
      Begin MSComctlLib.ListView lvwAlmacen 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3413
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
      Begin MSComctlLib.ListView lvwRemplasos 
         Height          =   1455
         Left            =   120
         TabIndex        =   0
         Top             =   2400
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2566
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
      Begin VB.Label lblAlmacen 
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
         Left            =   2520
         TabIndex        =   21
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Usados recientemente"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label lblEstado 
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
         Left            =   120
         TabIndex        =   19
         Top             =   6840
         Width           =   6495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Producto"
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
         Left            =   120
         TabIndex        =   18
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "PRODUCTOS DE ALMACEN"
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
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmJRAlterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim bBanderaLvw As Byte
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    If Puede_Agregar(Me.lvwAlmacen.SelectedItem.SubItems(1)) Then
        If bBanderaLvw = 1 Then
            sqlQuery = "UPDATE JR_TEMPORALES SET ID_PRODUCTO = '" & Trim(Me.lvwRemplasos.SelectedItem.SubItems(1)) & "', CANTIDAD = " & Me.txtCantidad.Text & " WHERE ID_COMANDA = " & frmReviComa.txtId_Comanda.Text & " AND ID_REPARACION = '" & Me.lblProducto.Caption & "' AND ID_PRODUCTO = '" & Me.lblPieza.Caption & "'"
            cnn.Execute (sqlQuery)
            If Me.Caption = "JUEGO DE REPARACION ALTERNO DE PRODUCTO DAÑADO" Then
                frmScrap.Text5(0).Text = Trim(Me.lvwRemplasos.SelectedItem.SubItems(1))
                frmScrap.Text5(1).Text = txtCantidad.Text
                Unload Me
            Else
                frmEditarJR.Llenar_Lista_Juego_Reparacion frmReviComa.txtId_Reparacion.Text
                frmEditarJR.Traer_Existencias
                frmEditarJR.Traer_Descripciones
                Unload Me
            End If
        Else
            sqlQuery = "INSERT INTO JR_ALTERNOS (ID_REPARACION, ID_PRODUCTO1, ID_PRODUCTO2, CANTIDAD) VALUES('" & Trim(Me.lblProducto.Caption) & "', '" & Trim(Me.lblPieza.Caption) & "', '" & Trim(Me.lvwAlmacen.SelectedItem) & "', " & Me.txtCantidad.Text & ")"
            cnn.Execute (sqlQuery)
            sqlQuery = "UPDATE JR_TEMPORALES SET ID_PRODUCTO = '" & Trim(Me.lvwAlmacen.SelectedItem) & "', CANTIDAD = " & Me.txtCantidad.Text & " WHERE ID_COMANDA = " & frmReviComa.txtId_Comanda.Text & " AND ID_REPARACION = '" & Me.lblProducto.Caption & "' AND ID_PRODUCTO = '" & Me.lblPieza.Caption & "'"
            cnn.Execute (sqlQuery)
            If Me.Caption = "JUEGO DE REPARACION ALTERNO DE PRODUCTO DAÑADO" Then
                frmScrap.Text5(0).Text = Trim(Me.lvwRemplasos.SelectedItem.SubItems(1))
                frmScrap.Text5(1).Text = txtCantidad.Text
                Unload Me
            Else
                frmEditarJR.Llenar_Lista_Juego_Reparacion frmReviComa.txtId_Reparacion.Text
                frmEditarJR.Traer_Existencias
                frmEditarJR.Traer_Descripciones
                frmEditarJR.lblEstado.Caption = ""
                Unload Me
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBorrar_Click()
On Error GoTo ManejaError
    If Me.lvwRemplasos.SelectedItem.Selected Then
        sqlQuery = "DELETE FROM JR_ALTERNOS WHERE ALTERNO = " & Me.lvwRemplasos.SelectedItem
        Set tRs = cnn.Execute(sqlQuery)
        Me.Llenar_Lista_Reemplazos Me.lblPieza.Caption, Me.lblProducto.Caption
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBuscar2_Click()
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando " & Me.txtIdProducto.Text & "... por favor espere..."
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    Me.Llenar_Lista_Almacen_Producto frmEditarJR.txtAlmacen.Text
    Me.Llenar_Lista_Existencias
    Me.lblEstado.Caption = "Listo, " & Me.lvwAlmacen.ListItems.Count & " registros encontrados"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    sBuscar = "DELETE FROM JR_TEMPORALES WHERE ID_PRODUCTO = '" & lblPieza.Caption & "' AND ID_COMANDA = " & frmReviComa.txtId_Comanda.Text
    cnn.Execute (sBuscar)
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim iAfectados As Long
    Dim tRs As ADODB.Recordset
    sBuscar = "UPDATE JR_TEMPORALES SET CANTIDAD = " & Text1.Text & " WHERE ID_PRODUCTO = '" & lblPieza.Caption & "' AND ID_COMANDA = " & frmReviComa.txtId_Comanda.Text
    Set tRs = cnn.Execute(sBuscar, iAfectados, adCmdText)
    If iAfectados = 0 Then
        sBuscar = "INSERT INTO JR_TEMPORALES (ID_REPARACION, ID_PRODUCTO, CANTIDAD, ID_COMANDA) VALUES ('" & Trim(Me.lblProducto.Caption) & "', '" & lblPieza.Caption & "', " & Text1.Text & ", " & frmReviComa.txtId_Comanda.Text & ");"
        cnn.Execute (sBuscar)
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    Llenar_Lista_Almacen lblAlmacen.Caption
    Llenar_Lista_Reemplazos lblPieza.Caption, lblProducto.Caption
    Llenar_Lista_Existencias
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwRemplasos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE REEMPLAZO", 0
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 3100
        .ColumnHeaders.Add , , "EXISTENCIA", 1000
    End With
    With lvwAlmacen
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 3100
        .ColumnHeaders.Add , , "EXISTENCIA", 1000
    End With
    sBuscar = "SELECT * FROM JR_TEMPORALES WHERE ID_REPARACION = '" & lblProducto.Caption & "' AND ID_COMANDA = " & frmReviComa.txtId_Comanda.Text
    Set tRs = cnn.Execute(sBuscar)
    For Cont = 1 To Cont > frmEditarJR.lvwJR.ListItems.Count
        sBuscar = "INSERT INTO JR_TEMPORALES (ID_REPARACION, ID_PRODUCTO, CANTIDAD, ID_COMANDA) VALUES ('" & lblProducto.Caption & "', '" & frmEditarJR.lvwJR.SelectedItem & "', " & frmEditarJR.lvwJR.SelectedItem.SubItems(2) * frmEditarJR.txtCantidad.Text & ", " & frmReviComa.txtId_Comanda.Text & ");"
        cnn.Execute (sBuscar)
    Next Cont
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Almacen(nAlmacen As Integer)
On Error GoTo ManejaError
    If nAlmacen = 1 Then
        sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwAlmacen.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwAlmacen.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End With
    Else
        sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwAlmacen.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwAlmacen.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Reemplazos(cId_Producto As String, cId_Reparacion As String)
On Error GoTo ManejaError
    sqlQuery = "SELECT ALTERNO, ID_PRODUCTO2, CANTIDAD FROM JR_ALTERNOS WHERE ID_REPARACION = '" & cId_Reparacion & "' AND ID_PRODUCTO1 = '" & cId_Producto & "'"
    Set tRs = cnn.Execute(sqlQuery)
    
    With tRs
        Me.lvwRemplasos.ListItems.Clear
        Do While Not .EOF
            Set tLi = Me.lvwRemplasos.ListItems.Add(, , .Fields("ALTERNO"))
            If Not IsNull(.Fields("ID_PRODUCTO2")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO2")
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwAlmacen_Click()
On Error GoTo ManejaError
    bBanderaLvw = 2
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwAlmacen_DblClick()
On Error GoTo ManejaError
    Me.txtCantidad.SetFocus
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Agregar(cId_Producto As String) As Boolean
On Error GoTo ManejaError
    If Me.txtCantidad.Text = "" Then
        MsgBox "TECLEE LA CANTIDAD", vbInformation, "SACC"
        Me.txtCantidad.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    Puede_Agregar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub lvwRemplasos_Click()
On Error GoTo ManejaError
    bBanderaLvw = 1
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwRemplasos_DblClick()
On Error GoTo ManejaError
    Me.txtCantidad.SetFocus
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.cmdAgregar.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCantidad_GotFocus()
    txtCantidad.BackColor = &HFFE1E1
End Sub
Private Sub TxtCantidad_LostFocus()
      txtCantidad.BackColor = &H80000005
End Sub
Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdBuscar2.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Almacen_Producto(nAlmacen As Integer)
On Error GoTo ManejaError
    If nAlmacen = 1 Then
        sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwAlmacen.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwAlmacen.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End With
    Else
        sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwAlmacen.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwAlmacen.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Existencias()
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Cargando existencias..."
    DoEvents
    Dim Cont As Integer
    Dim NoRe As Integer
    NoRe = Me.lvwAlmacen.ListItems.Count
    For Cont = 1 To NoRe
        sqlQuery = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Me.lvwAlmacen.ListItems.Item(Cont) & "' AND SUCURSAL = 'BODEGA'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.EOF And .BOF) Then
                Me.lvwAlmacen.ListItems.Item(Cont).SubItems(2) = .Fields("CANTIDAD")
            End If
        End With
    Next Cont
    NoRe = Me.lvwRemplasos.ListItems.Count
    For Cont = 1 To NoRe
        sqlQuery = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Me.lvwRemplasos.ListItems.Item(Cont) & "' AND SUCURSAL = 'BODEGA'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.EOF And .BOF) Then
                Me.lvwRemplasos.ListItems.Item(Cont).SubItems(2) = .Fields("CANTIDAD")
            End If
        End With
    Next Cont
    Me.lblEstado.Caption = "Listo"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
