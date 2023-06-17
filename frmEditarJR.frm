VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEditarJR 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDITAR JUEGO DE REPARACIÓN"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   9
      Top             =   4080
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmEditarJR.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmEditarJR.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmEditarJR.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblID_Rep"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEstado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblJR"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwJR"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCantidad"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAlmacen"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtId_Producto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSeleccionar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
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
         Left            =   6840
         Picture         =   "frmEditarJR.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtId_Producto 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   4560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtAlmacen 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   4560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   4560
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwJR 
         Height          =   3615
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6376
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
      Begin VB.Label lblJR 
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
         TabIndex        =   8
         Top             =   240
         Width           =   7815
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
         Left            =   240
         TabIndex        =   7
         Top             =   4320
         Width           =   5415
      End
      Begin VB.Label lblID_Rep 
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   4440
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEditarJR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim tRs2 As ADODB.Recordset
Private Sub cmdSeleccionar_Click()
On Error GoTo ManejaError
    If lvwJR.SelectedItem.Selected Then
        txtId_Producto.Text = lvwJR.SelectedItem
        txtAlmacen.Text = lvwJR.SelectedItem.SubItems(4)
        txtCantidad.Text = lvwJR.SelectedItem.SubItems(2)
        frmJRAlterno.txtCantidad.Text = frmEditarJR.txtCantidad.Text
        frmJRAlterno.lblProducto.Caption = lblID_Rep.Caption
        frmJRAlterno.lblPieza.Caption = frmEditarJR.txtId_Producto.Text
        frmJRAlterno.lblAlmacen.Caption = frmEditarJR.txtAlmacen.Text
        frmJRAlterno.Show vbModal, Me
    Else
        MsgBox "NO SE SELECCIONO NINGUN PRODUCTO", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Activate()
    Me.lblJR.Caption = "Juego de Reparación de " & frmReviComa.txtId_Reparacion.Text
    Me.lblEstado.Caption = "Buscando Juego de Reparación " & frmReviComa.txtId_Reparacion.Text
    DoEvents
    Llenar_Lista_Juego_Reparacion frmReviComa.txtId_Reparacion.Text
    Traer_Existencias
    Traer_Descripciones
    Me.lblEstado.Caption = ""
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
    With lvwJR
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4100
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Almacen", 0
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub Llenar_Lista_Juego_Reparacion(Id_Reparacion As String)
On Error GoTo ManejaError
    'Si no ha sido agregada la tabla temporal
    lblID_Rep.Caption = Id_Reparacion
    sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & frmReviComa.txtId_Comanda.Text & " AND ID_REPARACION = '" & Id_Reparacion & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.Fields("TEMPORAL") = 0 Then
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Id_Reparacion & "'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwJR.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwJR.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD")
                    'Insertar datos en tabla temporal
                    sqlQuery = "INSERT INTO JR_TEMPORALES (ID_COMANDA, ID_REPARACION, ID_PRODUCTO, CANTIDAD) VALUES (" & frmReviComa.txtId_Comanda.Text & ", '" & Trim(frmReviComa.txtId_Reparacion.Text) & "', '" & .Fields("ID_PRODUCTO") & "', " & .Fields("CANTIDAD") & ")"
                    cnn.Execute (sqlQuery)
            .MoveNext
            Loop
        End With
    Else
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_COMANDA = " & frmReviComa.txtId_Comanda.Text & " AND ID_REPARACION = '" & Id_Reparacion & "'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwJR.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwJR.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD")
            .MoveNext
            Loop
        End With
    End If

Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub Traer_Existencias()
On Error GoTo ManejaError
    Dim NoRe As Integer
    Dim Cont As Integer
    NoRe = Me.lvwJR.ListItems.Count
    For Cont = 1 To NoRe
        Me.lblEstado.Caption = "Buscando existencaias de " & Me.lvwJR.ListItems.Item(Cont)
        DoEvents
        sqlQuery = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Me.lvwJR.ListItems.Item(Cont) & "'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not .EOF And Not .BOF Then
                If Not IsNull(.Fields("CANTIDAD")) Then Me.lvwJR.ListItems(Cont).SubItems(3) = .Fields("CANTIDAD")
            Else
                Me.lvwJR.ListItems(Cont).SubItems(3) = "-"
            End If
        End With
    Next Cont
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub Traer_Descripciones()
On Error GoTo ManejaError
    Dim NoRe As Integer
    Dim Cont As Integer
    NoRe = Me.lvwJR.ListItems.Count
    For Cont = 1 To NoRe
        Me.lblEstado.Caption = "Buscando Descripcion de " & Me.lvwJR.ListItems.Item(Cont)
        DoEvents
        sqlQuery = "SELECT Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Me.lvwJR.ListItems.Item(Cont) & "'"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not .EOF And Not .BOF Then
                If Not IsNull(.Fields("Descripcion")) Then Me.lvwJR.ListItems(Cont).SubItems(1) = .Fields("Descripcion")
                Me.lvwJR.ListItems(Cont).SubItems(4) = "2"
            Else
                sqlQuery = "SELECT Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Me.lvwJR.ListItems.Item(Cont) & "'"
                Set tRs2 = cnn.Execute(sqlQuery)
                If Not tRs2.EOF And Not tRs2.BOF Then
                    If Not IsNull(.Fields("Descripcion")) Then Me.lvwJR.ListItems(Cont).SubItems(1) = tRs2.Fields("Descripcion")
                    Me.lvwJR.ListItems(Cont).SubItems(4) = "1"
                Else
                    Me.lvwJR.ListItems(Cont).SubItems(1) = "-"
                End If
                
            End If
        End With
    Next Cont
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
    frmReviComa.cmdTraer.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwJR_DblClick()
On Error GoTo ManejaError
    Me.cmdSeleccionar.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
