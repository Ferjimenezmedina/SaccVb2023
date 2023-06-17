VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form EditarJRVarios 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Editar Juegos de Reparacion "
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   20
      Top             =   5640
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "EditarJRVarios.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "EditarJRVarios.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Buscar J.R."
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
      Picture         =   "EditarJRVarios.frx":23EC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reemplazar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   4335
      Begin VB.TextBox txtIDPROD2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
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
         Height          =   285
         Left            =   2640
         Picture         =   "EditarJRVarios.frx":4DBE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVReemplazo 
         Height          =   1095
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ninguno"
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
         Top             =   1920
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lista de Productos a Reemplazar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtIDPROD 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2415
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
         Height          =   285
         Left            =   2640
         Picture         =   "EditarJRVarios.frx":7790
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
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
         Left            =   2640
         Picture         =   "EditarJRVarios.frx":A162
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVProductos 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LVSeleccion 
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin TabDlg.SSTab SSTab1 
         Height          =   3885
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6853
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   503
         BackColor       =   16777215
         TabCaption(0)   =   "Juego de Reparacion"
         TabPicture(0)   =   "EditarJRVarios.frx":CB34
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LVIndividual"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIndex"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin VB.CommandButton Command2 
            Caption         =   "Quitar de Lista"
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
            Left            =   4080
            Picture         =   "EditarJRVarios.frx":CB50
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtIndex 
            Height          =   285
            Left            =   360
            TabIndex        =   17
            Top             =   3480
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.ListView LVIndividual 
            Height          =   2775
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4895
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sustituir Lista "
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
         Left            =   2400
         Picture         =   "EditarJRVarios.frx":F522
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6240
         Width           =   1455
      End
      Begin MSComctlLib.ListView LVJuegos 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Juegos de Reparacion"
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
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "EditarJRVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub CmdQuitar_Click()
    LVProductos.ListItems.Remove (LVProductos.SelectedItem.Index)
    If LVProductos.ListItems.Count = 0 Then
        cmdQuitar.Enabled = False
    End If
End Sub
Private Sub Command1_Click()
    Dim sBus As String
    Dim Lista As String
    Dim ListaTmp As String
    Dim NoCambiados As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim Inicio As Long
    Lista = ""
    For Cont = 1 To LVJuegos.ListItems.Count
        If ListaTmp <> "" Then
            ListaTmp = ListaTmp & ", "
        End If
        ListaTmp = ListaTmp & "'" & LVProductos.ListItems.Item(Cont) & "'"
    Next Cont
    sBus = "SELECT ID_PRODUCTO, CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO IN (" & ListaTmp & ")"
    Set tRs = cnn.Execute(sBus)
    NoCambiados = ""
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not (.EOF)
                If .Fields("CANTIDAD") > 0 Then
                    Inicio = InStr(ListaTmp, .Fields("ID_PRODUCTO"))
                    If NoCambiados <> "" Then
                        NoCambiados = NoCambiados & ", "
                    End If
                    NoCambiados = NoCambiados & Mid(ListaTmp, Inicio, Len(.Fields("ID_PRODUCTO")))
                    Lista = "'" & .Fields("ID_PRODUCTO") & "'"
                    ListaTmp = Mid(ListaTmp, 1, Inicio - 2) & Mid(ListaTmp, Inicio + Len(Lista))
                End If
                .MoveNext
            Loop
        End If
    End With
    sBus = "UPDATE JUEGO_REPARACION SET ID_PRODUCTO = '" & Label2.Caption & "' WHERE ID_PRODUCTO IN (" & Lista & ")"
    cnn.Execute (sBus)
    If NoCambiados <> "" Then
        MsgBox "Productos no cambiados" & NoCambiados, , "SACC"
    End If
    txtIDPROD.Text = ""
    txtIDPROD2.Text = ""
    LVSeleccion.ListItems.Clear
    LVProductos.ListItems.Clear
    LVReemplazo.ListItems.Clear
    LVIndividual.ListItems.Clear
    LVJuegos.ListItems.Clear
    Lista = ListaTmp
    SSTab1.Caption = "Juego de Reparacion"
    txtIndex.Text = ""
    Command2.Enabled = False
End Sub
Private Sub Command2_Click()
    LVIndividual.ListItems.Clear
    LVJuegos.ListItems.Remove (Val(txtIndex.Text))
    txtIndex.Text = ""
    Command2.Enabled = False
    SSTab1.Caption = "Juego de Reparacion "
    If LVJuegos.ListItems.Count = 0 Then
        Command1.Enabled = False
    End If
End Sub
Private Sub Command3_Click()
    Dim sBus As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    LVSeleccion.ListItems.Clear
    sBus = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & txtIDPROD.Text & "%'"
    Set tRs = cnn.Execute(sBus)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = LVSeleccion.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End If
    End With
    sBus = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & txtIDPROD.Text & "%'"
    Set tRs = cnn.Execute(sBus)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = LVSeleccion.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Command4_Click()
    Dim sBus As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Label2.Caption = "Ninguno"
    LVReemplazo.ListItems.Clear
    sBus = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & txtIDPROD2.Text & "%'"
    Set tRs = cnn.Execute(sBus)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = LVReemplazo.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End If
    End With
    sBus = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & txtIDPROD2.Text & "%'"
    Set tRs = cnn.Execute(sBus)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = LVReemplazo.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Command5_Click()
    Dim sBus As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim Lista As String
    Dim Cont As Integer
    LVIndividual.ListItems.Clear
    LVJuegos.ListItems.Clear
    Lista = ""
    For Cont = 1 To LVProductos.ListItems.Count
        If Lista <> "" Then
            Lista = Lista & ", "
        End If
        Lista = Lista & "'" & LVProductos.ListItems.Item(Cont) & "'"
    Next Cont
    sBus = "SELECT ID_REPARACION FROM JUEGO_REPARACION WHERE ID_PRODUCTO IN (" & Lista & ")"
    Set tRs = cnn.Execute(sBus)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = LVJuegos.ListItems.Add(, , .Fields("ID_REPARACION"))
                .MoveNext
            Loop
            Command1.Enabled = True
        Else
            Command1.Enabled = False
            MsgBox "No existen juegos de reparacion con esos productos", , "SACC"
        End If
    End With
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Command5.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With LVSeleccion
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_PROD", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
    End With
    With LVProductos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_PROD", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
    End With
    With LVReemplazo
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_PROD", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
    End With
    With LVJuegos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_REPARACION", 3000
    End With
    With LVIndividual
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_PROD", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub LVJuegos_DblClick()
    Dim sBus As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If LVJuegos.ListItems.Count > 0 Then
        LVIndividual.ListItems.Clear
        txtIndex.Text = LVJuegos.SelectedItem.Index
        SSTab1.Caption = "Juego de Reparacion " & LVJuegos.SelectedItem
        sBus = "SELECT * FROM VsJRA1 WHERE ID_REPARACION = '" & LVJuegos.SelectedItem & "'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = LVIndividual.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        tLi.SubItems(1) = .Fields("Descripcion")
                    .MoveNext
                Loop
            End If
        End With
        sBus = "SELECT * FROM VsJRA2 WHERE ID_REPARACION = '" & LVJuegos.SelectedItem & "'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = LVIndividual.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        tLi.SubItems(1) = .Fields("Descripcion")
                    .MoveNext
                Loop
            End If
        End With
        Command2.Enabled = True
    End If
End Sub
Private Sub LVReemplazo_DblClick()
    If LVReemplazo.ListItems.Count > 0 Then
        Label2.Caption = LVReemplazo.SelectedItem
        txtIDPROD2.Text = LVReemplazo.SelectedItem
    End If
End Sub
Private Sub LVSeleccion_DblClick()
    Dim tLi As ListItem
    If LVSeleccion.ListItems.Count > 0 Then
        Set tLi = LVProductos.ListItems.Add(, , LVSeleccion.SelectedItem)
            tLi.SubItems(1) = LVSeleccion.SelectedItem.SubItems(1)
        cmdQuitar.Enabled = True
    End If
End Sub
Private Sub txtIDPROD_Change()
    If (txtIDPROD2.Text <> "") And (Label2.Caption <> "Ninguno") And (LVProductos.ListItems.Count > 0) Then
        Command5.Enabled = True
    Else
        Command5.Enabled = False
    End If
End Sub
Private Sub txtIDPROD2_Change()
    If (txtIDPROD2.Text <> "") And (Label2.Caption <> "Ninguno") And (LVProductos.ListItems.Count > 0) Then
        Command5.Enabled = True
    Else
        Command5.Enabled = False
    End If
End Sub
Private Sub txtIDPROD2_GotFocus()
    txtIDPROD2.BackColor = &HFFE1E1
End Sub
Private Sub txtIDPROD2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command4.Value = True
    End If
End Sub
Private Sub txtIDPROD2_LostFocus()
    txtIDPROD2.BackColor = &H80000005
End Sub
Private Sub txtIDPROD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command3.Value = True
    End If
End Sub
Private Sub txtIDPROD_GotFocus()
    txtIDPROD.BackColor = &HFFE1E1
End Sub
Private Sub txtIDPROD_LostFocus()
    txtIDPROD.BackColor = &H80000005
End Sub
