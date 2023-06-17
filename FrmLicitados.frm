VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmLicitados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos Licitados"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   4
      Top             =   3240
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmLicitados.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmLicitados.frx":030A
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8070
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   2
      Top             =   4440
      Width           =   975
      Begin VB.Label Label9 
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmLicitados.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmLicitados.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmLicitados.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5400
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
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
         Left            =   7080
         Picture         =   "FrmLicitados.frx":40D4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmLicitados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Public IdCliente As String
Private Sub cmdOk_Click()
    Actualiza
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwAutomatic
        .HideSelection = False
        .HotTracking = False
        .FullRowSelect = True
        .HoverSelection = False
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Clave Producto", 1500
        .ColumnHeaders.Add , , "Descripción", 5100
    End With
    Actualiza
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualiza()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.Value = True Then
        sBuscar = "SELECT LICITACIONES.ID_PRODUCTO, ALMACEN3.DESCRIPCION FROM LICITACIONES INNER JOIN ALMACEN3 ON LICITACIONES.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (LICITACIONES.ID_PRODUCTO LIKE '%" & Text1.Text & "%') AND (LICITACIONES.FECHA_FIN >= GETDATE()) AND (LICITACIONES.FECHA_INICIO <= GETDATE()) AND LICITACIONES.ID_CLIENTE = " & IdCliente & " ORDER BY ALMACEN3.DESCRIPCION"
    Else
        sBuscar = "SELECT LICITACIONES.ID_PRODUCTO, ALMACEN3.DESCRIPCION FROM LICITACIONES INNER JOIN ALMACEN3 ON LICITACIONES.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (ALMACEN3.DESCRIPCION LIKE '%" & Text1.Text & "%') AND (LICITACIONES.FECHA_FIN >= GETDATE()) AND (LICITACIONES.FECHA_INICIO <= GETDATE()) AND LICITACIONES.ID_CLIENTE = " & IdCliente & " ORDER BY ALMACEN3.DESCRIPCION"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , "")
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("DESCRIPCION")
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
    Dim NumReg As Integer
    Dim Can As Double
    Dim Desc As String
    Dim Tipo As String
    Dim Minima As String
    Dim Marca As String
    Dim PVenta As Double
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    Text1.SetFocus
    'ListView1.HoverSelection = True
    NumReg = ListView1.ListItems.Count
    For Con = 1 To NumReg
        If ListView1.ListItems(Con) <> "" And ListView1.ListItems(Con) <> "0" Then
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView1.ListItems(Con).SubItems(1) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Can = CDbl(tRs.Fields("CANTIDAD"))
            Else
                Can = 0
            End If
            sBuscar = "SELECT DESCRIPCION, TIPO, C_MINIMA, MARCA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Con).SubItems(1) & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Desc = tRs.Fields("DESCRIPCION")
                Tipo = tRs.Fields("TIPO")
                Minima = tRs.Fields("C_MINIMA")
                Marca = tRs.Fields("MARCA")
            End If
            sBuscar = "SELECT (PRECIO_COSTO * (1 + GANANCIA)) AS P_VENTA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Con).SubItems(1) & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                PVenta = CDbl(tRs.Fields("P_VENTA"))
            Else
                PVenta = 0
            End If
            Set tLi = Programadas.ListView2.ListItems.Add(, , ListView1.ListItems(Con).SubItems(1))
            tLi.SubItems(1) = ListView1.ListItems(Con)
            tLi.SubItems(2) = Can
            If CDbl(ListView1.ListItems(Con)) - Can >= 0 Then
                tLi.SubItems(3) = CDbl(ListView1.ListItems(Con)) - Can
            Else
                tLi.SubItems(3) = ListView1.ListItems(Con)
            End If
            tLi.SubItems(4) = Desc
            tLi.SubItems(5) = Tipo
            tLi.SubItems(6) = Minima
            tLi.SubItems(7) = Marca
            tLi.SubItems(8) = PVenta
        End If
    Next Con
    'ListView1.HoverSelection = False
    Actualiza
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    NewString = Replace(NewString, "A", "")
    NewString = Replace(NewString, "B", "")
    NewString = Replace(NewString, "C", "")
    NewString = Replace(NewString, "D", "")
    NewString = Replace(NewString, "E", "")
    NewString = Replace(NewString, "F", "")
    NewString = Replace(NewString, "G", "")
    NewString = Replace(NewString, "H", "")
    NewString = Replace(NewString, "I", "")
    NewString = Replace(NewString, "J", "")
    NewString = Replace(NewString, "K", "")
    NewString = Replace(NewString, "L", "")
    NewString = Replace(NewString, "M", "")
    NewString = Replace(NewString, "N", "")
    NewString = Replace(NewString, "Ñ", "")
    NewString = Replace(NewString, "O", "")
    NewString = Replace(NewString, "P", "")
    NewString = Replace(NewString, "Q", "")
    NewString = Replace(NewString, "R", "")
    NewString = Replace(NewString, "S", "")
    NewString = Replace(NewString, "T", "")
    NewString = Replace(NewString, "U", "")
    NewString = Replace(NewString, "V", "")
    NewString = Replace(NewString, "W", "")
    NewString = Replace(NewString, "X", "")
    NewString = Replace(NewString, "Y", "")
    NewString = Replace(NewString, "Z", "")
    NewString = Replace(NewString, ",", "")
    NewString = Replace(NewString, ";", "")
    NewString = Replace(NewString, ":", "")
    NewString = Replace(NewString, "-", "")
    NewString = Replace(NewString, "_", "")
    NewString = Replace(NewString, "(", "")
    NewString = Replace(NewString, ")", "")
    NewString = Replace(NewString, "&", "")
    NewString = Replace(NewString, "/", "")
    NewString = Replace(NewString, "%", "")
    NewString = Replace(NewString, "$", "")
    NewString = Replace(NewString, "#", "")
    NewString = Replace(NewString, "!", "")
    NewString = Replace(NewString, "¨", "")
    NewString = Replace(NewString, "+", "")
    NewString = Replace(NewString, "*", "")
    NewString = Replace(NewString, "[", "")
    NewString = Replace(NewString, "]", "")
    NewString = Replace(NewString, "^", "")
    NewString = Replace(NewString, "{", "")
    NewString = Replace(NewString, "}", "")
    NewString = Replace(NewString, "a", "")
    NewString = Replace(NewString, "b", "")
    NewString = Replace(NewString, "c", "")
    NewString = Replace(NewString, "d", "")
    NewString = Replace(NewString, "e", "")
    NewString = Replace(NewString, "f", "")
    NewString = Replace(NewString, "g", "")
    NewString = Replace(NewString, "h", "")
    NewString = Replace(NewString, "i", "")
    NewString = Replace(NewString, "j", "")
    NewString = Replace(NewString, "k", "")
    NewString = Replace(NewString, "l", "")
    NewString = Replace(NewString, "m", "")
    NewString = Replace(NewString, "n", "")
    NewString = Replace(NewString, "ñ", "")
    NewString = Replace(NewString, "o", "")
    NewString = Replace(NewString, "p", "")
    NewString = Replace(NewString, "q", "")
    NewString = Replace(NewString, "r", "")
    NewString = Replace(NewString, "s", "")
    NewString = Replace(NewString, "t", "")
    NewString = Replace(NewString, "u", "")
    NewString = Replace(NewString, "v", "")
    NewString = Replace(NewString, "w", "")
    NewString = Replace(NewString, "x", "")
    NewString = Replace(NewString, "y", "")
    NewString = Replace(NewString, "z", "")
End Sub
Private Sub ListView1_Click()
    Me.ListView1.StartLabelEdit
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
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
