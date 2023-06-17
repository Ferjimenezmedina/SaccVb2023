VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form BuscaProd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Producto"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   15
      Top             =   5880
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "BuscaProd.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "BuscaProd.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "BuscaProd.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Generales"
      TabPicture(1)   =   "BuscaProd.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
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
         Left            =   -67200
         Picture         =   "BuscaProd.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   " Tipo "
         Height          =   855
         Left            =   -71160
         TabIndex        =   14
         Top             =   480
         Width           =   3735
         Begin VB.OptionButton Option5 
            Caption         =   "Todos"
            Height          =   255
            Left            =   2520
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Compuesto"
            Height          =   255
            Left            =   1320
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Simple"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74160
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   10
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9340
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   720
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   6240
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripcion"
         Height          =   195
         Left            =   6240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
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
         Left            =   7800
         Picture         =   "BuscaProd.frx":4DF6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
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
      Begin VB.Label Label2 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "BuscaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    BuscaGrales
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Buscar
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 5200
        .ColumnHeaders.Add , , "TIPO", 1500
        .ColumnHeaders.Add , , "MARCA", 1500
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
        .ColumnHeaders.Add , , "ALMACEN", 2000
        .ColumnHeaders.Add , , "C. MAXIMA", 2000
        .ColumnHeaders.Add , , "C. MINIMA", 2000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 5200
        .ColumnHeaders.Add , , "TIPO", 1500
        .ColumnHeaders.Add , , "MARCA", 1500
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
        .ColumnHeaders.Add , , "ALMACEN", 2000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Text1.Text
    Me.ListView1.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT * FROM ALMACEN3 WHERE Descripcion LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("TIPO")
                tLi.SubItems(3) = .Fields("MARCA")
                If .Fields("PRECIO_COSTO") <> "" And .Fields("GANANCIA") <> "" Then
                    tLi.SubItems(4) = Format(CDbl(.Fields("PRECIO_COSTO")) * (1 + (CDbl(.Fields("GANANCIA")))), "###,###,##0.00")
                Else
                    tLi.SubItems(4) = "Precio no Disponible"
                End If
                tLi.SubItems(5) = "ALMACEN 3"
                tLi.SubItems(6) = .Fields("C_MAXIMA")
                tLi.SubItems(7) = .Fields("C_MINIMA")
                .MoveNext
            Loop
        End If
    End With
    sBuscar = Text1.Text
    If Option1.Value Then
        sBuscar = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT * FROM ALMACEN2 WHERE Descripcion LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("MARCA") & ""
                tLi.SubItems(4) = "Precio no Disponible"
                tLi.SubItems(5) = "ALMACEN 2"
                .MoveNext
            Loop
        End If
    End With
    sBuscar = Text1.Text
    If Option1.Value Then
        sBuscar = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT * FROM ALMACEN1 WHERE Descripcion LIKE '%" & sBuscar & "%' ORDER BY Descripcion"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("MARCA") & ""
                tLi.SubItems(4) = "Precio no Disponible"
                tLi.SubItems(5) = "ALMACEN 1"
                .MoveNext
            Loop
        End If
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
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    ListView1.SortOrder = 1 Xor ListView1.SortOrder
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    ListView2.SortOrder = 1 Xor ListView2.SortOrder
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then
            Buscar
        Else
            MsgBox "TIENE QUE ESPECIFICAR UNA BUSQUEDA!", vbInformation, "SACC"
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaGrales
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub BuscaGrales()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    If Option3.Value = True Then
        sBuscar = "SIMPLE"
    Else
        If Option4.Value = True Then
            sBuscar = "COMPUESTO"
        Else
            sBuscar = "%"
        End If
    End If
    sBuscar = "SELECT * FROM ALMACEN1 WHERE MARCA LIKE '%" & Text2.Text & "%' AND TIPO LIKE '" & sBuscar & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("MARCA") & ""
                If .Fields("PRECIO_COSTO") <> "" And .Fields("GANANCIA") <> "" Then
                    tLi.SubItems(4) = Format(CDbl(.Fields("PRECIO_COSTO")) * (1 + (CDbl(.Fields("GANANCIA")) / 100)), "###,###,##0.00")
                Else
                    tLi.SubItems(4) = "Precio no Disponible"
                End If
                tLi.SubItems(5) = "ALMACEN 3"
                .MoveNext
            Loop
        End If
    End With
    If Option3.Value = True Then
        sBuscar = "SIMPLE"
    Else
        If Option4.Value = True Then
            sBuscar = "COMPUESTO"
        Else
            sBuscar = "%"
        End If
    End If
    sBuscar = "SELECT * FROM ALMACEN2 WHERE MARCA LIKE '%" & Text2.Text & "%' AND TIPO LIKE '" & sBuscar & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("MARCA") & ""
                If .Fields("PRECIO_COSTO") <> "" And .Fields("GANANCIA") <> "" Then
                    tLi.SubItems(4) = Format(CDbl(.Fields("PRECIO_COSTO")) * (1 + (CDbl(.Fields("GANANCIA")) / 100)), "###,###,##0.00")
                Else
                    tLi.SubItems(4) = "Precio no Disponible"
                End If
                tLi.SubItems(5) = "ALMACEN 3"
                .MoveNext
            Loop
        End If
    End With
    If Option3.Value = True Then
        sBuscar = "SIMPLE"
    Else
        If Option4.Value = True Then
            sBuscar = "COMPUESTO"
        Else
            sBuscar = "%"
        End If
    End If
    sBuscar = "SELECT * FROM ALMACEN3 WHERE MARCA LIKE '%" & Text2.Text & "%' AND TIPO LIKE '" & sBuscar & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("Descripcion") & ""
                tLi.SubItems(2) = .Fields("TIPO") & ""
                tLi.SubItems(3) = .Fields("MARCA") & ""
                If .Fields("PRECIO_COSTO") <> "" And .Fields("GANANCIA") <> "" Then
                    tLi.SubItems(4) = Format(CDbl(.Fields("PRECIO_COSTO")) * (1 + (CDbl(.Fields("GANANCIA")) / 100)), "###,###,##0.00")
                Else
                    tLi.SubItems(4) = "Precio no Disponible"
                End If
                tLi.SubItems(5) = "ALMACEN 3"
                .MoveNext
            Loop
        End If
    End With
End Sub
