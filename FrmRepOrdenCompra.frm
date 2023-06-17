VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepOrdenCompra 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de ordenes de compra"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   10320
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   17
      Top             =   3720
      Width           =   975
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepOrdenCompra.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenCompra.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   8
      Top             =   4920
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepOrdenCompra.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenCompra.frx":2156
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ordenes de compra"
      TabPicture(0)   =   "FrmRepOrdenCompra.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmRepOrdenCompra.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   9340
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         Height          =   975
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   3615
         Begin VB.CheckBox Check1 
            Caption         =   "Filtrar por fecha"
            Height          =   195
            Left            =   1200
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60030977
            CurrentDate     =   39993
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60030977
            CurrentDate     =   39993
         End
         Begin VB.Label Label2 
            Caption         =   "Al"
            Height          =   255
            Left            =   1920
            TabIndex        =   13
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Del"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   6840
         Picture         =   "FrmRepOrdenCompra.frx":4270
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
         Begin VB.OptionButton Option4 
            Caption         =   "Todas"
            Height          =   195
            Left            =   1560
            TabIndex        =   6
            Top             =   600
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Indirectas"
            Height          =   195
            Left            =   1560
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Internacionales"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nacionales"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7646
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
   End
End
Attribute VB_Name = "FrmRepOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub cmdBuscar_Click()
    BuscaOrden
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Num. Orden", 1500
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Fecha", 1440
        .ColumnHeaders.Add , , "Total", 1440
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 1500
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio", 1440
        .ColumnHeaders.Add , , "Surtido", 1200
    End With
    BuscaOrden
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub BuscaOrden()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA FROM VsRepOrdenCompra WHERE TIPO = 'N'"
    End If
    If Option2.Value Then
        sBuscar = "SELECT NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA FROM VsRepOrdenCompra WHERE TIPO = 'I'"
    End If
    If Option3.Value Then
        sBuscar = "SELECT NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA FROM VsRepOrdenCompra WHERE TIPO = 'X'"
    End If
    If Option4.Value Then
        sBuscar = "SELECT NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA FROM VsRepOrdenCompra WHERE TIPO IN ('N', 'I', 'X')"
    End If
    If Check1.Value = 1 Then
        sBuscar = sBuscar & " AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'  GROUP BY NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA"
    Else
        sBuscar = sBuscar & " GROUP BY NUM_ORDEN, TIPO, NOMBRE, FECHA, TOTAL, CONFIRMADA"
    End If
    Set tRs = cnn.Execute(sBuscar)
    sBuscar = "SELECT TOP 1 VENTA FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(1) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            If tRs.Fields("TIPO") <> "I" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = Format(tRs.Fields("TOTAL"), "###,###,###,##0.00")
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = Format(tRs.Fields("TOTAL") * tRs1.Fields("VENTA"), "###,###,###,##0.00")
            End If
            tLi.Bold = True
            If tRs.Fields("CONFIRMADA") = "Y" Then
                tLi.ForeColor = &HC00000
                tLi.ListSubItems(1).ForeColor = &HC00000
                tLi.ListSubItems(2).ForeColor = &HC00000
                tLi.ListSubItems(3).ForeColor = &HC00000
                tLi.ListSubItems(4).ForeColor = &HC00000
            Else
                tLi.ForeColor = &HFF&
                tLi.ListSubItems(1).ForeColor = &HFF&
                tLi.ListSubItems(2).ForeColor = &HFF&
                tLi.ListSubItems(3).ForeColor = &HFF&
                tLi.ListSubItems(4).ForeColor = &HFF&
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If SSTab1.Tab = 0 Then
        If ListView1.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView1.ColumnHeaders.Count
                For Con = 1 To ListView1.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView1.ListItems.Count
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
                'archivo TXT
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
        End If
    Else
        If ListView2.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView2.ColumnHeaders.Count
                For Con = 1 To ListView2.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView2.ListItems.Count
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
                'archivo TXT
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT TOP 1 VENTA FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs1 = cnn.Execute(sBuscar)
    sBuscar = "SELECT * FROM VsRepOrdenCompra WHERE NUM_ORDEN = " & Item & " AND TIPO = '" & ListView1.SelectedItem.SubItems(1) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If tRs.Fields("TIPO") <> "I" Then
                If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(3) = Format(tRs.Fields("PRECIO"), "###,###,###,##0.00")
            Else
                If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(3) = Format(tRs.Fields("PRECIO") * tRs1.Fields("VENTA"), "###,###,###,##0.00")
            End If
            If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(4) = tRs.Fields("SURTIDO")
            If CDbl(tRs.Fields("SURTIDO")) = CDbl(tRs.Fields("CANTIDAD")) Then
                tLi.ForeColor = &HC00000
                tLi.ListSubItems(1).ForeColor = &HC00000
                tLi.ListSubItems(2).ForeColor = &HC00000
                tLi.ListSubItems(3).ForeColor = &HC00000
                tLi.ListSubItems(4).ForeColor = &HC00000
            Else
                tLi.ForeColor = &HFF&
                tLi.ListSubItems(1).ForeColor = &HFF&
                tLi.ListSubItems(2).ForeColor = &HFF&
                tLi.ListSubItems(3).ForeColor = &HFF&
                tLi.ListSubItems(4).ForeColor = &HFF&
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
