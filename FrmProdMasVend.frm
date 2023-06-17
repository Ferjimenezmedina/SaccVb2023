VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmProdMasVend 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos mas vendidos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   17
      Top             =   3240
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmProdMasVend.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdMasVend.frx":030A
         Top             =   240
         Width           =   720
      End
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
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   1
      Top             =   4440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProdMasVend.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdMasVend.frx":2156
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Almacen3"
      TabPicture(0)   =   "FrmProdMasVend.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Combo3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBuscar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ListView3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Almacen 1 y 2"
      TabPicture(1)   =   "FrmProdMasVend.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "Combo4"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "Combo6"
      Tab(1).ControlCount=   7
      Begin MSComctlLib.ListView ListView3 
         Height          =   1215
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
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
         Left            =   960
         TabIndex        =   32
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   -73440
         TabIndex        =   27
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rango de fechas"
         Height          =   1695
         Left            =   -69480
         TabIndex        =   21
         Top             =   480
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "Seleccionar Rango"
            Height          =   195
            Left            =   600
            TabIndex        =   22
            Top             =   1320
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1080
            TabIndex        =   23
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51576833
            CurrentDate     =   39296
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1080
            TabIndex        =   24
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51576833
            CurrentDate     =   39296
         End
         Begin VB.Label Label7 
            Caption         =   "Del :"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Al :"
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   -73440
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
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
         Left            =   -71040
         Picture         =   "FrmProdMasVend.frx":4270
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
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
         Left            =   4320
         Picture         =   "FrmProdMasVend.frx":6C42
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de fechas"
         Height          =   1695
         Left            =   6240
         TabIndex        =   6
         Top             =   480
         Width           =   2295
         Begin VB.CheckBox Check1 
            Caption         =   "Seleccionar Rango"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   600
            TabIndex        =   10
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51576833
            CurrentDate     =   39296
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   51576833
            CurrentDate     =   39296
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   28
         Top             =   2280
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
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
      Begin VB.Label Label9 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   -74280
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   -74280
         TabIndex        =   29
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmProdMasVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim IdClien As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
    If Check1.value = 1 Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
End Sub
Private Sub Check2_Click()
    If Check2.value = 1 Then
        DTPicker3.Enabled = True
        DTPicker4.Enabled = True
    Else
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
    End If
End Sub
Private Sub cmdBuscar_Click()
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD FROM VsMasVendidos "
    If Combo1.Text <> "" Or Combo2.Text <> "" Or Combo3.Text <> "" Or Check1.value = 1 Then
        sBuscar = sBuscar & "WHERE "
    End If
    If IdClien <> "" Then
        sBuscar = sBuscar & " ID_CLIENTE = " & IdClien
        If Combo1.Text <> "" Or Combo2.Text <> "" Or Combo3.Text <> "" Or Check1.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Combo1.Text <> "" Then
        sBuscar = sBuscar & " SUCURSAL = '" & Combo1.Text & "'"
        If Combo2.Text <> "" Or Combo3.Text <> "" Or Check1.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Combo2.Text <> "" Then
        sBuscar = sBuscar & " CLASIFICACION = '" & Combo2.Text & "'"
        If Combo3.Text <> "" Or Check1.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Combo3.Text <> "" Then
        sBuscar = sBuscar & " MARCA = '" & Combo3.Text & "'"
        If Check1.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Check1.value = 1 Then
        If DTPicker1.value > DTPicker2.value Then
            Dim AcomFecha As String
            AcomFecha = DTPicker1.value
            DTPicker1.value = DTPicker2.value
            DTPicker2.value = AcomFecha
        End If
        sBuscar = sBuscar & " FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & "'"
    End If
    sBuscar = sBuscar & " GROUP BY ID_PRODUCTO, Descripcion"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    Else
        MsgBox "LA BUSQUEDA NO GENERO RESULTADOS!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD FROM VsMasVendidosInsumos "
    If Combo4.Text <> "" Or Combo6.Text <> "" Or Check2.value = 1 Then
        sBuscar = sBuscar & "WHERE "
    End If
    If Combo6.Text <> "" Then
        sBuscar = sBuscar & " SUCURSAL = '" & Combo6.Text & "'"
        If Combo4.Text <> "" Or Check2.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Combo4.Text <> "" Then
        sBuscar = sBuscar & " MARCA = '" & Combo4.Text & "'"
        If Check1.value = 1 Then
            sBuscar = sBuscar & " AND "
        End If
    End If
    If Check2.value = 1 Then
        If DTPicker4.value > DTPicker3.value Then
            Dim AcomFecha As String
            AcomFecha = DTPicker3.value
            DTPicker3.value = DTPicker4.value
            DTPicker4.value = AcomFecha
        End If
        sBuscar = sBuscar & " FECHA BETWEEN '" & DTPicker4.value & "' AND '" & DTPicker3.value & "'"
    End If
    sBuscar = sBuscar & " GROUP BY ID_PRODUCTO, Descripcion"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    Else
        MsgBox "LA BUSQUEDA NO GENERO RESULTADOS!", vbInformation, "SACC"
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.value = Format(Date, "dd/mm/yyyy")
    DTPicker2.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker4.value = Format(Date, "dd/mm/yyyy")
    DTPicker3.value = Format(Date - 30, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
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
        .ColumnHeaders.Add , , "Clave", 1800
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Total", 1000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 1800
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Total", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 0
        .ColumnHeaders.Add , , "Nombre", 5500
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            Combo6.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT CLASIFICACION FROM CLASIFICACIONES ORDER BY CLASIFICACION"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("CLASIFICACION")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT MARCA FROM ALMACEN3 GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo3.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT MARCA FROM ALMACEN2 GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo4.AddItem tRs.Fields("MARCA")
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
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If Ruta <> "" Then
        NumColum = ListView1.ColumnHeaders.COUNT
        For Con = 1 To ListView1.ColumnHeaders.COUNT
            If SSTab1.Tab = 0 Then
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Else
                StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
            End If
        Next
        StrCopi = StrCopi & Chr(13)
        For Con = 1 To ListView1.ListItems.COUNT
            StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
            For Con2 = 1 To NumColum - 1
                If SSTab1.Tab = 0 Then
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Else
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                End If
            Next
            StrCopi = StrCopi & Chr(13)
        Next
        'archivo TXT
        Dim foo As Integer
        foo = FreeFile
        Open Ruta For Output As #foo
            Print #foo, StrCopi
        Close #foo
    End If
    ShellExecute Me.hWnd, "open", Ruta, "", "", 4
End Sub
Private Sub Image9_Click()
    Unload Me
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
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdClien = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView3.ListItems.Clear
        sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
