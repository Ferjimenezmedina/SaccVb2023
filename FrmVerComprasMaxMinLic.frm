VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVerComprasMaxMinLic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver pendientes de entrega en licitación"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8280
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmVerComprasMaxMinLic.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerComprasMaxMinLic.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8280
      TabIndex        =   9
      Top             =   4920
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmVerComprasMaxMinLic.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerComprasMaxMinLic.frx":2156
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "FrmVerComprasMaxMinLic.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Concentrado"
      TabPicture(1)   =   "FrmVerComprasMaxMinLic.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "Option2"
      Tab(1).Control(2)=   "Option1"
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(4)=   "ListView1"
      Tab(1).Control(5)=   "Label1"
      Tab(1).ControlCount=   6
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   4455
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
         Left            =   6240
         Picture         =   "FrmVerComprasMaxMinLic.frx":4270
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
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
         Left            =   -68760
         Picture         =   "FrmVerComprasMaxMinLic.frx":6C42
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   -70320
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   -70320
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -73920
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4815
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8493
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   1
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8493
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
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmVerComprasMaxMinLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS TOTAL, CANT_MIN, CANT_MAX From VsVentasLicitacionMaxMin WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'GROUP BY ID_PRODUCTO, DESCRIPCION, CANT_MIN, CANT_MAX ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS TOTAL, CANT_MIN, CANT_MAX From VsVentasLicitacionMaxMin WHERE Descripcion LIKE '%" & Text1.Text & "%'GROUP BY ID_PRODUCTO, DESCRIPCION, CANT_MIN, CANT_MAX ORDER BY ID_PRODUCTO"
    End If
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(3) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(4) = tRs.Fields("CANT_MAX")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT NOMBRE, ID_PRODUCTO, PRECIO_VENTA, SUM(CANTIDAD) AS TOTAL, CANT_MIN, CANT_MAX  From VsVentasLicitacionMaxMin WHERE NOMBRE LIKE '%" & Text2.Text & "%' GROUP BY NOMBRE, ID_PRODUCTO, PRECIO_VENTA, CANT_MIN, CANT_MAX ORDER BY NOMBRE, ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("NOMBRE"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(2) = tRs.Fields("PRECIO_VENTA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(4) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(5) = tRs.Fields("CANT_MAX")
            tRs.MoveNext
        Loop
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 6100
        .ColumnHeaders.Add , , "CANTIDAD ENTREGADA", 1200
        .ColumnHeaders.Add , , "MINIMO LICITACIÓN", 1200
        .ColumnHeaders.Add , , "MAXIMO LICITACIÓN", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 1200
        .ColumnHeaders.Add , , "CANTIDAD ENTREGADA", 1200
        .ColumnHeaders.Add , , "MINIMO LICITACIÓN", 1200
        .ColumnHeaders.Add , , "MAXIMO LICITACIÓN", 1200
    End With
End Sub
Private Sub Image10_Click()
    If SSTab1.Tab = 0 And ListView2.ListItems.Count <> 0 Then
        LlevaExcel1
    Else
        If ListView1.ListItems.Count <> 0 Then
            LlevaExcel2
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub LlevaExcel1()
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
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub LlevaExcel2()
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
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
