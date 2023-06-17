VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepCompras 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Compras"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   15
      Top             =   4800
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepCompras.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCompras.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   13
      Top             =   3600
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
         MouseIcon       =   "FrmRepCompras.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCompras.frx":26F6
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Por Producto"
      TabPicture(0)   =   "FrmRepCompras.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Por Proveedor"
      TabPicture(1)   =   "FrmRepCompras.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(3)=   "ListView3"
      Tab(1).Control(4)=   "ListView4"
      Tab(1).Control(5)=   "Line2"
      Tab(1).Control(6)=   "Label6"
      Tab(1).ControlCount=   7
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
         Left            =   -69240
         Picture         =   "FrmRepCompras.frx":4270
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha de compra"
         Height          =   1815
         Left            =   -70200
         TabIndex        =   21
         Top             =   600
         Width           =   3255
         Begin VB.CheckBox Check1 
            Caption         =   "Por Fecha"
            Height          =   255
            Left            =   960
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1080
            TabIndex        =   9
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39384
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1080
            TabIndex        =   8
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39384
         End
         Begin VB.Label Label5 
            Caption         =   "Del :"
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Al :"
            Height          =   255
            Left            =   720
            TabIndex        =   22
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73920
         TabIndex        =   6
         Top             =   600
         Width           =   3615
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
         Left            =   5760
         Picture         =   "FrmRepCompras.frx":6C42
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   600
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fecha de compra"
         Height          =   1815
         Left            =   4800
         TabIndex        =   17
         Top             =   600
         Width           =   3255
         Begin VB.CheckBox Check2 
            Caption         =   "Por Fecha"
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   1320
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1080
            TabIndex        =   3
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39384
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1080
            TabIndex        =   2
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39384
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   720
            TabIndex        =   20
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Del :"
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   480
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   7
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3625
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   11
         Top             =   3240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
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
      Begin VB.Line Line2 
         X1              =   -74760
         X2              =   -67080
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label6 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7920
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   8520
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "FrmRepCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ItmBUS1 As String
Dim sNoOrden As String
Dim sTipo As String
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Me.DTPicker3.Enabled = True
        Me.DTPicker4.Enabled = True
    Else
        Me.DTPicker3.Enabled = False
        Me.DTPicker4.Enabled = False
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Me.DTPicker1.Enabled = True
        Me.DTPicker2.Enabled = True
    Else
        Me.DTPicker1.Enabled = False
        Me.DTPicker2.Enabled = False
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim Cont As Integer
    sBuscar = ""
    For Cont = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(Cont).Checked Then
            sBuscar = sBuscar & " ID_PRODUCTO = '" & ListView2.ListItems(Cont) & "' OR"
        End If
    Next
    sBuscar = Mid(sBuscar, 1, Len(sBuscar) - 2)
    If Check2.Value = 1 Then
        If sBuscar = "" Then
            sBuscar = " FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
        Else
            sBuscar = sBuscar & " AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' "
        End If
    End If
    If sBuscar <> "" Then
        sBuscar = "SELECT ID_PRODUCTO, NOMBRE, FECHA, PRECIO, SURTIDO, CANTIDAD, NUM_ORDEN, TIPO FROM VsRepCompras WHERE " & sBuscar
    Else
        sBuscar = "SELECT ID_PRODUCTO, NOMBRE, FECHA, PRECIO, SURTIDO, CANTIDAD, NUM_ORDEN, TIPO FROM VsRepCompras " & sBuscar
    End If
    ListView1.ListItems.Clear
    sBuscar = sBuscar & " ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("ID_PRODUCTO")))
            If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(1) = tRs.Fields("PRECIO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(3) = tRs.Fields("SURTIDO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(4) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(6) = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("TIPO")) Then
                If tRs.Fields("TIPO") = "I" Then
                    tLi.SubItems(7) = "Internacional"
                Else
                    If tRs.Fields("TIPO") = "N" Then
                        tLi.SubItems(7) = "Nacional"
                    Else
                        tLi.SubItems(7) = "Indirecta"
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command2_Click()
    If ItmBUS1 <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView4.ListItems.Clear
        sBuscar = "SELECT ID_PRODUCTO, NOMBRE, FECHA, PRECIO, SURTIDO, CANTIDAD, NUM_ORDEN, TIPO  FROM VsRepCompras WHERE ID_PROVEEDOR = " & ItmBUS1 & ""
        If Check1.Value = 1 Then
            sBuscar = sBuscar & " AND FECHA BETWEEN '" & DTPicker4.Value & "' AND '" & DTPicker3.Value & "'"
        End If
        sBuscar = sBuscar & " ORDER BY ID_PROVEEDOR, ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView4.ListItems.Add(, , Trim(tRs.Fields("ID_PRODUCTO")))
                If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(1) = tRs.Fields("PRECIO")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(3) = tRs.Fields("SURTIDO")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(4) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(6) = tRs.Fields("NUM_ORDEN")
                If Not IsNull(tRs.Fields("TIPO")) Then
                    If tRs.Fields("TIPO") = "I" Then
                        tLi.SubItems(7) = "Internacional"
                    Else
                        If tRs.Fields("TIPO") = "N" Then
                            tLi.SubItems(7) = "Nacional"
                        Else
                            tLi.SubItems(7) = "Indirecta"
                        End If
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    Me.DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    Me.DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Me.DTPicker4.Value = Format(Date - 30, "dd/mm/yyyy")
    Me.DTPicker3.Value = Format(Date, "dd/mm/yyyy")
    Me.DTPicker1.Enabled = False
    Me.DTPicker2.Enabled = False
    Me.DTPicker3.Enabled = False
    Me.DTPicker4.Enabled = False
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
        .ColumnHeaders.Add , , "Producto", 2200
        .ColumnHeaders.Add , , "Precio", 1500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Surtido", 1500
        .ColumnHeaders.Add , , "Nombre", 4500
        .ColumnHeaders.Add , , "Fecha de Compra", 1840
        .ColumnHeaders.Add , , "No. Orden", 1840
        .ColumnHeaders.Add , , "Tipo", 1840
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 2440
        .ColumnHeaders.Add , , "Descripcion", 4440
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 0
        .ColumnHeaders.Add , , "Nombre", 4440
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Producto", 2200
        .ColumnHeaders.Add , , "Precio", 1500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Surtido", 1500
        .ColumnHeaders.Add , , "Nombre", 4500
        .ColumnHeaders.Add , , "Fecha de Compra", 1840
        .ColumnHeaders.Add , , "No. Orden", 1840
        .ColumnHeaders.Add , , "Tipo", 1840
    End With
End Sub
Private Sub Image10_Click()
On Error GoTo ManejaError
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
    End If
    If SSTab1.Tab = 1 Then
        If ListView4.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView4.ColumnHeaders.Count
                For Con = 1 To ListView4.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView4.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView4.ListItems.Count
                    StrCopi = StrCopi & ListView4.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_DblClick()
    If sTipo = "Nacional" Or sTipo = "Indirecta" Then
        FunImp
    Else
        FunImpr2
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    sNoOrden = Item.SubItems(6)
    sTipo = Item.SubItems(7)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ItmBUS1 = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView2.ListItems.Clear
        sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM vsALMACENES_123 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView2.ListItems.Add(, , Trim(tRs.Fields("ID_PRODUCTO")))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView3.ListItems.Clear
        sBuscar = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , Trim(tRs.Fields("ID_PROVEEDOR")))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub FunImpr2()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim Moneda As String
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If sTipo = "Nacional" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'N'"
    End If
    If sTipo = "Internacional" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'I'"
    End If
    If sTipo = "Indirecta" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'X'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        If Not IsNull(tRs1.Fields("MONEDA")) Then Moneda = tRs1.Fields("MONEDA")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR  where  STATUS='A' "
        Set tRs5 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Date:" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "PURCHASE ORDER : ", "F3", 8, hCenter
        oDoc.WTextBox 60, 390, 20, 250, sNoOrden, "F2", 11, hCenter
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 100, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 100, 175, "INVOICE TO :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 100, 175, "SHIP TO:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Person in charge :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("TELEFONO3"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 138, 205, 100, 175, "COl." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 175, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        'CAJA3
        oDoc.WTextBox 115, 390, 100, 175, tRs5.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 125, 390, 100, 175, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 145, 390, 100, 175, tRs5.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 155, 390, 100, 175, tRs5.Fields("TEL1"), "F3", 8, hCenter
        oDoc.WTextBox 165, 390, 100, 175, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUN", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, tRs3.Fields("ID_PRODUCTO"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 477, 20, 50, Format(CDbl(tRs3.Fields("PRECIO")) * CDbl(tRs3.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 7, hRight
                Posi = Posi + 12
                tRs3.MoveNext
                If Posi >= 650 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If sTipo = "Nacional" Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'N'"
                    End If
                    If sTipo = "Internacional" Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'I'"
                    End If
                    If sTipo = "Indirecta" Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 30, 340, 20, 250, "PURCHASE ORDER # :", "F3", 9, hCenter, , , 1, vbBlack
                        oDoc.WTextBox 30, 390, 20, 250, sNoOrden, "F3", 11, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "ITEM#", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "QTY", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "AMOUN", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 12
                    End If
                End If
            Loop
        End If
        ' Linea
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 20, 100, 275, "Please include country of origin for all items", "F3", 10, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 100, 275, "Please deliver to our shipping Address. Any questions, contact Purchasing Dept at " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox Posi, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((tRs1.Fields("TOTAL")), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Less Discount:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((tRs1.Fields("DISCOUNT")), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Other Charges:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((tRs1.Fields("OTROS_CARGOS")), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((tRs1.Fields("FREIGHT")), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Sales Tax:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((tRs1.Fields("TAX")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 20, 100, 275, "COMENTARIOS:", "F2", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox 690, 20, 100, 300, Format((tRs1.Fields("COMENTARIO")), "###,###,##0.00"), "F3", 11, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox Posi, 400, 20, 70, "TOTAL AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(tRs1.Fields("TOTAL")) + CDbl(tRs1.Fields("DISCOUNT") + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")) + CDbl(tRs1.Fields("TAX"))), "###,###,##0.00"), "F3", 8, hRight
        'totales
        Posi = Posi + 6
        'oDoc.WTextBox Posi, 200, 20, 250, "Mr Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox Posi, 15, 20, 250, "Prices expressed in " & Moneda, "F3", 10, hCenter
        Posi = Posi + 10
        oDoc.WTextBox Posi, 200, 20, 250, "Autorized Signature", "F3", 8, hCenter
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImp()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If sTipo = "Nacional" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'N'"
    End If
    If sTipo = "Internacional" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'I'"
    End If
    If sTipo = "Indirecta" Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'X'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompranacioa.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        ' Encabezado del reporte
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 38, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : ", "F3", 10, hCenter
        oDoc.WTextBox 60, 395, 20, 250, sNoOrden, "F2", 10, hCenter
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 100, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 100, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 100, 175, "ENVIAR A:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Responsable :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("TELEFONO3"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 138, 205, 100, 175, "COl." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 175, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        'CAJA3
        oDoc.WTextBox 125, 390, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 135, 390, 100, 175, "COl." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 145, 390, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 155, 390, 100, 175, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 165, 390, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "CLAVE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "CANTIDAD", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "PRECIO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "TOTAL", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, tRs3.Fields("ID_PRODUCTO"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 160, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 477, 20, 50, Format(CDbl(tRs3.Fields("PRECIO")) * CDbl(tRs3.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 8, hRight
                Posi = Posi + 12
                tRs3.MoveNext
                If Posi >= 650 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'N'"
                    If sTipo = "Indirecta" Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & sNoOrden & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : " & sNoOrden, "F3", 10, hCenter
                        If tRs1.Fields("REVISION") <> 0 Then
                            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
                        End If
                        oDoc.WTextBox 30, 380, 20, 250, Text2.Text, "F3", 8, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "CLAVE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "CANTIDAD", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPCION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "PRECIO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "TOTAL", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 12
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format((tRs1.Fields("TOTAL")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 640, 488, 20, 50, Format((tRs1.Fields("DISCOUNT")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 660, 488, 20, 50, Format((tRs1.Fields("OTROS_CARGOS")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 680, 488, 20, 50, Format((tRs1.Fields("FREIGHT")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 700, 488, 20, 50, Format((tRs1.Fields("TAX")), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 720, 488, 20, 50, Format(CDbl(tRs1.Fields("TOTAL")) + CDbl(tRs1.Fields("DISCOUNT") + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")) + CDbl(tRs1.Fields("TAX"))), "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 730, 200, 20, 250, "Lic. Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox 730, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 740, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 10, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
