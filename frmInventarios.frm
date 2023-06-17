VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmInventarios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventarios"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmInventarios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListView1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CommonDialog1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   840
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1560
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Almacen 3"
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Almacen 2"
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Almacen 1"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   1560
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Clasificación"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   9
      Top             =   2640
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmInventarios.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmInventarios.frx":0326
         Top             =   120
         Width           =   720
      End
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmInventarios.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "frmInventarios.frx":2712
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "frmInventarios.frx":4254
         MousePointer    =   99  'Custom
         Picture         =   "frmInventarios.frx":455E
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmInventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "Id Producto", 0
        .ColumnHeaders.Add , , "Descripcion", 1440
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Sucursal", 1000
        .ColumnHeaders.Add , , "Clasificacion", 0
        .ColumnHeaders.Add , , "Almacen", 0
    End With
    sBuscar = "SELECT CLASIFICACION FROM ALMACEN3 GROUP BY CLASIFICACION ORDER BY CLASIFICACION"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("CLASIFICACION")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Combo1.AddItem "<TODAS>"
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Consulta()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Check1.Value = 1 Then
        sBuscar = "SELECT ALMACEN1.ID_PRODUCTO, ALMACEN1.DESCRIPCION, EXISTENCIAS.CANTIDAD, EXISTENCIAS.SUCURSAL FROM ALMACEN1 INNER JOIN EXISTENCIAS ON ALMACEN1.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (ALMACEN1.CLASIFICACION <> 'SERVICIOS') AND (ALMACEN1.CLASIFICACION <> 'SERVICIO') "
        If Combo1.Text <> "" And Combo1.Text <> "<TODAS>" Then
            sBuscar = sBuscar & "AND EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' "
        End If
        sBuscar = sBuscar & "ORDER BY EXISTENCIAS.SUCURSAL, ALMACEN1.ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
                tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                tLi.SubItems(4) = "Materia Prima"
                tLi.SubItems(5) = "1"
                tRs.MoveNext
            Loop
        End If
    End If
    If Check2.Value = 1 Then
        sBuscar = "SELECT ALMACEN2.ID_PRODUCTO, ALMACEN2.DESCRIPCION, EXISTENCIAS.CANTIDAD, EXISTENCIAS.SUCURSAL FROM ALMACEN2 INNER JOIN EXISTENCIAS ON ALMACEN2.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (ALMACEN2.CLASIFICACION <> 'SERVICIOS') AND (ALMACEN2.CLASIFICACION <> 'SERVICIO') "
        If Combo1.Text <> "" And Combo1.Text <> "<TODAS>" Then
            sBuscar = sBuscar & "AND EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' "
        End If
        sBuscar = sBuscar & "ORDER BY EXISTENCIAS.SUCURSAL, ALMACEN2.ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
                tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                tLi.SubItems(4) = "Materia Prima"
                tLi.SubItems(5) = "2"
                tRs.MoveNext
            Loop
        End If
    End If
    If Check3.Value = 1 Then
        sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.DESCRIPCION, EXISTENCIAS.CANTIDAD, EXISTENCIAS.SUCURSAL, ALMACEN3.CLASIFICACION FROM ALMACEN3 INNER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (ALMACEN3.CLASIFICACION <> 'SERVICIOS') AND (ALMACEN3.CLASIFICACION <> 'SERVICIO') "
        If Combo1.Text <> "" And Combo1.Text <> "<TODAS>" Then
            sBuscar = sBuscar & "AND EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' "
        End If
        If Combo2.Text <> "" Then
            sBuscar = sBuscar & "AND ALMACEN3.CLASIFICACION  = '" & Combo2.Text & "' "
        End If
        sBuscar = sBuscar & "ORDER BY EXISTENCIAS.SUCURSAL, ALMACEN3.ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
                tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                tLi.SubItems(4) = tRs.Fields("CLASIFICACION")
                tLi.SubItems(5) = "3"
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Consulta
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
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Posi As Integer
    Dim Con As Integer
    Set oDoc = New cPDF
    Consulta
    Posi = 120
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\Inventarios.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Courier_Bold, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F1", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F1", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F1", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "REPORTE DE INVENTARIOS", "F2", 10, hCenter
    oDoc.WTextBox 60, 400, 20, 250, "FECHA DE IMPRESION", "F2", 8, hCenter
    oDoc.WTextBox 70, 510, 20, 250, Format(Date, "dd/mm/yyyy"), "F2", 8, hLeft
' Encabezado de pagina
' Cuerpo del reporte
    oDoc.WTextBox 100, 5, 40, 145, "CLAVE", "F2", 10, hLeft
    oDoc.WTextBox 100, 100, 40, 300, "DESCRIPCION", "F2", 10, hLeft
    oDoc.WTextBox 100, 350, 40, 70, "CANTIDAD", "F2", 10, hRight
    oDoc.WTextBox 100, 370, 40, 100, "ALMACEN.", "F2", 10, hRight
    oDoc.WTextBox 100, 470, 40, 80, "SUCURSAL", "F2", 10, hRight
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 185
    oDoc.WLineTo 580, 185
    For Con = 1 To ListView1.ListItems.Count
        oDoc.WTextBox Posi, 5, 40, 145, ListView1.ListItems(Con), "F1", 9, hLeft
        oDoc.WTextBox Posi, 100, 9, 300, Mid(ListView1.ListItems(Con).SubItems(1), 1, 58), "F1", 9, hLeft
        oDoc.WTextBox Posi, 350, 40, 70, ListView1.ListItems(Con).SubItems(2), "F1", 9, hRight
        oDoc.WTextBox Posi, 370, 40, 100, ListView1.ListItems(Con).SubItems(3), "F1", 9, hRight
        oDoc.WTextBox Posi, 470, 40, 80, ListView1.ListItems(Con).SubItems(4), "F1", 9, hRight
        If Posi > 750 Then
            Posi = 120
            oDoc.NewPage A4_Vertical
            oDoc.WImage 80, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
            oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F1", 7, hCenter
            oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F1", 7, hCenter
            oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F1", 7, hCenter
            oDoc.WTextBox 90, 200, 20, 250, "REPORTE DE INVENTARIOS", "F2", 10, hCenter
            oDoc.WTextBox 60, 400, 20, 250, "FECHA DE IMPRESION", "F2", 8, hCenter
            oDoc.WTextBox 70, 510, 20, 250, Format(Date, "dd/mm/yyyy"), "F2", 8, hLeft
        ' Encabezado de pagina
        ' Cuerpo del reporte
            oDoc.WTextBox 100, 5, 40, 145, "CLAVE", "F2", 10, hLeft
            oDoc.WTextBox 100, 100, 40, 300, "DESCRIPCION", "F2", 10, hLeft
            oDoc.WTextBox 100, 350, 40, 70, "CANTIDAD", "F2", 10, hRight
            oDoc.WTextBox 100, 370, 40, 100, "ALMACEN.", "F2", 10, hRight
            oDoc.WTextBox 100, 470, 40, 80, "SUCURSAL", "F2", 10, hRight
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 185
            oDoc.WLineTo 580, 185
        End If
        Posi = Posi + 10
    Next Con
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi + 20
    oDoc.WLineTo 580, Posi + 20
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
