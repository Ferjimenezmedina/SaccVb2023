VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepOrdenCancel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ordenes Canceladas"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9855
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   5
      Top             =   5040
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepOrdenCancel.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenCancel.frx":030A
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   3
      Top             =   2640
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepOrdenCancel.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenCancel.frx":26F6
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmRepOrdenCancel.frx":4238
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenCancel.frx":4542
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepOrdenCancel.frx":6114
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame3 
         Height          =   2775
         Left            =   2640
         TabIndex        =   19
         Top             =   120
         Width           =   5775
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
            Left            =   4440
            Picture         =   "FrmRepOrdenCancel.frx":6130
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   20
            Top             =   240
            Width           =   4695
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2778
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
         Begin VB.Label Label3 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clasificacion"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "Orden Rapida"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nacional"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Internacional"
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Indirecta"
            Height          =   195
            Left            =   2880
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Inicio"
         Height          =   195
         Left            =   7920
         TabIndex        =   13
         Top             =   5880
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango del Reporte"
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   720
            TabIndex        =   9
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50659329
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50659329
            CurrentDate     =   39576
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5106
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   8760
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "FrmRepOrdenCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim sIdProv As String
Private Sub Option1_Click()
    cmdBuscar.Enabled = True
    Option2.Value = False
    Check3.Value = 0
    Option3.Value = False
End Sub
Private Sub Option2_Click()
    cmdBuscar.Enabled = True
    Option1.Value = False
    Check3.Value = 0
    Option3.Value = False
End Sub
Private Sub Option3_Click()
    cmdBuscar.Enabled = True
    Option1.Value = False
    Check3.Value = 0
    Option2.Value = False
End Sub
Private Sub Check3_Click()
    cmdBuscar.Enabled = True
    Option2.Value = False
    Option1.Value = False
    Option3.Value = False
End Sub
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    ListView1.ListItems.Clear
    If Option1.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_ORDEN_COMPRA = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND (ORDEN_COMPRA.CONFIRMADA = 'E') ORDER BY ORDEN_COMPRA.NUM_ORDEN"
        StrRep = sBuscar
    End If
    StrRep = sBuscar
    If Option2.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_ORDEN_COMPRA = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND (ORDEN_COMPRA.CONFIRMADA = 'E') ORDER BY ORDEN_COMPRA.NUM_ORDEN"
        StrRep2 = sBuscar
    End If
    If Option3.Value = True Then
        sBuscar = "SELECT  ORDEN_RAPIDA.ID_ORDEN_RAPIDA AS NUM_ORDEN, 'R' AS TIPO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, (SELECT SUM(TOTAL) AS Expr1 From ORDEN_RAPIDA_DETALLE WHERE (ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA)) AS TOTAL FROM ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY ORDEN_RAPIDA.ID_ORDEN_RAPIDA"
        StrRep4 = sBuscar
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
            If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(1) = tRs.Fields("TIPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
    Dim sBuscar As String
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Orden", 1000
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Proveedor", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "Nombre", 5500
    End With
End Sub
Private Sub Image1_Click()
    Dim Path As String
    Dim SelectionFormula As Date
    Path = App.Path
    Imprimir
End Sub
Private Sub Image10_Click()
    If ListView1.ListItems.Count > 0 Then
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
        StrCopi = "Id_Proveedor" & Chr(9) & "Nombre" & Chr(9) & "Fecha" & Chr(9) & "Total" & Chr(9) & " Num_Orden" & Chr(9) & "Confirmada" & Chr(13)
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
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
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = ListView2.SelectedItem.SubItems(1)
    Text1.SetFocus
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    ListView2.ListItems.Clear
    If KeyAscii = 13 Then
        If Option3.Value Then
            sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        Else
            sqlQuery = "SELECT * FROM PROVEEDOR_CONSUMO WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        End If
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.BOF And .EOF) Then
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                    .MoveNext
                Loop
            End If
        End With
   End If
End Sub
Private Sub Imprimir()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumor As Double
    Dim sumpr As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim Total As Double
    Dim Total1 As Double
    Dim totor As Double
    Dim totpr As Double
    Dim Conta As Integer
    Dim totgen As Double
    Dim totalgen As Double
    Dim totalgenpro As Double
    Dim ConPag As Integer
    ConPag = 1
    totalgen = 0
    totalgenpro = 0
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\RepCuentasPagadas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 8, hCenter
    oDoc.WTextBox 60, 380, 20, 250, "Fecha del " & DTPicker1.Value & " al " & DTPicker2.Value, "F3", 8, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 8, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 8, hCenter
    If Option1.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Ordenes Internacionales Canceladas", "F2", 10, hCenter
    End If
    If Option2.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Ordenes Nacionales Canceladas", "F2", 10, hCenter
    End If
    If Option3.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Ordenes Rapidas Canceladas", "F2", 10, hCenter
    End If
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
    ' Encabezado de pagina
    oDoc.WTextBox 110, 10, 20, 50, "#", "F2", 10, hLeft
    oDoc.WTextBox 110, 40, 20, 50, "F. Orden", "F2", 10, hCenter
    oDoc.WTextBox 110, 90, 20, 250, "Proveedor", "F2", 10, hCenter
    oDoc.WTextBox 110, 340, 20, 60, "Tot. Orden", "F2", 10, hCenter
    oDoc.WTextBox 110, 400, 20, 50, "F. Pago", "F2", 10, hCenter
    oDoc.WTextBox 110, 450, 20, 50, "Cheque", "F2", 10, hCenter
    oDoc.WTextBox 110, 500, 20, 50, "Importe", "F2", 10, hCenter
    ' Cuerpo del reporte
    sumor = 0
    sumpr = 0
    totor = 0
    totpr = 0
    Conta = 0
    sumtoabono = 0
    If Option1.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_ORDEN_COMPRA = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND (ORDEN_COMPRA.CONFIRMADA = 'E') ORDER BY ABONOS_PAGO_OC.NUM_ORDEN"
        StrRep = sBuscar
    End If
    StrRep = sBuscar
    If Option2.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_ORDEN_COMPRA = PROVEEDOR.ID_PROVEEDOR WHERE PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%' AND (ORDEN_COMPRA.CONFIRMADA = 'E') ORDER BY ABONOS_PAGO_OC.NUM_ORDEN"
        StrRep2 = sBuscar
    End If
    If Option3.Value = True Then
        sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA AS NUM_ORDEN, 'R' AS TIPO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA FROM ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY ABONOS_PAGO_OC.NUM_ORDEN"
        StrRep4 = sBuscar
    End If
    Set tRs = cnn.Execute(sBuscar)
    Posi = 120
    Total = 0
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 100
    oDoc.WLineTo 580, 100
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 580, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If sNombre <> tRs.Fields("NOMBRE") Then
                sNombre = tRs.Fields("NOMBRE")
                Conta = 1
                Posi = Posi + 15
                oDoc.WTextBox Posi, 20, 20, 500, tRs.Fields("NOMBRE"), "F2", 9, hLeft
                Posi = Posi + 5
                If Conta = 1 Then
                    Posi = Posi + 6
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 15, Posi
                    oDoc.WLineTo 280, Posi
                    oDoc.LineStroke
                    Posi = Posi + 6
                    Conta = 0
                End If
                If sumor > 0 Then
                    Posi = Posi - 20
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 490, Posi
                    oDoc.WLineTo 560, Posi
                    oDoc.LineStroke
                    oDoc.WTextBox Posi, 510, 40, 1000, Format((sumor), "###,###,###,##0.00"), "F2", 9, hLeft
                    sumor = 0
                    Posi = Posi + 15
                End If
            End If
            Posi = Posi + 10
            Total = CDbl(tRs.Fields("TOT"))
            oDoc.WTextBox Posi, 10, 20, 50, tRs.Fields("NUM_ORDEN"), "F2", 8, hLeft
            oDoc.WTextBox Posi, 40, 20, 50, tRs.Fields("FECHA_ORDEN"), "F2", 8, hLeft
            oDoc.WTextBox Posi, 90, 20, 250, tRs.Fields("NOMBRE"), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("TOT")) Then oDoc.WTextBox Posi, 340, 20, 60, Format(CDbl(tRs.Fields("TOT")), "###,###,###,##0.00"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("FECHA_PAGO")) Then oDoc.WTextBox Posi, 400, 20, 50, tRs.Fields("FECHA_PAGO"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("NO_CHEQUE")) Then oDoc.WTextBox Posi, 450, 20, 50, tRs.Fields("NO_CHEQUE"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then oDoc.WTextBox Posi, 500, 20, 50, Format(CDbl(tRs.Fields("CANT_ABONO")), "###,###,###,##0.00"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then totalgen = totalgen + CDbl(tRs.Fields("CANT_ABONO"))
            sumor = sumor + CDbl(tRs.Fields("CANT_ABONO"))
            If Posi >= 760 Then
                oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 120
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 8, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Fecha del " & DTPicker1.Value & " al " & DTPicker2.Value, "F3", 8, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 8, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 8, hCenter
                If Option4.Value Then
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Ordenes Pagadas (Por Fecha de Orden)", "F2", 10, hCenter
                Else
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Ordenes Pagadas (Por Fecha de Pago)", "F2", 10, hCenter
                End If
                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 110, 10, 20, 50, "#", "F2", 10, hLeft
                oDoc.WTextBox 110, 40, 20, 50, "F. Orden", "F2", 10, hCenter
                oDoc.WTextBox 110, 90, 20, 250, "Proveedor", "F2", 10, hCenter
                oDoc.WTextBox 110, 340, 20, 60, "Tot. Orden", "F2", 10, hCenter
                oDoc.WTextBox 110, 400, 20, 50, "F. Pago", "F2", 10, hCenter
                oDoc.WTextBox 110, 450, 20, 50, "Cheque", "F2", 10, hCenter
                oDoc.WTextBox 110, 500, 20, 50, "Importe", "F2", 10, hCenter
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 100
                oDoc.WLineTo 580, 100
                oDoc.LineStroke
                oDoc.MoveTo 10, 125
                oDoc.WLineTo 580, 125
                oDoc.LineStroke
            End If
            tRs.MoveNext
        Loop
        If sumor > 0 Then
            Posi = Posi + 15
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 490, Posi
            oDoc.WLineTo 560, Posi
            oDoc.LineStroke
            oDoc.WTextBox Posi, 510, 40, 1000, Format((sumor), "###,###,###,##0.00"), "F2", 9, hLeft
            sumor = 0
            Posi = Posi + 15
        End If
        Posi = Posi + 20
        Cont = Cont + 1
        Posi = Posi + 30
        oDoc.WTextBox Posi, 370, 40, 900, "TOTAL GENERAL :", "F2", 9, hLeft
        oDoc.WTextBox Posi, 480, 40, 1000, Format(totalgen, "###,###,###,##0.00"), "F2", 10, hLeft
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub


