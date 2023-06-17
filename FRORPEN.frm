VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRORPEN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ordenes Pedientes de Pago"
   ClientHeight    =   6375
   ClientLeft      =   2445
   ClientTop       =   1170
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   18
      Top             =   5040
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FRORPEN.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FRORPEN.frx":030A
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   16
      Top             =   2640
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FRORPEN.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FRORPEN.frx":26F6
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   13
      Top             =   3840
      Width           =   975
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FRORPEN.frx":4238
         MousePointer    =   99  'Custom
         Picture         =   "FRORPEN.frx":4542
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   9
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
      TabPicture(0)   =   "FRORPEN.frx":6114
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
         Height          =   2295
         Left            =   2640
         TabIndex        =   23
         Top             =   120
         Width           =   5655
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
            Picture         =   "FRORPEN.frx":6130
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Top             =   240
            Width           =   4575
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1095
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1931
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
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clasificacion"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "Orden Rapida"
            Height          =   195
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nacional"
            Height          =   195
            Left            =   480
            TabIndex        =   5
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Internacional"
            Height          =   195
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Indirecta"
            Height          =   195
            Left            =   2880
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Inicio"
         Height          =   195
         Left            =   8340
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango del Reporte"
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   3
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49872897
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
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
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
Attribute VB_Name = "FRORPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim sIdProveedor As String
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
    Dim Dolar As Double
    sBuscar = "SELECT TOP 1 VENTA AS DOLAR From Dolar ORDER BY ID_DOLAR DESC"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
    End If
    ListView1.ListItems.Clear
    If Option1.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ISNULL(SUM(ABONOS_PAGO_OC.CANT_ABONO), 0) AS ABONO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR LEFT OUTER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO AND ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'I') AND (ORDEN_COMPRA.CONFIRMADA IN ('X')) AND (ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
        If sIdProveedor <> "" Then
            sBuscar = sBuscar & " AND ORDEN_COMPRA.ID_PROVEEDOR = " & sIdProveedor
        End If
        sBuscar = sBuscar & " GROUP BY ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA"
        StrRep = sBuscar
    End If
    StrRep = sBuscar
    If Option2.Value = True Then
        sBuscar = "SELECT ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ISNULL(SUM(ABONOS_PAGO_OC.CANT_ABONO), 0) AS ABONO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR LEFT OUTER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO AND ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'N') AND (ORDEN_COMPRA.CONFIRMADA IN ('X')) AND (ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
        If sIdProveedor <> "" Then
            sBuscar = sBuscar & " AND ORDEN_COMPRA.ID_PROVEEDOR = " & sIdProveedor
        End If
        sBuscar = sBuscar & " GROUP BY ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA"
        StrRep2 = sBuscar
    End If
    If Option3.Value = True Then
        sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, ORDEN_RAPIDA.ESTADO, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) - ISNULL((SELECT SUM(CANT_ABONO) AS CANT_ABONO From ABONOS_PAGO_OC WHERE (NUM_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA) AND (TIPO = 'R')), 0) AS DEUDA FROM ORDEN_RAPIDA INNER JOIN  ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (ORDEN_RAPIDA.ESTADO IN ('A', 'M')) "
        If sIdProveedor <> "" Then
            sBuscar = sBuscar & " AND ORDEN_RAPIDA.ID_PROVEEDOR = " & sIdProveedor & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado, ORDEN_RAPIDA.ID_ORDEN_RAPIDA ORDER BY PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA"
        Else
            sBuscar = sBuscar & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado, ORDEN_RAPIDA.ID_ORDEN_RAPIDA ORDER BY  PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA"
        End If
        StrRep4 = sBuscar
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Option3.Value = True Then
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("FECHA")
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    tLi.SubItems(3) = Format(tRs.Fields("DEUDA") * Dolar, "###,###,##0.00")
                Else
                    tLi.SubItems(3) = Format(tRs.Fields("DEUDA"), "###,###,##0.00")
                End If
                tLi.SubItems(4) = tRs.Fields("ESTADO")
                tLi.SubItems(5) = tRs.Fields("MONEDA")
                tRs.MoveNext
            Loop
        End If
    Else
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = tRs.Fields("FECHA")
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    tLi.SubItems(3) = Format(((tRs.Fields("TOTAL") - tRs.Fields("DISCOUNT") + tRs.Fields("FREIGHT") + tRs.Fields("TAX") + tRs.Fields("OTROS_CARGOS")) * Dolar) - tRs.Fields("ABONO"), "###,###,##0.00")
                Else
                    tLi.SubItems(3) = Format(tRs.Fields("TOTAL") - tRs.Fields("DISCOUNT") + tRs.Fields("FREIGHT") + tRs.Fields("TAX") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("ABONO"), "###,###,##0.00")
                End If
                tLi.SubItems(4) = tRs.Fields("NUM_ORDEN")
                tLi.SubItems(5) = tRs.Fields("CONFIRMADA")
                tRs.MoveNext
            Loop
        End If
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
        .ColumnHeaders.Add , , "Id.", 1000
        .ColumnHeaders.Add , , "Nombre", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Total", 1200
        .ColumnHeaders.Add , , "Num_Orden", 1500
        .ColumnHeaders.Add , , "Tipo Confirmada", 1000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id.", 1000
        .ColumnHeaders.Add , , "Nombre", 3000
    End With
    If Option3.Value = True Then
        With ListView1
            .View = lvwReport
            .GridLines = True
            .LabelEdit = lvwManual
            .HideSelection = False
            .HotTracking = False
            .HoverSelection = False
            .ColumnHeaders.Add , , "Id.", 1000
            .ColumnHeaders.Add , , "Nombre", 1200
            .ColumnHeaders.Add , , "Fecha", 1200
            .ColumnHeaders.Add , , "Total", 1200
            .ColumnHeaders.Add , , "Estado", 1500
            .ColumnHeaders.Add , , "Moneda", 1000
            .ColumnHeaders.Add , , "Factura", 1000
            .ColumnHeaders.Add , , "Total_Factura", 1000
            .ColumnHeaders.Add , , "Num_Entrada", 1000
        End With
    End If
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
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = ListView2.SelectedItem.SubItems(1)
    Text1.SetFocus
    sIdProveedor = Item
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
            sqlQuery = "SELECT * FROM PROVEEDOR_CONSUMO WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        Else
            sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
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
    Dim sDolar As Double
    ConPag = 1
    totalgen = 0
    totalgenpro = 0
    sBuscar = "SELECT TOP 1 VENTA FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sDolar = tRs.Fields("VENTA")
    End If
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\Repcuentasporpagar.pdf") Then
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
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    If Option1.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar (Internacionales)", "F2", 10, hCenter
    End If
    If Option2.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar (Nacionales)", "F2", 10, hCenter
    End If
    If Option3.Value Then
        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar (Rapidas)", "F2", 10, hCenter
    End If
    oDoc.WTextBox 60, 380, 20, 250, "Fecha del " & DTPicker1.Value & " al " & DTPicker2.Value, "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
    ' Encabezado de pagina
    ' Cuerpo del reporte
    sumor = 0
    sumpr = 0
    totor = 0
    totpr = 0
    Conta = 0
    sumtoabono = 0
    If Option2.Value Or Option1.Value Then
        'oDoc.WTextBox 160, 20, 20, 100, "Nombre", "F2", 10, hCenter
        oDoc.WTextBox 110, 10, 10, 50, "Num Orden", "F2", 8, hLeft
        oDoc.WTextBox 110, 60, 20, 50, "Fecha", "F2", 8, hCenter
        oDoc.WTextBox 110, 120, 20, 50, "Subtotal", "F2", 8, hCenter
        oDoc.WTextBox 110, 170, 20, 50, "Descuento", "F2", 8, hCenter
        oDoc.WTextBox 110, 210, 20, 50, "Flete", "F2", 8, hCenter
        oDoc.WTextBox 110, 20, 20, 510, "Otros Car.", "F2", 8, hCenter
        oDoc.WTextBox 110, 10, 20, 610, "Tax", "F2", 8, hCenter
        oDoc.WTextBox 110, 20, 20, 740, "No Factura", "F2", 8, hCenter
        oDoc.WTextBox 110, 20, 30, 900, "Total de Orden", "F2", 8, hCenter
        oDoc.WTextBox 110, 20, 30, 1040, "Total de Proveedor", "F2", 8, hCenter
        If Option1.Value = True Then
            sBuscar = "SELECT ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ISNULL(SUM(ABONOS_PAGO_OC.CANT_ABONO), 0) AS ABONO, ORDEN_COMPRA.FACT_PROVE, ORDEN_COMPRA.TOT_PRO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR LEFT OUTER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO AND ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'I') AND (ORDEN_COMPRA.CONFIRMADA IN ('X')) AND (ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"""
        End If
        If Option2.Value Then
            sBuscar = "SELECT ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ISNULL(SUM(ABONOS_PAGO_OC.CANT_ABONO), 0) AS ABONO, ORDEN_COMPRA.FACT_PROVE, ORDEN_COMPRA.TOT_PRO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR LEFT OUTER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO AND ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'N') AND (ORDEN_COMPRA.CONFIRMADA IN ('X')) AND (ORDEN_COMPRA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
        End If
        If sIdProveedor <> "" Then
            sBuscar = sBuscar & " AND ORDEN_COMPRA.ID_PROVEEDOR = " & sIdProveedor & " GROUP BY ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.FACT_PROVE, ORDEN_COMPRA.TOT_PRO ORDER BY PROVEEDOR.NOMBRE"
        Else
            sBuscar = sBuscar & " GROUP BY ORDEN_COMPRA.ID_PROVEEDOR, PROVEEDOR.NOMBRE, ORDEN_COMPRA.FECHA, ORDEN_COMPRA.MONEDA, ORDEN_COMPRA.TOTAL, ORDEN_COMPRA.DISCOUNT, ORDEN_COMPRA.FREIGHT, ORDEN_COMPRA.TAX, ORDEN_COMPRA.OTROS_CARGOS, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.ID_ORDEN_COMPRA, ORDEN_COMPRA.FACT_PROVE, ORDEN_COMPRA.TOT_PRO ORDER BY PROVEEDOR.NOMBRE"
        End If
        Set tRs = cnn.Execute(sBuscar)
        Posi = 120
        Total = 0
        totalgen = 0
        totalgenpro = 0
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 110
        oDoc.WLineTo 580, 110
        oDoc.LineStroke
        oDoc.MoveTo 10, 130
        oDoc.WLineTo 580, 130
        oDoc.LineStroke
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                If sNombre <> tRs.Fields("NOMBRE") Then
                    Conta = 1
                    Posi = Posi + 15
                    oDoc.WTextBox Posi, 20, 20, 500, tRs.Fields("NOMBRE"), "F2", 9, hLeft
                    If Conta = 1 Then
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 15, Posi
                        oDoc.WLineTo 280, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        Conta = 0
                    End If
                    If sumor > 0 Then
                        Posi = Posi - 30
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 510, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        oDoc.WTextBox Posi, 440, 30, 130, Format(sumor, "###,###,##0.00"), "F2", 9, hRight
                        sumor = 0
                        Posi = Posi + 15
                    End If
                    If sumpr > 0 Then
                        Posi = Posi - 15
                        oDoc.WTextBox Posi, 480, 30, 1000, Format(sumpr, "###,###,##0.00"), "F2", 9, hCenter
                        sumpr = 0
                        Posi = Posi + 15
                    End If
                End If
                Posi = Posi + 10
                oDoc.WTextBox Posi, 20, 30, 30, tRs.Fields("NUM_ORDEN"), "F2", 8, hLeft
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    If Not IsNull(tRs.Fields("FECHA")) Then oDoc.WTextBox Posi, 40, 30, 60, Format(tRs.Fields("FECHA"), "dd/mm/yyyy"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TOTAL")) Then oDoc.WTextBox Posi, 100, 30, 60, Format(tRs.Fields("TOTAL") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("DISCOUNT")) Then oDoc.WTextBox Posi, 160, 30, 50, Format(tRs.Fields("DISCOUNT") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("FREIGHT")) Then oDoc.WTextBox Posi, 190, 30, 60, Format(tRs.Fields("FREIGHT") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("OTROS_CARGOS")) Then oDoc.WTextBox Posi, 230, 30, 60, Format(tRs.Fields("OTROS_CARGOS") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TAX")) Then oDoc.WTextBox Posi, 270, 30, 70, Format(tRs.Fields("TAX") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("FACT_PROVE")) Then oDoc.WTextBox Posi, 325, 30, 60, tRs.Fields("FACT_PROVE"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TOT_PRO")) Then oDoc.WTextBox Posi, 380, 30, 60, Format(tRs.Fields("TOT_PRO") * sDolar, "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TOT_PRO")) Then oDoc.WTextBox Posi, 500, 30, 60, Format((CDbl(tRs.Fields("TOTAL")) - CDbl(tRs.Fields("DISCOUNT")) + CDbl(tRs.Fields("FREIGHT")) + CDbl(tRs.Fields("OTROS_CARGOS")) + CDbl(tRs.Fields("TAX"))), "###,###,###,##0.00") & " USD", "F2", 8, hRight
                    PosVer = Posi
                    Total = (CDbl(tRs.Fields("TOTAL")) - CDbl(tRs.Fields("DISCOUNT")) + CDbl(tRs.Fields("FREIGHT")) + CDbl(tRs.Fields("OTROS_CARGOS")) + CDbl(tRs.Fields("TAX"))) * sDolar
                    sumor = CDbl(sumor) + CDbl(Total)
                    totalgen = CDbl(totalgen) + CDbl(Total)
                    If Not IsNull(tRs.Fields("TOT_PRO")) Then totalgenpro = CDbl(totalgenpro) + CDbl(tRs.Fields("TOT_PRO") * sDolar)
                    If Total > 0 Then
                        oDoc.WTextBox Posi, 440, 30, 60, Format((Total), "###,###,###,##0.00"), "F2", 9, hRight
                        Total = 0
                        Posi = Posi + 15
                    End If
                    sNombre = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("TOT_PRO")) Then sumpr = CDbl(sumpr) + CDbl(tRs.Fields("TOT_PRO") * sDolar)
                Else
                    If Not IsNull(tRs.Fields("FECHA")) Then oDoc.WTextBox Posi, 40, 30, 60, Format(tRs.Fields("FECHA"), "dd/mm/yyyy"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TOTAL")) Then oDoc.WTextBox Posi, 100, 30, 60, Format(tRs.Fields("TOTAL"), "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("DISCOUNT")) Then oDoc.WTextBox Posi, 160, 30, 50, Format(tRs.Fields("DISCOUNT"), "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("FREIGHT")) Then oDoc.WTextBox Posi, 190, 30, 60, Format(tRs.Fields("FREIGHT"), "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("OTROS_CARGOS")) Then oDoc.WTextBox Posi, 280, 30, 60, Format(tRs.Fields("OTROS_CARGOS"), "###,###,###,##0.00"), "F2", 8, hLeft
                    If Not IsNull(tRs.Fields("TAX")) Then oDoc.WTextBox Posi, 270, 30, 70, Format(tRs.Fields("TAX"), "###,###,###,##0.00"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("FACT_PROVE")) Then oDoc.WTextBox Posi, 365, 30, 70, tRs.Fields("FACT_PROVE"), "F2", 8, hRight
                    If Not IsNull(tRs.Fields("TOTAL")) Then oDoc.WTextBox Posi, 420, 30, 80, Format(tRs.Fields("TOTAL") - tRs.Fields("DISCOUNT") + tRs.Fields("FREIGHT") + tRs.Fields("TAX") + tRs.Fields("OTROS_CARGOS"), "###,###,##0.00"), "F2", 8, hRight
                    PosVer = Posi
                    Total = CDbl(tRs.Fields("TOTAL")) - CDbl(tRs.Fields("DISCOUNT")) + CDbl(tRs.Fields("FREIGHT")) + CDbl(tRs.Fields("OTROS_CARGOS")) + CDbl(tRs.Fields("TAX") - tRs.Fields("ABONO"))
                    sumor = CDbl(sumor) + CDbl(Total)
                    totalgen = CDbl(totalgen) + CDbl(Total)
                    If Not IsNull(tRs.Fields("TOTAL")) Then totalgenpro = Format(tRs.Fields("TOTAL") - tRs.Fields("DISCOUNT") + tRs.Fields("FREIGHT") + tRs.Fields("TAX") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("ABONO"), "###,###,##0.00")
                    If Total > 0 Then
                        oDoc.WTextBox Posi, 510, 30, 60, Format((Total), "###,###,###,##0.00"), "F2", 9, hRight
                        Total = 0
                        Posi = Posi + 15
                    End If
                    sNombre = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("TOTAL")) Then sumpr = CDbl(sumpr) + Format(tRs.Fields("TOTAL") - tRs.Fields("DISCOUNT") + tRs.Fields("FREIGHT") + tRs.Fields("TAX") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("ABONO"), "0.00")
                End If
                tRs.MoveNext
                If Posi >= 700 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    ' Encabezado del reporte
                    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                    oDoc.WTextBox 60, 380, 20, 250, "Fecha del " & DTPicker1.Value & " al " & DTPicker2.Value, "F3", 8, hCenter
                    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                    ' Encabezado de pagina
                    oDoc.WTextBox 110, 10, 10, 50, "Num Orden", "F2", 8, hLeft
                    oDoc.WTextBox 110, 60, 20, 50, "Fecha", "F2", 8, hCenter
                    oDoc.WTextBox 110, 120, 20, 50, "Subtotal", "F2", 8, hCenter
                    oDoc.WTextBox 110, 170, 20, 50, "Descuento", "F2", 8, hCenter
                    oDoc.WTextBox 110, 210, 20, 50, "Flete", "F2", 8, hCenter
                    oDoc.WTextBox 110, 20, 20, 510, "Otros Car.", "F2", 8, hCenter
                    oDoc.WTextBox 110, 10, 20, 610, "Tax", "F2", 8, hCenter
                    oDoc.WTextBox 110, 20, 20, 740, "No Factura", "F2", 8, hCenter
                    oDoc.WTextBox 110, 20, 30, 900, "Total de Orden", "F2", 8, hCenter
                    oDoc.WTextBox 110, 20, 30, 1040, "Total de Proveedor", "F2", 8, hCenter
                    Posi = 180
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, 110
                    oDoc.WLineTo 580, 110
                    oDoc.LineStroke
                    oDoc.MoveTo 10, 130
                    oDoc.WLineTo 580, 130
                    oDoc.LineStroke
                End If
            Loop
            Conta = 1
            Posi = Posi + 35
            If sumor > 0 Then
                Posi = Posi - 30
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 510, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                oDoc.WTextBox Posi, 440, 30, 130, Format(sumor, "###,###,##0.00"), "F2", 9, hRight
                sumor = 0
                Posi = Posi + 15
            End If
            If sumpr > 0 Then
                Posi = Posi - 15
                oDoc.WTextBox Posi, 710, 30, 1000, Format(sumpr, "###,###,##0.00"), "F2", 9, hCenter
                sumpr = 0
                Posi = Posi + 15
            End If
            Posi = Posi + 30
            oDoc.WTextBox Posi, 20, 20, 720, "Total en Ordenes", "F2", 9, hCenter
            oDoc.WTextBox Posi, 20, 20, 1040, Format(totalgen, "###,###,##0.00"), "F2", 9, hCenter

            Cont = Cont + 1
        End If
    End If
    If Option3.Value = True Then
        'oDoc.WTextBox 160, 20, 20, 100, "Nombre", "F2", 10, hCenter
        oDoc.WTextBox 110, 20, 20, 100, "ID_ORDEN_RAPIDA", "F2", 10, hCenter
        oDoc.WTextBox 110, 20, 20, 360, "FECHA", "F2", 10, hCenter
        oDoc.WTextBox 110, 20, 20, 700, "MONEDA", "F2", 10, hCenter
        oDoc.WTextBox 110, 20, 20, 950, "TOTAL", "F2", 10, hCenter
        Dim totreb As Double
        'sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, ORDEN_RAPIDA.Estado FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (ORDEN_RAPIDA.ESTADO IN ('A', 'M')) "
        'If sIdProveedor <> "" Then
        '    sBuscar = sBuscar & " AND ORDEN_RAPIDA.ID_PROVEEDOR = " & sIdProveedor & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado ORDER BY PROVEEDOR_CONSUMO.NOMBRE"
        'Else
        '    sBuscar = sBuscar & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado ORDER BY PROVEEDOR_CONSUMO.NOMBRE"
        'End If
        
        sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, ORDEN_RAPIDA.ESTADO, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) - ISNULL((SELECT SUM(CANT_ABONO) AS CANT_ABONO From ABONOS_PAGO_OC WHERE (NUM_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA) AND (TIPO = 'R')), 0) AS DEUDA FROM ORDEN_RAPIDA INNER JOIN  ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (ORDEN_RAPIDA.ESTADO IN ('A', 'M')) "
        If sIdProveedor <> "" Then
            sBuscar = sBuscar & " AND ORDEN_RAPIDA.ID_PROVEEDOR = " & sIdProveedor & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado, ORDEN_RAPIDA.ID_ORDEN_RAPIDA ORDER BY PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA"
        Else
            sBuscar = sBuscar & " GROUP BY ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_PROVEEDOR, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA , PROVEEDOR_CONSUMO.Nombre, ORDEN_RAPIDA.fecha, ORDEN_RAPIDA.Estado, ORDEN_RAPIDA.ID_ORDEN_RAPIDA ORDER BY PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA"
        End If
        Set tRs = cnn.Execute(sBuscar)
        Posi = 120
        Total = 0
        sumor = 0
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
                    Conta = 1
                    Posi = Posi + 25
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
                        Posi = Posi - 30
                        Posi = Posi + 3
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 450, Posi
                        oDoc.WLineTo 550, Posi
                        oDoc.LineStroke
                        Posi = Posi + 3
                        oDoc.WTextBox Posi, 480, 40, 1000, Format((sumor), "###,###,###,##0.00"), "F2", 10, hLeft
                        sumor = 0
                        Posi = Posi + 15
                    End If
                End If
                Posi = Posi + 15
                oDoc.WTextBox Posi, 20, 20, 100, tRs.Fields("ID_ORDEN_RAPIDA"), "F2", 10, hCenter
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    sumtoabono = Format(CDbl(tRs.Fields("DEUDA")) * sDolar, "###,###,###,##0.00")
                Else
                    sumtoabono = Format(CDbl(tRs.Fields("DEUDA")), "###,###,###,##0.00")
                End If
                oDoc.WTextBox Posi, 20, 20, 350, tRs.Fields("FECHA"), "F2", 8, hCenter
                oDoc.WTextBox Posi, 20, 20, 700, Format(tRs.Fields("MONEDA"), "###,###,###,##0.00"), "F2", 8, hCenter
                oDoc.WTextBox Posi, 430, 40, 100, Format(sumtoabono, "###,###,###,##0.00"), "F2", 8, hRight
                PosVer = Posi
                sumor = CDbl(sumor) + CDbl(sumtoabono)
                totgen = CDbl(totgen) + CDbl(sumtoabono)
                sNombre = tRs.Fields("NOMBRE")
                tRs.MoveNext
                If Posi >= 700 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    ' Encabezado del reporte
                    Posi = 120
                    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                    ' Encabezado de pagina
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, 100
                    oDoc.WLineTo 580, 100
                    oDoc.LineStroke
                    oDoc.MoveTo 10, 125
                    oDoc.WLineTo 580, 125
                    oDoc.LineStroke
                    oDoc.WTextBox 110, 20, 20, 100, "ID_ORDEN_RAPIDA", "F2", 10, hCenter
                    oDoc.WTextBox 110, 20, 20, 360, "FECHA", "F2", 10, hCenter
                    oDoc.WTextBox 110, 20, 20, 700, "MONEDA", "F2", 10, hCenter
                    oDoc.WTextBox 110, 20, 20, 950, "TOTAL", "F2", 10, hCenter
                End If
            Loop
            Conta = 1
            Posi = Posi + 36
            If sumor > 0 Then
                Posi = Posi - 30
                Posi = Posi + 3
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 450, Posi
                oDoc.WLineTo 550, Posi
                oDoc.LineStroke
                Posi = Posi + 3
                'oDoc.WTextBox Posi, 340, 20, 700, Total, "F2", 9, hLeft
                oDoc.WTextBox Posi, 480, 40, 1000, Format((sumor), "###,###,###,##0.00"), "F2", 10, hLeft
                sumor = 0
                Posi = Posi + 15
            End If
            Posi = Posi + 30
            oDoc.WTextBox Posi, 370, 40, 900, "TOTAL GENERAL :", "F2", 9, hLeft
            oDoc.WTextBox Posi, 480, 40, 1000, Format(totgen, "###,###,###,##0.00"), "F2", 10, hLeft
            Cont = Cont + 1
        End If
    End If
    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
    oDoc.PDFClose
    oDoc.Show
End Sub
