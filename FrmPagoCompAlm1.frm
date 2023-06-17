VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmPagoCompAlm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pagos a Proveedores (Almacen 1)"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   22
      Top             =   5640
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   21
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox txtTrans 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   20
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtCheque 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   6120
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9480
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   16
      Top             =   3120
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmPagoCompAlm1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoCompAlm1.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   14
      Top             =   4320
      Width           =   975
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptar"
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmPagoCompAlm1.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoCompAlm1.frx":2156
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   12
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmPagoCompAlm1.frx":3C08
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoCompAlm1.frx":3F12
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6480
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Pagos Pendientes"
      TabPicture(0)   =   "FrmPagoCompAlm1.frx":5FF4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmPagoCompAlm1.frx":6010
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(2)=   "Text1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "No Aprobados"
      TabPicture(2)   =   "FrmPagoCompAlm1.frx":602C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "ListView3"
      Tab(2).Control(2)=   "Text2"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -67680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4920
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7646
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -67680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8281
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
         Caption         =   "Importe"
         Height          =   255
         Left            =   -68400
         TabIndex        =   7
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Total a Pagar"
         Height          =   255
         Left            =   -68880
         TabIndex        =   3
         Top             =   4920
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Banco"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo de Pago"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. de Transferencia"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. de Cheque"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total a Pagar :"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   615
   End
End
Attribute VB_Name = "FrmPagoCompAlm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim NomProv As String
Dim NoFolio As String
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
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
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Nombre", 5950
        .ColumnHeaders.Add , , "Telefono", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Total", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Revision", 0
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Id Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Cant. Minima", 1000
        .ColumnHeaders.Add , , "Cant. Maxima", 1000
        .ColumnHeaders.Add , , "Proveedor", 3000
        .ColumnHeaders.Add , , "Precio Ofertado", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Revision", 0
        .ColumnHeaders.Add , , "Folio", 0
        .ColumnHeaders.Add , , "Id Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Cant. Minima", 1000
        .ColumnHeaders.Add , , "Cant. Maxima", 1000
        .ColumnHeaders.Add , , "Proveedor", 3000
        .ColumnHeaders.Add , , "Precio Ofertado", 1000
    End With
    sBuscar = "SELECT * FROM BANCOS"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo2.Clear
            Do While Not .EOF
                Combo2.AddItem (.Fields("NOMBRE"))
                .MoveNext
            Loop
        Else
            MsgBox "NO EXISTEN BANCOS REGISTRADOS, NO PUEDE REGISTRAR PAGOS", vbInformation, "SACC"
        End If
        .Close
    End With
    sBuscar = "SELECT * FROM TPAGOS_OC"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo1.Clear
            Do While Not .EOF
                Combo1.AddItem (.Fields("Descripcion"))
                .MoveNext
            Loop
        Else
            MsgBox "FALLO DE INFORMACION, FAVOR DE LLAMAR A SOPORTE", vbInformation, "SACC"
        End If
        .Close
    End With
    sBuscar = "UPDATE REV_COMPRA_ALMACEN1 SET APROVADO = 'F' WHERE (PRECIO_COMPRA = 0) AND (APROVADO = 'A')"
    cnn.Execute (sBuscar)
    BuscarPagos
End Sub
Private Sub Image1_Click()
    Dim sBuscar As String
    Dim Cont As Integer
    Dim NReg As Integer
    Dim tRs As ADODB.Recordset
    NReg = ListView2.ListItems.Count
    For Cont = 1 To NReg
        sBuscar = "UPDATE REV_COMPRA_ALMACEN1 SET APROVADO = 'F', FECHA_APROVADO = '" & Format(Date, "dd/mm/yyyy") & "' WHERE ID_REVISION = " & ListView2.ListItems(Cont)
        Set tRs = cnn.Execute(sBuscar)
    Next Cont
    ListView1.ListItems.Clear
    BuscarPagos
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
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
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.Count
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ProgressBar1.Visible = False
        ProgressBar1.Value = 0
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub imgLeer_Click()
    If Not ArchivoEnUso(App.Path & "\Cheque.pdf") Then
        If MsgBox("ESTA POR CERRAR LA ORDEN DE COMPRA SELECCIONADA, REGISTRARA UN PAGO ¿ESTA SEGURO QUE DESEA CONTINUAR?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            Dim sBuscar As String
            Dim Cont As Integer
            Dim tRs As ADODB.Recordset
            FrmCheque.TxtNUM_ORDEN.Text = NoFolio 'numero de orden de compra
            FrmCheque.TxtTIPO_ORDEN.Text = "ALMACEN1" 'tipo de orden de compra
            FrmCheque.txtNum2Let(0).Text = Text4.Text 'total de la orden de compra
            FrmCheque.TxtNOMBRE.Text = NomProv 'nombre del proveedor a recibir el pago
            FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text
            FrmCheque.Combo1.Text = Combo2.Text
            FrmCheque.Show vbModal
            sBuscar = "UPDATE REV_COMPRA_ALMACEN1 SET APROVADO = 'F', FECHA_APROVADO = '" & Format(Date, "dd/mm/yyyy") & "' WHERE GRUPO = " & NoFolio
            Set tRs = cnn.Execute(sBuscar)
            ListView1.ListItems.Clear
            BuscarPagos
            NoFolio = ""
            NomProv = ""
            ListView2.ListItems.Clear
            ListView3.ListItems.Clear
        End If
    Else
        MsgBox "Cierre la poliza de cheque que tiene abierta para poder continuar", vbExclamation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim TotPaga As String
    TotPaga = "0"
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    NomProv = Item.SubItems(2)
    NoFolio = Item.SubItems(1)
    sBuscar = "SELECT * FROM VsComprasAlm1reporte WHERE FECHA = '" & Item.SubItems(4) & "' AND ID_PROVEEDOR = " & Item & " AND GRUPO = '" & Item.SubItems(1) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_REVISION") & "")
                If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(1) = tRs.Fields("GRUPO")
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(3) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD_APROVADA")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD_APROVADA")
                If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(5) = tRs.Fields("EXISTENCIA")
                If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(6) = tRs.Fields("C_MINIMA")
                If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(7) = tRs.Fields("C_MAXIMA")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(8) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then tLi.SubItems(9) = tRs.Fields("PRECIO_COMPRA")
                If Not IsNull(tRs.Fields("CANTIDAD_APROVADA")) And Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then
                 TotPaga = Format(CDbl(TotPaga) + (CDbl(tRs.Fields("PRECIO_COMPRA")) * CDbl(tRs.Fields("CANTIDAD_APROVADA"))), "###,###,##0.00")
                End If
                If tRs.Fields("APROVADO") = "R" Then
                    ListView2.ListItems(ListView2.ListItems.Count).ForeColor = &HC0C0FF
                End If
            tRs.MoveNext
        Loop
        Text1.Text = TotPaga
        Text4.Text = TotPaga
    End If
    tRs.MoveFirst
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_REVISION") & "")
            If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(1) = tRs.Fields("GRUPO")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(3) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD_APROVADA")) And Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD") - tRs.Fields("CANTIDAD_APROVADA")
            If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(5) = tRs.Fields("EXISTENCIA")
            If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(6) = tRs.Fields("C_MINIMA")
            If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(7) = tRs.Fields("C_MAXIMA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(8) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then tLi.SubItems(9) = tRs.Fields("PRECIO_COMPRA")
            If Not IsNull(tRs.Fields("CANTIDAD_APROVADA")) And Not IsNull(tRs.Fields("PRECIO_COMPRA")) Then
                If CDbl(tRs.Fields("CANTIDAD")) - CDbl(tRs.Fields("CANTIDAD_APROVADA")) > 0 Then
                    Text2.Text = Format((CDbl(TotPaga) + CDbl(tRs.Fields("PRECIO_COMPRA")) * (CDbl(tRs.Fields("CANTIDAD")) - CDbl(tRs.Fields("CANTIDAD_APROVADA")))), "###,###,##0.00")
                Else
                    Text2.Text = "0.00"
                End If
            Else
                Text2.Text = "0.00"
            End If
            tRs.MoveNext
        Loop
    Text3.Text = Format(CDbl(Text2.Text) + CDbl(TotPaga), "###,###,##0.00")
    End If
End Sub
Public Sub BuscarPagos()
    sBuscar = "SELECT NOMBRE, FECHA, ID_PROVEEDOR, TELEFONO, GRUPO, SUM(TOTAL) AS TOTAL FROM VsGroupComAlma1 WHERE APROVADO = 'A' GROUP BY NOMBRE, FECHA, ID_PROVEEDOR, TELEFONO, GRUPO ORDER BY GRUPO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR") & "")
            If Not IsNull(tRs.Fields("GRUPO")) Then tLi.SubItems(1) = tRs.Fields("GRUPO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(3) = tRs.Fields("TELEFONO")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(5) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    On Error Resume Next
    filenum = FreeFile()
    Open sFileName For Input Lock Read As #filenum
    Close filenum
    errnum = Err
    On Error GoTo 0
    Select Case errnum
        Case 0
            ArchivoEnUso = False
        Case 70
            ArchivoEnUso = True
        Case Else
            Error errnum
    End Select
End Function
