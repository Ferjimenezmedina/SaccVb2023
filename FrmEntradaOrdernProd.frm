VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEntradaOrdernProd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Entrada de Ordenes de Producción"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   10560
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16318465
      CurrentDate     =   43606
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   10
      Top             =   4320
      Width           =   975
      Begin VB.Label Label11 
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image11 
         Height          =   675
         Left            =   120
         MouseIcon       =   "FrmEntradaOrdernProd.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmEntradaOrdernProd.frx":030A
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   255
      Left            =   10680
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   4
      Top             =   6360
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmEntradaOrdernProd.frx":1A80
         MousePointer    =   99  'Custom
         Picture         =   "FrmEntradaOrdernProd.frx":1D8A
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Producciones Terminadas"
      TabPicture(0)   =   "FrmEntradaOrdernProd.frx":3E6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Pendientes de Entrega a Almacén"
      TabPicture(1)   =   "FrmEntradaOrdernProd.frx":3E88
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "Recibir"
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
         Left            =   9120
         Picture         =   "FrmEntradaOrdernProd.frx":3EA4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6960
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5655
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9975
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Left            =   3960
         Picture         =   "FrmEntradaOrdernProd.frx":6876
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11880
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Orden de Produccion :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16318465
      CurrentDate     =   43606
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Al:"
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Del:"
      Height          =   255
      Left            =   10560
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "FrmEntradaOrdernProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim sId As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT COMANDAS_DETALLES_2.I_S, COMANDAS_DETALLES_2.ID_COMANDA, COMANDAS_DETALLES_2.ID_PRODUCTO, ALMACEN3.Descripcion, SUM(COMANDAS_DETALLES_2.CANTIDAD - COMANDAS_DETALLES_2.CANTIDAD_NO_SIRVIO) AS CANTIDAD, COMANDAS_DETALLES_2.FECHA_FIN FROM COMANDAS_DETALLES_2 INNER JOIN ALMACEN3 ON COMANDAS_DETALLES_2.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (COMANDAS_DETALLES_2.CLASIFICACION = 'P') AND (COMANDAS_DETALLES_2.ESTADO_ACTUAL IN ('L', 'N')) AND ID_COMANDA = '" & Text1.Text & "' GROUP BY COMANDAS_DETALLES_2.ID_COMANDA, COMANDAS_DETALLES_2.ID_PRODUCTO, ALMACEN3.Descripcion, COMANDAS_DETALLES_2.FECHA_FIN, COMANDAS_DETALLES_2.I_S"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(2) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(4) = tRs.Fields("FECHA_FIN")
            If Not IsNull(tRs.Fields("I_S")) Then tLi.SubItems(5) = tRs.Fields("I_S")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    If ListView2.ListItems.Count > 0 Then
        Dim Cont As Integer
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sId = ""
        Cont = 1
        Do While Cont <= ListView2.ListItems.Count
            If ListView2.ListItems(Cont).Checked Then
                sBuscar = "SELECT  CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView2.ListItems(Cont).SubItems(1) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.BOF And tRs.EOF) Then
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & ListView2.ListItems(Cont).SubItems(3) & " WHERE SUCURSAL = '" & VarMen.Text4(0).Text & " ' AND ID_PRODUCTO = '" & ListView2.ListItems(Cont).SubItems(1) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                Else
                    sBuscar = "INSERT INTO EXISTENCIAS(ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES('" & ListView2.ListItems(Cont).SubItems(1) & "','" & ListView2.ListItems(Cont).SubItems(3) & "', '" & VarMen.Text4(0).Text & "')"
                    Set tRs = cnn.Execute(sBuscar)
                End If
                sBuscar = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'I', FECHA_FIN = '" & Date & "' WHERE ID_COMANDA = " & ListView2.ListItems(Cont) & " AND ID_PRODUCTO = '" & ListView2.ListItems(Cont).SubItems(1) & "'"
                cnn.Execute (sBuscar)
            End If
            sId = sId & ListView2.ListItems(Cont).SubItems(5) & ", "
            Cont = Cont + 1
        Loop
        If Len(sId) > 1 Then
            sId = Mid(sId, 1, Len(sId) - 2)
        End If
        Text1.Text = ListView2.ListItems(1)
        cmdBuscar.Value = True
        ImrpimeEntrada
        Actualiza
    End If
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
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
        .ColumnHeaders.Add , , "No. Comanda", 1000
        .ColumnHeaders.Add , , "Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Fecha Finalizo", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Comanda", 1000
        .ColumnHeaders.Add , , "Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Fecha Finalizo", 1500
        .ColumnHeaders.Add , , "ID", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Comanda", 1000
        .ColumnHeaders.Add , , "Producto", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Fecha Finalizo", 1500
    End With
    Actualiza
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT COMANDAS_DETALLES_2.ID_COMANDA, COMANDAS_DETALLES_2.ID_PRODUCTO, ALMACEN3.Descripcion, SUM(COMANDAS_DETALLES_2.CANTIDAD - COMANDAS_DETALLES_2.CANTIDAD_NO_SIRVIO) AS CANTIDAD, COMANDAS_DETALLES_2.FECHA_FIN FROM COMANDAS_DETALLES_2 INNER JOIN ALMACEN3 ON COMANDAS_DETALLES_2.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN COMANDAS_2 ON COMANDAS_DETALLES_2.ID_COMANDA = COMANDAS_2.ID_COMANDA WHERE (COMANDAS_DETALLES_2.ESTADO_ACTUAL IN ('L', 'N')) AND (COMANDAS_2.TIPO = 'P') AND (COMANDAS_DETALLES_2.CANTIDAD - COMANDAS_DETALLES_2.CANTIDAD_NO_SIRVIO > 0) GROUP BY COMANDAS_DETALLES_2.ID_COMANDA, COMANDAS_DETALLES_2.ID_PRODUCTO, ALMACEN3.Descripcion, COMANDAS_DETALLES_2.FECHA_FIN, COMANDAS_2.TIPO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(2) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(4) = tRs.Fields("FECHA_FIN")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image11_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView3.ListItems.Clear
    sBuscar = "SELECT COMANDAS_2.ID_COMANDA, COMANDAS_DETALLES_2.ID_PRODUCTO, COMANDAS_DETALLES_2.CANTIDAD, COMANDAS_DETALLES_2.FECHA_FIN FROM COMANDAS_2 INNER JOIN COMANDAS_DETALLES_2 ON COMANDAS_2.ID_COMANDA = COMANDAS_DETALLES_2.ID_COMANDA WHERE (COMANDAS_2.TIPO = 'P') AND (COMANDAS_DETALLES_2.ESTADO_ACTUAL = 'I') AND (COMANDAS_DETALLES_2.FECHA_FIN BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(3) = tRs.Fields("FECHA_FIN")
            tRs.MoveNext
        Loop
    End If
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
    If ListView3.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView3.ColumnHeaders.Count
            For Con = 1 To ListView3.ColumnHeaders.Count
                StrCopi = StrCopi & ListView3.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView3.ListItems.Count
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub ImrpimeEntrada()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM COMANDAS_DETALLES_2 WHERE I_S IN (" & sId & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Dim POSY As Integer
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "             ORDEN DE PRODUCCION : " & tRs.Fields("ID_COMANDA")
        Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE ORDEN DE PRODUCCION")) / 2
        Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE ORDEN DE PRODUCCION"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = 3000
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Clave del Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3300
        Printer.Print "Cant. Registrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 4100
        Printer.Print "Sucursal"
        Printer.CurrentY = POSY
        Printer.CurrentX = 5300
        Printer.Print "Orden"
        POSY = POSY + 200
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 3300
            Printer.Print tRs.Fields("CANTIDAD") - tRs.Fields("CANTIDAD_NO_SIRVIO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 4100
            Printer.Print VarMen.Text4(0).Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 5500
            Printer.Print tRs.Fields("ID_COMANDA")
            If POSY >= 14200 Then
                POSY = 100
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "             ORDEN DE PRODUCCION : " & tRs.Fields("ID_COMANDA")
                Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE ORDEN DE PRODUCCION")) / 2
                Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE ORDEN DE PRODUCCION"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 3000
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Clave del Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 3300
                Printer.Print "Cant. Registrada"
                Printer.CurrentY = POSY
                Printer.CurrentX = 4100
                Printer.Print "Sucursal"
                Printer.CurrentY = POSY
                Printer.CurrentX = 5300
                Printer.Print "Orden"
                POSY = POSY + 200
            End If
            tRs.MoveNext
        Loop
        Printer.Print "FIN DEL LISTADO"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    Else
        MsgBox "NO SE ENCONTRO EL REGISTRO DE LA ENTRADA DE LOS PRODUCTOS", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub

