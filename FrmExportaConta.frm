VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmExportaConta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   7
      Top             =   2160
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
         MouseIcon       =   "FrmExportaConta.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmExportaConta.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   5
      Top             =   3360
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmExportaConta.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmExportaConta.frx":2156
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmExportaConta.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DTPicker1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
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
         Left            =   2040
         Picture         =   "FrmExportaConta.frx":4254
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50593793
         CurrentDate     =   44400
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6165
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
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmExportaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    'ORDENES RAPIDAS
    sBuscar = "SELECT ORDEN_RAPIDA.FECHA, PRODUCTOS_CONSUMIBLES.CUENTA_CONTABLE AS CUENTA, PROVEEDOR_CONSUMO.NOMBRE, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS CARGO, 0 AS ABONO, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_ORDEN_RAPIDA AS DOCUMENTO FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PRODUCTOS_CONSUMIBLES ON ORDEN_RAPIDA_DETALLE.ID_PRODUCTO = PRODUCTOS_CONSUMIBLES.ID_PRODUCTO INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE ORDEN_RAPIDA.FECHA = '" & DTPicker1.Value & "' GROUP BY ORDEN_RAPIDA.FECHA, PRODUCTOS_CONSUMIBLES.CUENTA_CONTABLE, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.ID_ORDEN_RAPIDA"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "OCR - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ORDENES NACIONALES
    sBuscar = "SELECT ORDEN_COMPRA.FECHA, 'XXXXXXXXXX' AS CUENTA, PROVEEDOR.NOMBRE, ORDEN_COMPRA.TOTAL AS CARGO, 0 AS ABONO, ORDEN_COMPRA.Moneda , ORDEN_COMPRA.num_orden AS DOCUMENTO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA.TIPO = 'N') AND ORDEN_COMPRA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "OCN - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ORDENES INTERNACIONALES
    sBuscar = "SELECT ORDEN_COMPRA.FECHA, 'XXXXXXXXXX' AS CUENTA, PROVEEDOR.NOMBRE, ORDEN_COMPRA.TOTAL AS CARGO, 0 AS ABONO, ORDEN_COMPRA.Moneda , ORDEN_COMPRA.num_orden AS DOCUMENTO FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA.TIPO = 'I') AND ORDEN_COMPRA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "OCI - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'VENTAS DE CONTADO FACTURADAS
    sBuscar = "SELECT FECHA, 'XXXXXXXXX' AS CUENTA, NOMBRE, 0 AS CARGO, TOTAL AS ABONO, 'MXN' AS MONEDA, FOLIO AS DOCUMENTO From Ventas WHERE (FACTURADO = '1') AND (UNA_EXIBICION = 'S') FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "V.CONTADO - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ABONOS A VENTAS DE CREDITO FACTURADAS
    sBuscar = "SELECT ABONOS_CUENTA.FECHA, 'XXXXXXXXXXXXX' AS CUENTA, VENTAS.NOMBRE, 0 AS CARGO, ABONOS_CUENTA.CANT_ABONO AS ABONO, 'MXN' AS MONEDA, VENTAS.FOLIO AS DOCUMENTO FROM ABONOS_CUENTA INNER JOIN CUENTA_VENTA ON ABONOS_CUENTA.ID_CUENTA = CUENTA_VENTA.ID_CUENTA INNER JOIN VENTAS ON CUENTA_VENTA.ID_VENTA = VENTAS.ID_VENTA WHERE ABONOS_CUENTA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "V.CREDITO - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ABONOS A COMPRAS NACIONALES
    sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, 'XXXXXXXXXX' AS CUENTA, PROVEEDOR.NOMBRE, 0 AS CARGO, ABONOS_PAGO_OC.CANTIDAD AS ABONO, 'MXN' AS MONEDA, ORDEN_COMPRA.NUM_ORDEN AS DOCUMENTO FROM ABONOS_PAGO_OC INNER JOIN ORDEN_COMPRA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_COMPRA.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ORDEN_COMPRA.TIPO = 'N') AND ABONOS_PAGO_OC = 'N' AND ABONOS_CUENTA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "A.CREDITO - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ABONOS A COMPRAS INTERNACIONALES
    sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, 'XXXXXXXXXX' AS CUENTA, PROVEEDOR.NOMBRE, 0 AS CARGO, ABONOS_PAGO_OC.CANTIDAD AS ABONO, 'MXN' AS MONEDA, ORDEN_COMPRA.NUM_ORDEN AS DOCUMENTO FROM ABONOS_PAGO_OC INNER JOIN ORDEN_COMPRA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_COMPRA.ID_ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR WHERE (ABONOS_PAGO_OC.TIPO = 'I') AND ABONOS_CUENTA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "A.CREDITO - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
    'ABONOS A COMPRAS RAPIDAS
    sBuscar = "SELECT ABONOS_PAGO_OC.FECHA, 'XXXXXXXXXX' AS CUENTA, PROVEEDOR_CONSUMO.NOMBRE, 0 AS CARGO, ABONOS_PAGO_OC.CANTIDAD AS ABONO, 'MXN' AS MONEDA, ORDEN_RAPIDA.ID_ORDEN_RAPIDA AS DOCUMENTO FROM ABONOS_PAGO_OC INNER JOIN ORDEN_RAPIDA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ABONOS_PAGO_OC.TIPO = 'R') AND ABONOS_CUENTA.FECHA = '" & DTPicker1.Value & "'"
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Trim(tRs.Fields("FECHA")))
            If Not IsNull(tRs.Fields("CUENTA")) Then tLi.SubItems(1) = tRs.Fields("CUENTA")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("CARGO")) Then tLi.SubItems(3) = tRs.Fields("CARGO")
            If Not IsNull(tRs.Fields("ABONO")) Then tLi.SubItems(4) = tRs.Fields("ABONO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(5) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("DOCUMENTO")) Then tLi.SubItems(6) = "A.CREDITO - " & tRs.Fields("DOCUMENTO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Cuenta", 1000
        .ColumnHeaders.Add , , "Descripcion", 4500
        .ColumnHeaders.Add , , "Cargo", 1000
        .ColumnHeaders.Add , , "Abono", 1000
        .ColumnHeaders.Add , , "Moneda", 1000
        .ColumnHeaders.Add , , "Documento", 4000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
' Necesario para el correcto funcionmiento agregar al form lo siguiente :
' - Funcion ShellExecute (para abrir el archivo al terminar de ejecutar)
' - CommonDialog
' - ProgressBar
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
Private Sub Image9_Click()
    Unload Me
End Sub
