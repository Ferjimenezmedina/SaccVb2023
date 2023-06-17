VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepAbonos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Abonos a Ordenes"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepAbonos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton Command1 
         Caption         =   "Busca"
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
         Left            =   6480
         Picture         =   "FrmRepAbonos.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de fechas"
         Height          =   855
         Left            =   2640
         TabIndex        =   19
         Top             =   720
         Width           =   3615
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   600
            TabIndex        =   3
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50987009
            CurrentDate     =   44389
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1440
            TabIndex        =   2
            Top             =   0
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2280
            TabIndex        =   4
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50987009
            CurrentDate     =   44389
         End
         Begin VB.Label Label4 
            Caption         =   "Al :"
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Rapida"
         Height          =   255
         Left            =   6360
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Internacional"
         Height          =   255
         Left            =   6360
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   6360
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6800
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
         Caption         =   "No. de Orden"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   14
      Top             =   2160
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepAbonos.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepAbonos.frx":2CF8
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   12
      Top             =   3360
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepAbonos.frx":3287
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepAbonos.frx":3591
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   10
      Top             =   4560
      Width           =   975
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepAbonos.frx":50D3
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepAbonos.frx":53DD
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   8280
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmRepAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    If Option1.Value = True Then
        BuscaOrdenNacional
    End If
    If Option2.Value = True Then
        BuscaOrdenInternacional
    End If
    If Option3.Value = True Then
        BuscaOrdenRapida
    End If
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.DTPicker1 = Format(Date - 15, "dd/mm/yyyy")
    Me.DTPicker2 = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "Id Abono", 1500
        .ColumnHeaders.Add , , "No. Orden", 1500
        .ColumnHeaders.Add , , "Tipo", 1500
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Fecha Abono", 2500
        .ColumnHeaders.Add , , "Fecha Orden", 2500
        .ColumnHeaders.Add , , "Banco", 3500
        .ColumnHeaders.Add , , "No. Transferencia", 2500
        .ColumnHeaders.Add , , "No. Cheque", 2500
        .ColumnHeaders.Add , , "Abono", 2500
        .ColumnHeaders.Add , , "Total Orden", 2500
    End With
End Sub
Sub BuscaOrdenRapida()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sWhere As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ABONOS_PAGO_OC.ID_ABONO, ABONOS_PAGO_OC.NUM_ORDEN, ABONOS_PAGO_OC.TIPO, PROVEEDOR_CONSUMO.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_RAPIDA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, SUM(ORDEN_RAPIDA_DETALLE.Total) As Total FROM ABONOS_PAGO_OC INNER JOIN ORDEN_RAPIDA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE ABONOS_PAGO_OC.TIPO = 'R' "
    If Check1.Value = 1 Then
        sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        sWhere = sWhere & " AND PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
    End If
    sBuscar = sBuscar & sWhere & " GROUP BY ABONOS_PAGO_OC.ID_ABONO, ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ABONOS_PAGO_OC.TIPO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, ABONOS_PAGO_OC.NUM_ORDEN ORDER BY ABONOS_PAGO_OC.ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        Me.ListView1.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_ABONO"))
                If Not IsNull(.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                If Not IsNull(.Fields("TIPO")) Then tLi.SubItems(2) = Trim(.Fields("TIPO"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("FECHA_ABONO")) Then tLi.SubItems(4) = Trim(.Fields("FECHA_ABONO"))
                If Not IsNull(.Fields("FECHA_ORDEN")) Then tLi.SubItems(5) = Trim(.Fields("FECHA_ORDEN"))
                If Not IsNull(.Fields("BANCO")) Then tLi.SubItems(6) = Trim(.Fields("BANCO"))
                If Not IsNull(.Fields("NUMTRANS")) Then tLi.SubItems(7) = Trim(.Fields("NUMTRANS"))
                If Not IsNull(.Fields("NUMCHEQUE")) Then tLi.SubItems(8) = Trim(.Fields("NUMCHEQUE"))
                If Not IsNull(.Fields("CANT_ABONO")) Then tLi.SubItems(9) = Trim(.Fields("CANT_ABONO"))
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(10) = Trim(.Fields("TOTAL"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub BuscaOrdenNacional()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sWhere As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ABONOS_PAGO_OC.ID_ABONO, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_COMPRA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE ABONOS_PAGO_OC.TIPO = 'N' "
    If Check1.Value = 1 Then
        sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        sWhere = sWhere & " AND PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
    End If
    sBuscar = sBuscar & sWhere & " ORDER BY ABONOS_PAGO_OC.ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        Me.ListView1.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_ABONO"))
                If Not IsNull(.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                If Not IsNull(.Fields("TIPO")) Then tLi.SubItems(2) = Trim(.Fields("TIPO"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("FECHA_ABONO")) Then tLi.SubItems(4) = Trim(.Fields("FECHA_ABONO"))
                If Not IsNull(.Fields("FECHA_ORDEN")) Then tLi.SubItems(5) = Trim(.Fields("FECHA_ORDEN"))
                If Not IsNull(.Fields("BANCO")) Then tLi.SubItems(6) = Trim(.Fields("BANCO"))
                If Not IsNull(.Fields("NUMTRANS")) Then tLi.SubItems(7) = Trim(.Fields("NUMTRANS"))
                If Not IsNull(.Fields("NUMCHEQUE")) Then tLi.SubItems(8) = Trim(.Fields("NUMCHEQUE"))
                If Not IsNull(.Fields("CANT_ABONO")) Then tLi.SubItems(9) = Trim(.Fields("CANT_ABONO"))
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(10) = Trim(.Fields("TOTAL"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub BuscaOrdenInternacional()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sWhere As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ABONOS_PAGO_OC.ID_ABONO, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_COMPRA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE ABONOS_PAGO_OC.TIPO = 'N' "
    If Check1.Value = 1 Then
        sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        sWhere = sWhere & " AND PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
    End If
    sBuscar = sBuscar & sWhere & " ORDER BY ABONOS_PAGO_OC.ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        Me.ListView1.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_ABONO"))
                If Not IsNull(.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                If Not IsNull(.Fields("TIPO")) Then tLi.SubItems(2) = Trim(.Fields("TIPO"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("FECHA_ABONO")) Then tLi.SubItems(4) = Trim(.Fields("FECHA_ABONO"))
                If Not IsNull(.Fields("FECHA_ORDEN")) Then tLi.SubItems(5) = Trim(.Fields("FECHA_ORDEN"))
                If Not IsNull(.Fields("BANCO")) Then tLi.SubItems(6) = Trim(.Fields("BANCO"))
                If Not IsNull(.Fields("NUMTRANS")) Then tLi.SubItems(7) = Trim(.Fields("NUMTRANS"))
                If Not IsNull(.Fields("NUMCHEQUE")) Then tLi.SubItems(8) = Trim(.Fields("NUMCHEQUE"))
                If Not IsNull(.Fields("CANT_ABONO")) Then tLi.SubItems(9) = Trim(.Fields("CANT_ABONO"))
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(10) = Trim(.Fields("TOTAL"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image26_Click()
    Reporte
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Reporte()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim sBuscar As String
    Dim sBuscarC As String
    Dim ConPag As Integer
    Dim Suma As String
    Dim Total As String
    Dim sWhere As String
    Dim sBuscaC As String
    ConPag = 1
    Total = "0"
    Suma = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Option1.Value Then
        sBuscar = "SELECT PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'N')"
        sBuscarC = "SELECT ABONOS_PAGO_OC.ID_ABONO, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_COMPRA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'N')"
        If Check1.Value = 1 Then
            sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
        End If
        If Text1.Text <> "" Then
            sWhere = sWhere & " AND PROVEEDOR.NOMBRE LIKE '%" & Text1.Text & "%'"
        End If
        If Text2.Text <> "" Then
            sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
        End If
        sBuscar = sBuscar & sWhere & " GROUP BY PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE ORDER BY dbo.PROVEEDOR.NOMBRE"
    End If
    If Option2.Value Then
        sBuscar = "SELECT PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'I')"
        sBuscarC = "SELECT ABONOS_PAGO_OC.ID_ABONO, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.TIPO, PROVEEDOR.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_COMPRA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ORDEN_COMPRA.Total FROM ORDEN_COMPRA INNER JOIN PROVEEDOR ON ORDEN_COMPRA.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ABONOS_PAGO_OC ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ABONOS_PAGO_OC.ID_ORDEN WHERE (ORDEN_COMPRA.TIPO = 'I')"
        If Check1.Value = 1 Then
            sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
        End If
        If Text1.Text <> "" Then
            sWhere = sWhere & " AND PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
        End If
        If Text2.Text <> "" Then
            sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
        End If
        sBuscar = sBuscar & sWhere & " GROUP BY PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE ORDER BY dbo.PROVEEDOR.NOMBRE"
    End If
    If Option3.Value Then
        sBuscar = "SELECT PROVEEDOR_CONSUMO.ID_PROVEEDOR, PROVEEDOR_CONSUMO.NOMBRE FROM ABONOS_PAGO_OC INNER JOIN ORDEN_RAPIDA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ABONOS_PAGO_OC.TIPO = 'R')"
        sBuscarC = "SELECT ABONOS_PAGO_OC.ID_ABONO, ABONOS_PAGO_OC.NUM_ORDEN, ABONOS_PAGO_OC.TIPO, PROVEEDOR_CONSUMO.NOMBRE, ABONOS_PAGO_OC.FECHA AS FECHA_ABONO, ORDEN_RAPIDA.FECHA AS FECHA_ORDEN, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, SUM(ORDEN_RAPIDA_DETALLE.Total) As Total FROM ABONOS_PAGO_OC INNER JOIN ORDEN_RAPIDA ON ABONOS_PAGO_OC.ID_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ABONOS_PAGO_OC.TIPO = 'R')"
        If Check1.Value = 1 Then
            sWhere = " AND ABONOS_PAGO_OC.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
        End If
        If Text1.Text <> "" Then
            sWhere = sWhere & " AND PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
        End If
        If Text2.Text <> "" Then
            sWhere = sWhere & " AND ABONOS_PAGO_OC.NUM_ORDEN = " & Text2.Text
        End If
        sBuscar = sBuscar & sWhere & " GROUP BY PROVEEDOR_CONSUMO.ID_PROVEEDOR, PROVEEDOR_CONSUMO.NOMBRE ORDER BY PROVEEDOR_CONSUMO.NOMBRE"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\RepAbonos.pdf") Then
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
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 70, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 205, 100, 175, "ABONOS A ORDEN DE COMPRA", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 50, "NO. ORDEN", "F2", 8, hCenter
        oDoc.WTextBox Posi, 80, 20, 60, "F. ORDEN", "F2", 8, hCenter
        oDoc.WTextBox Posi, 140, 20, 60, "F. ABONO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 200, 20, 100, "BANCO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 300, 20, 60, "NO. TRANS.", "F2", 8, hCenter
        oDoc.WTextBox Posi, 360, 20, 60, "NO. CHEQUE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 420, 20, 80, "CANTIDAD ABONO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 500, 20, 80, "TOTAL ORDEN", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then oDoc.WTextBox Posi, 10, 20, 400, tRs.Fields("NOMBRE"), "F2", 10, hLeft
            Posi = Posi + 12
            If Option3.Value Then
                sBuscar = sBuscarC & sWhere & " AND PROVEEDOR_CONSUMO.ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR") & " GROUP BY ABONOS_PAGO_OC.ID_ABONO, ABONOS_PAGO_OC.FECHA, ABONOS_PAGO_OC.BANCO, ABONOS_PAGO_OC.NUMTRANS, ABONOS_PAGO_OC.NUMCHEQUE, ABONOS_PAGO_OC.CANT_ABONO, ABONOS_PAGO_OC.TIPO, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.FECHA, ABONOS_PAGO_OC.NUM_ORDEN ORDER BY ABONOS_PAGO_OC.NUM_ORDEN "
            Else
                sBuscar = sBuscarC & sWhere & " AND PROVEEDOR.ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR") & " ORDER BY ORDEN_COMPRA.NUM_ORDEN"
            End If
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.EOF And tRs2.BOF) Then
                Do While Not tRs2.EOF
                    If Not IsNull(tRs2.Fields("NUM_ORDEN")) Then oDoc.WTextBox Posi, 30, 20, 50, tRs2.Fields("NUM_ORDEN"), "F3", 8, hLeft
                    If Not IsNull(tRs2.Fields("FECHA_ORDEN")) Then oDoc.WTextBox Posi, 80, 20, 60, tRs2.Fields("FECHA_ORDEN"), "F3", 8, hLeft
                    If Not IsNull(tRs2.Fields("FECHA_ABONO")) Then oDoc.WTextBox Posi, 140, 20, 60, tRs2.Fields("FECHA_ABONO"), "F3", 8, hLeft
                    If Not IsNull(tRs2.Fields("BANCO")) Then oDoc.WTextBox Posi, 200, 20, 100, tRs2.Fields("BANCO"), "F3", 7, hLeft
                    If Not IsNull(tRs2.Fields("NUMTRANS")) Then oDoc.WTextBox Posi, 320, 20, 60, tRs2.Fields("NUMTRANS"), "F3", 8, hLeft
                    If Not IsNull(tRs2.Fields("NUMCHEQUE")) Then oDoc.WTextBox Posi, 380, 20, 60, tRs2.Fields("NUMCHEQUE"), "F3", 8, hLeft
                    If Not IsNull(tRs2.Fields("CANT_ABONO")) Then oDoc.WTextBox Posi, 420, 20, 60, Format(tRs2.Fields("CANT_ABONO"), "###,###,##0.00"), "F3", 8, hRight
                    If Not IsNull(tRs2.Fields("TOTAL")) Then oDoc.WTextBox Posi, 500, 20, 60, Format(tRs2.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
                    Suma = Suma + tRs2.Fields("CANT_ABONO")
                    Posi = Posi + 12
                    If Posi >= 650 Then
                        oDoc.NewPage A4_Vertical
                        oDoc.WImage 70, 40, 43, 161, "Logo"
                        sBuscar = "SELECT * FROM EMPRESA  "
                        Set tRs1 = cnn.Execute(sBuscar)
                        oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                        oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                        oDoc.WTextBox 70, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                        oDoc.WTextBox 90, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                        oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.WTextBox 100, 205, 100, 175, "ABONOS A ORDEN DE COMPRA", "F3", 8, hCenter
                        Posi = 120
                        oDoc.WTextBox Posi, 10, 20, 50, "NO. ORDEN", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 80, 20, 60, "F. ORDEN", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 140, 20, 60, "F. ABONO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 200, 20, 100, "BANCO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 300, 20, 60, "NO. TRANS.", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 360, 20, 60, "NO. CHEQUE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 420, 20, 80, "CANTIDAD ABONO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 500, 20, 80, "TOTAL ORDEN", "F2", 8, hCenter
                        Posi = Posi + 12
                        ' Linea
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 6
                    End If
                    tRs2.MoveNext
                Loop
                ' Linea
                Posi = Posi + 6
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                oDoc.WTextBox Posi, 450, 20, 50, "TOTAL", "F3", 7, hLeft
                If Not IsNull(Suma) Then oDoc.WTextBox Posi, 515, 20, 55, Format(Suma, "$ #,###,##0.00"), "F3", 7, hRight
                Posi = Posi + 16
            End If
            tRs.MoveNext
        Loop
        
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 205, 100, 175, "COMENTARIOS", "F3", 8, hCenter
        Posi = Posi + 20
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
