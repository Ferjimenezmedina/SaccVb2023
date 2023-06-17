VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMaxMin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Máximos y Mínimos Almacén 2"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   6
      Top             =   2160
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmMaxMin.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMin.frx":030A
         Top             =   240
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   4
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmMaxMin.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMin.frx":20C6
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmMaxMin.frx":3C08
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMin.frx":3F12
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Por Pedir"
      TabPicture(0)   =   "FrmMaxMin.frx":5FF4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdTodo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "En proceso de compras"
      TabPicture(1)   =   "FrmMaxMin.frx":6010
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8070
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
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
         Left            =   4320
         Picture         =   "FrmMaxMin.frx":602C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdTodo 
         Caption         =   "+"
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
         Left            =   4680
         Picture         =   "FrmMaxMin.frx":89FE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5280
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8281
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmMaxMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdTodo_Click()
    Dim Cont As Double
    For Cont = 1 To ListView1.ListItems.Count
        ListView1.ListItems(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command1_Click()
    Dim Cont As Double
    For Cont = 1 To ListView1.ListItems.Count
        ListView1.ListItems(Cont).Checked = False
    Next Cont
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
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
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Producto", 2900
        .ColumnHeaders.Add , , "Existencia", 1500
        .ColumnHeaders.Add , , "Pedir", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Producto", 1900
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Estado", 1500
    End With
    Actualiza
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
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    Else
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
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub imgLeer_Click()
    Dim sBuscar As String
    Dim Cont As Integer
    Dim CONT2 As String
    Dim tRs As ADODB.Recordset
    CONT2 = 1
    For Cont = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Cont).Checked Then
            sBuscar = "SELECT DESCRIPCION, MARCA FROM ALMACEN2 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Cont) & "'"
            Set tRs = cnn.Execute(sBuscar)
            sBuscar = "INSERT INTO REQUISICION (ID_PRODUCTO, DESCRIPCION, COMENTARIO, CANTIDAD, FECHA, CONTADOR, URGENTE, MARCA, ACTIVO, FOLIO, ALMACEN) VALUES ('" & ListView1.ListItems(Cont) & "', '" & tRs.Fields("Descripcion") & "', 'PEDIDO POR MAXIMOS Y MINIMOS SACC A " & Date & "', " & ListView1.ListItems(Cont).SubItems(2) & ", GETDATE(), '0', 'S', '" & tRs.Fields("MARCA") & "', '0', '0', 'A2')"
            cnn.Execute (sBuscar)
            CONT2 = CONT2 + 1
        End If
    Next Cont
    Actualiza
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "UPDATE REQUISICION SET CONTADOR = '0' WHERE (CONTADOR = '*')"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT ALMACEN2.ID_PRODUCTO, ALMACEN2.C_MAXIMA, EXISTENCIAS.CANTIDAD FROM ALMACEN2, EXISTENCIAS WHERE EXISTENCIAS.CANTIDAD <= ALMACEN2.C_MINIMA AND EXISTENCIAS.ID_PRODUCTO = ALMACEN2.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = 'BODEGA' AND ALMACEN2.ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM REQUISICION WHERE ACTIVO = 0 AND CONTADOR = 0 AND COTIZADA = 0)  AND ALMACEN2.ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'A') Union SELECT ID_PRODUCTO, C_MAXIMA, (0) AS CANTIDAD FROM ALMACEN2 WHERE ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM EXISTENCIAS WHERE SUCURSAL = 'BODEGA') AND C_MAXIMA > 0 AND ALMACEN2.ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM REQUISICION WHERE ACTIVO = 0 AND CONTADOR = 0 AND COTIZADA = 0)  AND ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'A')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("CANTIDAD")
            tLi.SubItems(2) = tRs.Fields("C_MAXIMA")
            tRs.MoveNext
        Loop
    End If
    ListView2.ListItems.Clear
    'sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, 'REQUISICION' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO IN (SELECT ID_PRODUCTO From REQUISICION WHERE (ACTIVO = 0) AND (CONTADOR = 0) AND (COTIZADA = 0) AND (URGENTE = 'N'))) AND (C_MAXIMA > 0) Union SELECT ID_PRODUCTO,  DESCRIPCION, 'COTIZACION' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO From COTIZA_REQUI WHERE (ESTADO_ACTUAL = 'A'))) AND (C_MAXIMA > 0) Union SELECT ID_PRODUCTO,  DESCRIPCION, 'ORDEN DE COMPRA' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO NOT IN (SELECT ORDEN_COMPRA_DETALLE.ID_PRODUCTO FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA_DETALLE.SURTIDO = 0) AND (ORDEN_COMPRA.CONFIRMADA NOT IN ('E', 'C', 'D')))) AND (C_MAXIMA > 0)"
    sBuscar = "SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'REQUISICION' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO IN (SELECT ID_PRODUCTO From REQUISICION WHERE (ACTIVO = 0) AND (CONTADOR = 0) AND (COTIZADA = 0) AND (URGENTE = 'N'))) AND (C_MAXIMA > 0) Union SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'COTIZACION' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO IN (SELECT ID_PRODUCTO From dbo.COTIZA_REQUI WHERE (ESTADO_ACTUAL = 'A'))) AND (C_MAXIMA > 0) Union SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'ORDEN' AS ESTADO From ALMACEN2 WHERE (ID_PRODUCTO IN (SELECT ORDEN_COMPRA_DETALLE.ID_PRODUCTO FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA_DETALLE.SURTIDO = 0) AND (ORDEN_COMPRA.CONFIRMADA NOT IN ('E', 'C', 'D'))))"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            tLi.SubItems(2) = tRs.Fields("ESTADO")
            tRs.MoveNext
        Loop
    End If
End Sub
