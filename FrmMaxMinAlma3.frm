VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMaxMinAlma3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmMaxMinAlma3.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMinAlma3.frx":030A
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
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   2
      Top             =   3360
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmMaxMinAlma3.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMinAlma3.frx":26F6
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5400
      TabIndex        =   0
      Top             =   2160
      Width           =   975
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmMaxMinAlma3.frx":4238
         MousePointer    =   99  'Custom
         Picture         =   "FrmMaxMinAlma3.frx":4542
         Top             =   240
         Width           =   705
      End
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Faltantes"
      TabPicture(0)   =   "FrmMaxMinAlma3.frx":5FF4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdTodo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "En Compras"
      TabPicture(1)   =   "FrmMaxMinAlma3.frx":6010
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8705
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Picture         =   "FrmMaxMinAlma3.frx":602C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5280
         Width           =   255
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
         Picture         =   "FrmMaxMinAlma3.frx":89FE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5280
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   9
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
Attribute VB_Name = "FrmMaxMinAlma3"
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
        .ColumnHeaders.Add , , "C. Mínima", 1500
        .ColumnHeaders.Add , , "C. Máxima", 1500
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
        .ColumnHeaders.Add , , "Producto", 1500
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
            sBuscar = "SELECT Descripcion, MARCA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Cont) & "'"
            Set tRs = cnn.Execute(sBuscar)
            sBuscar = "INSERT INTO REQUISICION (ID_PRODUCTO, DESCRIPCION, COMENTARIO, CANTIDAD, FECHA, CONTADOR, URGENTE, MARCA, ACTIVO, FOLIO, ALMACEN) VALUES ('" & ListView1.ListItems(Cont) & "', '" & tRs.Fields("Descripcion") & "', 'PEDIDO POR MAXIMOS Y MINIMOS SACC A " & Date & "', " & ListView1.ListItems(Cont).SubItems(4) & ", GETDATE(), '0', 'N', '" & tRs.Fields("MARCA") & "', '0', '0', 'A3')"
            cnn.Execute (sBuscar)
            CONT2 = CONT2 + 1
        End If
    Next Cont
    Actualiza
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim Existencia As Double
    ListView1.ListItems.Clear
    sBuscar = "UPDATE REQUISICION SET CONTADOR = '0' WHERE (CONTADOR = '*')"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT ID_PRODUCTO, C_MINIMA, C_MAXIMA From ALMACEN3 WHERE (ID_PRODUCTO NOT IN(SELECT ID_PRODUCTO From REQUISICION WHERE (ACTIVO = 0) AND (CONTADOR = 0) AND (COTIZADA = 0) AND (URGENTE = 'N'))) AND (C_MAXIMA > 0) AND (ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO From COTIZA_REQUI WHERE (ESTADO_ACTUAL = 'A'))) AND (C_MAXIMA > 0) AND (ID_PRODUCTO NOT IN (SELECT ORDEN_COMPRA_DETALLE.ID_PRODUCTO FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA_DETALLE.SURTIDO = 0) AND (ORDEN_COMPRA.CONFIRMADA NOT IN ('E', 'C', 'D')))) ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            'If tRs.Fields("ID_PRODUCTO") = "HPTCE255ACOMGEN" Then
            '    MsgBox "ya"
            'End If
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                Existencia = CDbl(tRs1.Fields("CANTIDAD"))
            Else
                Existencia = 0
            End If
            If Existencia <= CDbl(tRs.Fields("C_MINIMA")) Then
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("C_MINIMA")
                tLi.SubItems(2) = tRs.Fields("C_MAXIMA")
                tLi.SubItems(3) = Existencia
                tLi.SubItems(4) = CDbl(tRs.Fields("C_MAXIMA")) - CDbl(Existencia)
            End If
            tRs.MoveNext
        Loop
    End If
    ListView2.ListItems.Clear
    sBuscar = "SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'REQUISICION' AS ESTADO From ALMACEN3 WHERE (ID_PRODUCTO IN (SELECT ID_PRODUCTO From REQUISICION WHERE (ACTIVO = 0) AND (CONTADOR = 0) AND (COTIZADA = 0) AND (URGENTE = 'N'))) AND (C_MAXIMA > 0) Union SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'COTIZACION' AS ESTADO From ALMACEN3 WHERE (ID_PRODUCTO IN (SELECT ID_PRODUCTO From dbo.COTIZA_REQUI WHERE (ESTADO_ACTUAL = 'A'))) AND (C_MAXIMA > 0) Union SELECT TOP (100) PERCENT ID_PRODUCTO, DESCRIPCION, 'ORDEN' AS ESTADO From ALMACEN3 WHERE (ID_PRODUCTO IN (SELECT ORDEN_COMPRA_DETALLE.ID_PRODUCTO FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA_DETALLE.SURTIDO = 0) AND (ORDEN_COMPRA.CONFIRMADA NOT IN ('E', 'C', 'D'))))"
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


