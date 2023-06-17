VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepCXC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de CXC (Detalle)"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepCXC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdBuscar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CheckBox Check1 
         Caption         =   "Solo Facturadas"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de fechas"
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   40014
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   40014
         End
         Begin VB.Label Label6 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Al :"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
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
         Left            =   4080
         Picture         =   "FrmRepCXC.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8493
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
      Left            =   10200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   2
      Top             =   4080
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepCXC.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCXC.frx":2CF8
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   0
      Top             =   5280
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepCXC.frx":483A
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCXC.frx":4B44
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmRepCXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT VENTAS.NOMBRE, VENTAS.FOLIO, VENTAS.ID_VENTA, VENTAS.FECHA, VENTAS.TOTAL, DATEDIFF(day, VENTAS.FECHA, GETDATE()) AS DIAS, ISNULL(SUM(ABONOS_CUENTA.CANT_ABONO), 0) AS ABONOS FROM VENTAS INNER JOIN CUENTA_VENTA ON VENTAS.ID_VENTA = CUENTA_VENTA.ID_VENTA INNER JOIN CUENTAS ON CUENTA_VENTA.ID_CUENTA = CUENTAS.ID_CUENTA LEFT OUTER JOIN ABONOS_CUENTA ON CUENTA_VENTA.ID_CUENTA = ABONOS_CUENTA.ID_CUENTA WHERE (CUENTAS.PAGADA = 'N') AND VENTAS.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " '"
    If Check1.Value = 1 Then
        sBuscar = sBuscar & " AND VENTAS.FACTURADO = 1"
    End If
    sBuscar = sBuscar & " GROUP BY VENTAS.NOMBRE, VENTAS.FOLIO, VENTAS.ID_VENTA, VENTAS.FECHA, VENTAS.TOTAL"
    Set tRs = cnn.Execute(sBuscar)
    StrRep = sBuscar
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NOMBRE"))
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(2) = tRs.Fields("ID_VENTA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            If tRs.Fields("DIAS") <= 15 Then
                tLi.SubItems(5) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            End If
            If tRs.Fields("DIAS") > 15 And tRs.Fields("DIAS") <= 30 Then
                tLi.SubItems(6) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            End If
            If tRs.Fields("DIAS") > 30 And tRs.Fields("DIAS") <= 45 Then
                tLi.SubItems(7) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            End If
            If tRs.Fields("DIAS") > 45 And tRs.Fields("DIAS") <= 60 Then
                tLi.SubItems(8) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            End If
            If tRs.Fields("DIAS") > 60 Then
                tLi.SubItems(9) = tRs.Fields("TOTAL") - tRs.Fields("ABONOS")
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker3.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker4.Value = Format(Date, "dd/mm/yyyy")
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "NOMBRE", 4500
        .ColumnHeaders.Add , , "FOLIO", 1200
        .ColumnHeaders.Add , , "VENTA", 1200
        .ColumnHeaders.Add , , "FECHA", 1300
        .ColumnHeaders.Add , , "TOTAL", 1300
        .ColumnHeaders.Add , , "15 DIAS", 1300
        .ColumnHeaders.Add , , "16 - 30 DIAS", 1300
        .ColumnHeaders.Add , , "31 - 45 DIAS", 1300
        .ColumnHeaders.Add , , "46 - 60 DIAS", 1300
        .ColumnHeaders.Add , , "60 + DIAS", 1300
    End With
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
