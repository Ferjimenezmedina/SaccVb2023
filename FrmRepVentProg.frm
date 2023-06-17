VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepVentProg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ventas programadas"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepVentProg.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Opciones"
         Height          =   2295
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton Option2 
            Caption         =   "Por fecha de entrega"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por fecha de captura"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de fechas"
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   0
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   40403
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   600
            TabIndex        =   1
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   40403
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
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
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   6
      Top             =   480
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepVentProg.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentProg.frx":0326
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepVentProg.frx":1E68
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentProg.frx":2172
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
End
Attribute VB_Name = "FrmRepVentProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Pedido", 1800
        .ColumnHeaders.Add , , "No. Cliente", 1800
        .ColumnHeaders.Add , , "Cliente", 7450
        .ColumnHeaders.Add , , "Producto", 2450
        .ColumnHeaders.Add , , "Cantidad", 500
        .ColumnHeaders.Add , , "Marca", 1500
        .ColumnHeaders.Add , , "Fecha de Captura", 1500
        .ColumnHeaders.Add , , "Fecha de Entrega", 1500
    End With
    DTPicker1.Value = Date - 7
    DTPicker2.Value = Date
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
        "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, PED_CLIEN.ID_CLIENTE, CLIENTE.NOMBRE, PED_CLIEN_DETALLE.ID_PRODUCTO, PED_CLIEN_DETALLE.CANTIDAD_PEDIDA , ALMACEN3.Marca, PED_CLIEN.fecha, PED_CLIEN.FECHA_CAPTURA FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN ALMACEN3 ON PED_CLIEN_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO"
    If Option1.Value Then
        sBuscar = sBuscar & " WHERE PED_CLIEN.FECHA_CAPTURA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    Else
        sBuscar = sBuscar & " WHERE PED_CLIEN.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NO_PEDIDO"))
            If Not IsNull(tRs.Fields("ID_CLIENTE")) Then tLi.SubItems(1) = tRs.Fields("ID_CLIENTE")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD_PEDIDA")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD_PEDIDA")
            If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(5) = tRs.Fields("MARCA")
            If Not IsNull(tRs.Fields("FECHA_CAPTURA")) Then tLi.SubItems(6) = tRs.Fields("FECHA_CAPTURA")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(7) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
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
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
