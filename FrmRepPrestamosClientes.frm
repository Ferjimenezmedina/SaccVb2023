VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepPrestamosClientes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Prestamos a Clientes"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   17
      Top             =   3240
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepPrestamosClientes.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepPrestamosClientes.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepPrestamosClientes.frx":1E4C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPicker1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
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
         Left            =   7800
         Picture         =   "FrmRepPrestamosClientes.frx":1E68
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "N. Comercial"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   975
         Left            =   6360
         TabIndex        =   10
         Top             =   120
         Width           =   1335
         Begin VB.OptionButton Option3 
            Caption         =   "Todo"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cerrados"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Abiertos"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   44712
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   44712
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   0
      Top             =   4440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepPrestamosClientes.frx":483A
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepPrestamosClientes.frx":4B44
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmRepPrestamosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command2_Click()
    BuscaPrestamo
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "Folio Prestamo", 1500
        .ColumnHeaders.Add , , "Cliente", 5500
        .ColumnHeaders.Add , , "Fecha de Prestamo", 1500
        .ColumnHeaders.Add , , "Fecha de Entrega", 1500
        .ColumnHeaders.Add , , "Deposito", 1500
        .ColumnHeaders.Add , , "Notas", 7500
        .ColumnHeaders.Add , , "Estado", 1500
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Serie", 3500
    End With
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
    If ListView1.ListItems.COUNT > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.COUNT
            For Con = 1 To ListView1.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.COUNT
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
Private Sub BuscaPrestamo()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT PRESTAMOS_CLIENTES.ID_CLIENTE, PRESTAMOS_CLIENTES.ID_PRESTAMO, CLIENTE.NOMBRE, PRESTAMOS_CLIENTES.FECHA_PRESTAMO, PRESTAMOS_CLIENTES.FECHA_ENTREGA, PRESTAMOS_CLIENTES.DEPOSITO, PRESTAMOS_CLIENTES.NOTAS, PRESTAMOS_CLIENTES.ESTADO, PRESTAMOS_CLIENTES_DETALLE.ID_PRODUCTO, PRESTAMOS_CLIENTES_DETALLE.CANTIDAD , PRESTAMOS_CLIENTES_DETALLE.NO_SERIE FROM PRESTAMOS_CLIENTES INNER JOIN CLIENTE ON PRESTAMOS_CLIENTES.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN PRESTAMOS_CLIENTES_DETALLE ON PRESTAMOS_CLIENTES.ID_PRESTAMO = PRESTAMOS_CLIENTES_DETALLE.ID_PRESTAMO WHERE PRESTAMOS_CLIENTES.FECHA_ENTREGA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & "'"
    If Option1.value Then
        sBuscar = sBuscar & " AND PRESTAMOS_CLIENTES.ESTADO = 'P'"
    Else
        If Option2.value Then
            sBuscar = sBuscar & " AND PRESTAMOS_CLIENTES.ESTADO = 'C'"
        End If
    End If
    If Text1.Text <> "" Then
        If Option4.value Then
            sBuscar = sBuscar & " AND CLIENTE.NOMBRE LIKE '%" & Text1.Text & "%'"
        Else
            sBuscar = sBuscar & " AND CLIENTE.NOMBRE_COMERCIAL LIKE '%" & Text1.Text & "%'"
        End If
    End If
    sBuscar = sBuscar & " ORDER BY ID_PRESTAMO"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRESTAMO"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA_PRESTAMO")) Then tLi.SubItems(2) = tRs.Fields("FECHA_PRESTAMO")
            If Not IsNull(tRs.Fields("FECHA_ENTREGA")) Then tLi.SubItems(3) = tRs.Fields("FECHA_ENTREGA")
            If Not IsNull(tRs.Fields("DEPOSITO")) Then tLi.SubItems(4) = tRs.Fields("DEPOSITO")
            If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(5) = tRs.Fields("NOTAS")
            If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(6) = tRs.Fields("ESTADO")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(7) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(8) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("NO_SERIE")) Then tLi.SubItems(9) = tRs.Fields("NO_SERIE")
            tRs.MoveNext
        Loop
    End If
End Sub
