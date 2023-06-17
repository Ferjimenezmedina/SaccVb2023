VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCorteSemana 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corte por Periodo"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   19
      Top             =   3120
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmCorteSemana.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteSemana.frx":030A
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   10
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCorteSemana.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteSemana.frx":0BA3
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   8
      Top             =   4320
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCorteSemana.frx":2C85
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteSemana.frx":2F8F
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Seleccion del Reporte"
      TabPicture(0)   =   "FrmCorteSemana.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblMensaje"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ProgressBar1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox Check2 
         Caption         =   "VENTAS DE CONTADO"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "VENTAS DE CREDITO"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7435
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
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Reporte"
         Height          =   1575
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   8295
         Begin VB.CommandButton Command1 
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
            Left            =   6840
            Picture         =   "FrmCorteSemana.frx":4AED
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   960
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   6720
            TabIndex        =   1
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39171
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4800
            TabIndex        =   3
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50331649
            CurrentDate     =   39171
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1080
            TabIndex        =   0
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Al :"
            Height          =   255
            Left            =   6480
            TabIndex        =   15
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Del :"
            Height          =   255
            Left            =   4440
            TabIndex        =   14
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal :"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Label LblMensaje 
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Reporte de Ventas :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9600
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "FrmCorteSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim Sucursal As String
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRsREC As ADODB.Recordset
    Dim tRsREM As ADODB.Recordset
    Dim tRsINT As ADODB.Recordset
    Dim tRsCOM As ADODB.Recordset
    Dim tRsCOMAP As ADODB.Recordset
    Dim tRsTOT As ADODB.Recordset
    Dim tLi As ListItem
    Dim Acum As Double
    If Combo1.Text = "<TODAS>" Then
        sBuscar = "SELECT COUNT(ID_VENTA) AS CONTA FROM VENTAS WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997' AND FOLIO <> 'CANCELADO' AND SUBTOTAL <> 0"
    Else
        sBuscar = "SELECT COUNT(ID_VENTA) AS CONTA FROM VENTAS WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997' AND SUCURSAL = '" & Combo1.Text & "' AND FOLIO <> 'CANCELADO' AND SUBTOTAL <> 0"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If tRs.Fields("CONTA") > 0 Then
        LblMensaje.Caption = "PROCESANDO..."
        If Combo1.Text = "<TODAS>" Then
            sBuscar = "SELECT VENTAS.FECHA, VENTAS.ID_VENTA, VENTAS.NOMBRE as NOMBREc, VENTAS.SUBTOTAL, VENTAS.UNA_EXIBICION, VENTAS.FOLIO, USUARIOS.NOMBRE " & _
                      "FROM VENTAS, USUARIOS WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997' AND FOLIO <> 'CANCELADO' AND SUBTOTAL <> 0  AND VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO ORDER BY FECHA"
            Set tRs = cnn.Execute(sBuscar)
        Else
            sBuscar = "SELECT VENTAS.FECHA, VENTAS.ID_VENTA, VENTAS.NOMBRE as NOMBREc, VENTAS.SUBTOTAL, VENTAS.UNA_EXIBICION, VENTAS.FOLIO, USUARIOS.NOMBRE " & _
                      "FROM VENTAS, USUARIOS WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997' AND SUCURSAL = '" & Combo1.Text & "' AND FOLIO <> 'CANCELADO' AND SUBTOTAL <> 0  AND VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO ORDER BY FECHA"
            Set tRs = cnn.Execute(sBuscar)

        End If
        Sucursal = Combo1.Text
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Acum = 0
                If tRs.Fields("UNA_EXIBICION") = "N" Then
                    If Check1.Value = 1 Then
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("FECHA") & "")
                        If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(1) = tRs.Fields("ID_VENTA")
                        If Not IsNull(tRs.Fields("NOMBREc")) Then tLi.SubItems(2) = tRs.Fields("NOMBREc")
                        If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(3) = tRs.Fields("SUBTOTAL") & ""
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%REC'"
                        Set tRsREC = cnn.Execute(sBuscar)
                        If Not IsNull(tRsREC.Fields("TOT")) Then
                            tLi.SubItems(4) = tRsREC.Fields("TOT")
                            Acum = Acum + Val(tRsREC.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%REM'"
                        Set tRsREM = cnn.Execute(sBuscar)
                        If Not IsNull(tRsREM.Fields("TOT")) Then
                            tLi.SubItems(5) = tRsREM.Fields("TOT")
                            Acum = Acum + Val(tRsREM.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%CAMAP'"
                        Set tRsINT = cnn.Execute(sBuscar)
                        If Not IsNull(tRsINT.Fields("TOT")) Then
                            tLi.SubItems(6) = tRsINT.Fields("TOT")
                            Acum = Acum + Val(tRsINT.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%COMGEN'"
                        Set tRsCOM = cnn.Execute(sBuscar)
                        If Not IsNull(tRsCOM.Fields("TOT")) Then
                            tLi.SubItems(7) = tRsCOM.Fields("TOT")
                            Acum = Acum + Val(tRsCOM.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%COMAP'"
                        Set tRsCOMAP = cnn.Execute(sBuscar)
                        If Not IsNull(tRsCOMAP.Fields("TOT")) Then
                            tLi.SubItems(8) = tRsCOMAP.Fields("TOT")
                            Acum = Acum + Val(tRsCOMAP.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA")
                        Set tRsTOT = cnn.Execute(sBuscar)
                        If Not IsNull(tRsTOT.Fields("TOT")) Then
                            If Val(tRsTOT.Fields("TOT")) - Acum <> 0 Then
                                tLi.SubItems(9) = Val(tRsTOT.Fields("TOT")) - Acum
                            End If
                        End If
                        If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(10) = tRs.Fields("FOLIO")
                        If tRs.Fields("UNA_EXIBICION") = "S" Then
                            tLi.SubItems(11) = "CONTADO"
                        Else
                            tLi.SubItems(11) = "CREDITO"
                            tLi.ForeColor = &HFF0000
                            tLi.ListSubItems.Item(1).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(2).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(3).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(4).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(5).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(6).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(7).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(8).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(9).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(10).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(11).ForeColor = &HFF0000
                        End If
                        If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(12) = tRs.Fields("NOMBRE")
                    End If
                Else
                    If Check2.Value = 1 Then
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("FECHA") & "")
                        If Not IsNull(tRs.Fields("ID_VENTA")) Then tLi.SubItems(1) = tRs.Fields("ID_VENTA")
                        If Not IsNull(tRs.Fields("NOMBREc")) Then tLi.SubItems(2) = tRs.Fields("NOMBREc")
                        If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(3) = tRs.Fields("SUBTOTAL") & ""
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%REC'"
                        Set tRsREC = cnn.Execute(sBuscar)
                        If Not IsNull(tRsREC.Fields("TOT")) Then
                            tLi.SubItems(4) = tRsREC.Fields("TOT")
                            Acum = Acum + Val(tRsREC.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%REM'"
                        Set tRsREM = cnn.Execute(sBuscar)
                        If Not IsNull(tRsREM.Fields("TOT")) Then
                            tLi.SubItems(5) = tRsREM.Fields("TOT")
                            Acum = Acum + Val(tRsREM.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%CAMAP'"
                        Set tRsINT = cnn.Execute(sBuscar)
                        If Not IsNull(tRsINT.Fields("TOT")) Then
                            tLi.SubItems(6) = tRsINT.Fields("TOT")
                            Acum = Acum + Val(tRsINT.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%COMGEN'"
                        Set tRsCOM = cnn.Execute(sBuscar)
                        If Not IsNull(tRsCOM.Fields("TOT")) Then
                            tLi.SubItems(7) = tRsCOM.Fields("TOT")
                            Acum = Acum + Val(tRsCOM.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO LIKE '%COMAP'"
                        Set tRsCOMAP = cnn.Execute(sBuscar)
                        If Not IsNull(tRsCOMAP.Fields("TOT")) Then
                            tLi.SubItems(8) = tRsCOMAP.Fields("TOT")
                            Acum = Acum + Val(tRsCOMAP.Fields("TOT"))
                        End If
                        sBuscar = "SELECT SUM(PRECIO_VENTA) AS TOT FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA")
                        Set tRsTOT = cnn.Execute(sBuscar)
                        If Not IsNull(tRsTOT.Fields("TOT")) Then
                            If Val(tRsTOT.Fields("TOT")) - Acum <> 0 Then
                                tLi.SubItems(9) = Val(tRsTOT.Fields("TOT")) - Acum
                            End If
                        End If
                        If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(10) = tRs.Fields("FOLIO")
                        If tRs.Fields("UNA_EXIBICION") = "S" Then
                            tLi.SubItems(11) = "CONTADO"
                        Else
                            tLi.SubItems(11) = "CREDITO"
                            tLi.ForeColor = &HFF0000
                            tLi.ListSubItems.Item(1).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(2).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(3).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(4).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(5).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(6).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(7).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(8).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(9).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(10).ForeColor = &HFF0000
                            tLi.ListSubItems.Item(11).ForeColor = &HFF0000
                        End If
                        If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(12) = tRs.Fields("NOMBRE")
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        LblMensaje.Caption = ""
    Else
        MsgBox "LA CONSULTA NO TRAJO RESULTADOS!", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 7, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "FECHA", 1000
        .ColumnHeaders.Add , , "FOLIO", 1000
        .ColumnHeaders.Add , , "CLIENTE", 5200
        .ColumnHeaders.Add , , "SUBTOTAL", 1200
        .ColumnHeaders.Add , , "RECARGAS", 1200
        .ColumnHeaders.Add , , "REMANUFACTURAS", 1200
        .ColumnHeaders.Add , , "CAMBIO", 1200
        .ColumnHeaders.Add , , "COMPATIBLES", 1200
        .ColumnHeaders.Add , , "COMPATIBLES AP TONER", 1200
        .ColumnHeaders.Add , , "ORIGINALES", 1200
        .ColumnHeaders.Add , , "FACTURA", 1200
        .ColumnHeaders.Add , , "TIPO", 1200
        .ColumnHeaders.Add , , "AGENTE", 1200
    End With
End Sub
Private Sub Image10_Click()
On Error GoTo ManejaError
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
Exit Sub
ManejaError:
    If Err.Number <> 1004 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim Suma As String
    Dim Total As Double
    ConPag = 1
    Total = "0"
    Suma = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Not (ListView1.ListItems.Count = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\VentasPeriodo.pdf") Then
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
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS POR PERIODO DE SUCURSAL " & Sucursal, "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "Venta", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 340, "Cliente", "F2", 8, hCenter
        oDoc.WTextBox Posi, 405, 20, 55, "Folio", "F2", 8, hCenter
        oDoc.WTextBox Posi, 460, 20, 55, "Subtotal", "F2", 8, hCenter
        oDoc.WTextBox Posi, 515, 20, 55, "Tipo", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        For Cont = 1 To ListView1.ListItems.Count
            oDoc.WTextBox Posi, 10, 20, 60, ListView1.ListItems(Cont).SubItems(1), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 340, ListView1.ListItems(Cont).SubItems(2), "F3", 7, hLeft
            oDoc.WTextBox Posi, 405, 20, 55, ListView1.ListItems(Cont).SubItems(10), "F3", 7, hLeft
            oDoc.WTextBox Posi, 460, 20, 50, Format(ListView1.ListItems(Cont).SubItems(3), "###,###,##0.00"), "F3", 7, hRight
            Total = Total + ListView1.ListItems(Cont).SubItems(3)
            oDoc.WTextBox Posi, 515, 20, 55, ListView1.ListItems(Cont).SubItems(11), "F3", 7, hLeft
            Posi = Posi + 12
            If Posi >= 700 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs1 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS POR PERIODO DE SUCURSAL " & Sucursal, "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "Venta", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 340, "Cliente", "F2", 8, hCenter
                oDoc.WTextBox Posi, 405, 20, 55, "Folio", "F2", 8, hCenter
                oDoc.WTextBox Posi, 460, 20, 55, "Subtotal", "F2", 8, hCenter
                oDoc.WTextBox Posi, 515, 20, 55, "Tipo", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
        Next
        ' Linea
        Posi = Posi + 15
        oDoc.WTextBox Posi, 400, 20, 120, Format(Total, "###,###,##0.00"), "F3", 10, hRight
        Posi = Posi + 26
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 16
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
Private Sub Image9_Click()
    Unload Me
End Sub
