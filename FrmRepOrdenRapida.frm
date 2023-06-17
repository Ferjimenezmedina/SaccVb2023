VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepOrdenRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ordenes Rapidas"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
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
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   17
      Top             =   120
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
         MouseIcon       =   "FrmRepOrdenRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenRapida.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   4
      Top             =   1320
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepOrdenRapida.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenRapida.frx":2156
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   2
      Top             =   2520
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
         MouseIcon       =   "FrmRepOrdenRapida.frx":26E5
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepOrdenRapida.frx":29EF
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepOrdenRapida.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPicker1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPicker2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CommonDialog1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.OptionButton Option6 
         Caption         =   "Todo lo activo"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de reporte"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   6375
         Begin VB.OptionButton Option5 
            Caption         =   "General"
            Height          =   255
            Left            =   1800
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "De Gastos"
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6000
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pendientes de Pago"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pagadas"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todo"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   2280
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Filtrar por fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49807361
         CurrentDate     =   39253
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2400
         TabIndex        =   0
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49807361
         CurrentDate     =   39253
      End
      Begin VB.Label Label1 
         Caption         =   "No. Orden :"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Al :"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Del :"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1800
         Width           =   375
      End
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   6480
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmRepOrdenRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
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
        .ColumnHeaders.Add , , "No. Orden", 2000
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Estado", 1500
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "Impuesto", 1500
        .ColumnHeaders.Add , , "Total", 1500
    End With
End Sub
Private Sub Image10_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim SUMA As String
    Dim Total As String
    Dim Estado As String
    Dim sWhere As String
    Dim tLi As ListItem
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    ConPag = 1
    Total = "0"
    SUMA = "0"
    sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, PROVEEDOR_CONSUMO.RFC, ORDEN_RAPIDA.ESTADO, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.IVA) AS IVA, SUM(ORDEN_RAPIDA_DETALLE.IVARETENIDO) AS RET_IVA, SUM(ORDEN_RAPIDA_DETALLE.ISR) AS ISR, SUM(ORDEN_RAPIDA_DETALLE.IVADIEZ) AS IVA_DIEZ, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL"
    If Check1.Value = 1 Then
        sWhere = " WHERE ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ID_ORDEN_RAPIDA = " & Text2.Text
    End If
    If Option2.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO = 'F'"
    End If
    If Option3.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO IN ('A', 'M')"
    End If
    If Option6.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO NOT IN ('C', 'D')"
    End If
    sBuscar = sBuscar & sWhere & " GROUP BY ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ESTADO, ORDEN_RAPIDA_DETALLE.IVA, ORDEN_RAPIDA_DETALLE.IVARETENIDO, ORDEN_RAPIDA_DETALLE.ISR, ORDEN_RAPIDA_DETALLE.IVADIEZ, ORDEN_RAPIDA_DETALLE.TOTAL, PROVEEDOR_CONSUMO.RFC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If tRs.Fields("ESTADO") = "A" Then
                Estado = "PENDIENTE"
            Else
                If tRs.Fields("ESTADO") = "M" Then
                    Estado = "MODIFICACION"
                Else
                    If tRs.Fields("ESTADO") = "C" Then
                        Estado = "CANCELADA"
                    Else
                        Estado = "PAGADA"
                    End If
                End If
            End If
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            tLi.SubItems(3) = Estado
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(4) = Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("IMPUESTO")) Then tLi.SubItems(5) = Format(tRs.Fields("IMPUESTO"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(6) = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            tRs.MoveNext
        Loop
    End If
    
    
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
Private Sub Image26_Click()
    If Option5.Value = True Then
        RepGeneral
    Else
        RepDepto
    End If
End Sub
Private Sub RepGeneral()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim SUMA As String
    Dim Total As String
    Dim Estado As String
    Dim sWhere As String
    ConPag = 1
    Total = "0"
    SUMA = "0"
    sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ESTADO, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.TOTAL - ORDEN_RAPIDA_DETALLE.SUBTOTAL)AS IMPUESTO FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR"
    If Check1.Value = 1 Then
        sWhere = " WHERE ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ID_ORDEN_RAPIDA = " & Text2.Text
    End If
    If Option2.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO = 'F'"
    End If
    If Option3.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO IN ('A', 'M')"
    End If
    If Option6.Value Then
        If sWhere = "" Then
            sWhere = " WHERE"
        Else
            sWhere = sWhere & " AND"
        End If
        sWhere = sWhere & " ORDEN_RAPIDA.ESTADO NOT IN ('C', 'D')"
    End If
    sBuscar = sBuscar & sWhere & " GROUP BY ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ESTADO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\ReporteOrdenRapida.pdf") Then
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
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ESTADO DE ORDENES DE COMPRA RAPIDAS", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "No. Orden", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 230, "Proveedor", "F2", 8, hCenter
        oDoc.WTextBox Posi, 295, 20, 50, "Fecha", "F2", 8, hCenter
        oDoc.WTextBox Posi, 345, 20, 55, "Estado", "F2", 8, hLeft
        oDoc.WTextBox Posi, 395, 20, 55, "Subtotal", "F2", 8, hLeft
        oDoc.WTextBox Posi, 460, 20, 55, "Impuesto", "F2", 8, hLeft
        oDoc.WTextBox Posi, 525, 20, 55, "Total", "F2", 8, hLeft
        'oDoc.WTextBox Posi, 515, 20, 55, "Surtido", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Not tRs.EOF
            If tRs.Fields("ESTADO") = "A" Then
                Estado = "PENDIENTE"
            Else
                If tRs.Fields("ESTADO") = "M" Then
                    Estado = "MODIFICACION"
                Else
                    If tRs.Fields("ESTADO") = "C" Then
                        Estado = "CANCELADA"
                    Else
                        Estado = "PAGADA"
                    End If
                End If
            End If
            oDoc.WTextBox Posi, 20, 20, 50, tRs.Fields("ID_ORDEN_RAPIDA"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 230, tRs.Fields("NOMBRE"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 295, 20, 50, tRs.Fields("FECHA"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 345, 20, 55, Estado, "F3", 7, hLeft
            oDoc.WTextBox Posi, 395, 20, 50, Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 450, 20, 55, Format(tRs.Fields("IMPUESTO"), "###,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 505, 20, 55, Format(tRs.Fields("TOTAL"), "###,###,##0.00"), "F3", 7, hRight
            SUMA = CDbl(tRs.Fields("TOTAL")) + CDbl(SUMA)
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
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ESTADO DE ORDENES DE COMPRA RAPIDAS", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "No. Orden", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 230, "Proveedor", "F2", 8, hCenter
                oDoc.WTextBox Posi, 295, 20, 50, "Fecha", "F2", 8, hCenter
                oDoc.WTextBox Posi, 345, 20, 55, "Estado", "F2", 8, hLeft
                oDoc.WTextBox Posi, 395, 20, 55, "Subtotal", "F2", 8, hLeft
                oDoc.WTextBox Posi, 460, 20, 55, "Impuesto", "F2", 8, hLeft
                oDoc.WTextBox Posi, 525, 20, 55, "Total", "F2", 8, hLeft
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
            tRs.MoveNext
        Loop
        ' Linea
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 460, 20, 55, "Total: " & Format(SUMA, "###,###,##0.00"), "F3", 8, hLeft
        Posi = Posi + 16
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
    Else
        MsgBox "No se encontraron resultados coincidentes a la busqueda", vbCritical, "SACC"
    End If
End Sub
Private Sub RepDepto()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim SUMA As String
    Dim Total As String
    Dim Estado As String
    Dim sWhere As String
    Dim Depto As String
    Depto = "3A4F5AE2SS"
    ConPag = 1
    Total = "0"
    SUMA = "0"
    sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ESTADO, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.TOTAL - ORDEN_RAPIDA_DETALLE.SUBTOTAL)AS IMPUESTO, ORDEN_RAPIDA.DEPARTAMENTO FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR"
    sWhere = " WHERE ORDEN_RAPIDA.ESTADO = 'F'"
    If Check1.Value = 1 Then
        sWhere = " AND ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    End If
    If Text1.Text <> "" Then
        sWhere = sWhere & " AND PROVEEDOR_CONSUMO.NOMBRE LIKE '%" & Text1.Text & "%'"
    End If
    If Text2.Text <> "" Then
        sWhere = sWhere & " AND ORDEN_RAPIDA.ID_ORDEN_RAPIDA = " & Text2.Text
    End If
    If Option3.Value Then
        sWhere = sWhere & " AND ORDEN_RAPIDA.ESTADO IN ('A', 'M')"
    End If
    sBuscar = sBuscar & sWhere & " GROUP BY ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.ESTADO, ORDEN_RAPIDA.DEPARTAMENTO ORDER BY ORDEN_RAPIDA.DEPARTAMENTO, ORDEN_RAPIDA.FECHA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\RepGastosDepto.pdf") Then
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
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE GASTOS POR DEPARTAMENTO", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "No. Orden", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 330, "Proveedor", "F2", 8, hCenter
        oDoc.WTextBox Posi, 395, 20, 50, "Fecha", "F2", 8, hCenter
        oDoc.WTextBox Posi, 445, 20, 65, "Departamento", "F2", 8, hCenter
        'oDoc.WTextBox Posi, 395, 20, 55, "Subtotal", "F2", 8, hLeft
        'oDoc.WTextBox Posi, 460, 20, 55, "Impuesto", "F2", 8, hLeft
        oDoc.WTextBox Posi, 525, 20, 55, "Total", "F2", 8, hLeft
        'oDoc.WTextBox Posi, 515, 20, 55, "Surtido", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Not tRs.EOF
            'If tRs.Fields("ESTADO") = "A" Then
            '    Estado = "PENDIENTE"
            'Else
            '    If tRs.Fields("ESTADO") = "M" Then
            '        Estado = "MODIFICACION"
            '    Else
            '        If tRs.Fields("ESTADO") = "C" Then
            '            Estado = "CANCELADA"
            '        Else
            '            Estado = "PAGADA"
            '        End If
            '    End If
            'End If
            If Depto <> "3A4F5AE2SS" Then
                If Depto <> tRs.Fields("DEPARTAMENTO") Then
                    Posi = Posi + 6
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 580, Posi
                    oDoc.LineStroke
                    Posi = Posi + 6
                    oDoc.WTextBox Posi, 405, 20, 155, "Total : $ " & Format(SUMA, "###,###,##0.00"), "F3", 7, hRight
                    SUMA = 0
                    Posi = Posi + 20
                End If
            End If
            oDoc.WTextBox Posi, 20, 20, 50, tRs.Fields("ID_ORDEN_RAPIDA"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 330, tRs.Fields("NOMBRE"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 395, 20, 50, tRs.Fields("FECHA"), "F3", 7, hLeft
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then oDoc.WTextBox Posi, 445, 20, 65, tRs.Fields("DEPARTAMENTO"), "F3", 7, hLeft
            'oDoc.WTextBox Posi, 395, 20, 50, Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00"), "F3", 7, hRight
            'oDoc.WTextBox Posi, 450, 20, 55, Format(tRs.Fields("IMPUESTO"), "###,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 505, 20, 55, Format(tRs.Fields("TOTAL"), "###,###,##0.00"), "F3", 7, hRight
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Depto = tRs.Fields("DEPARTAMENTO")
            SUMA = CDbl(tRs.Fields("TOTAL")) + CDbl(SUMA)
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
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE GASTOS POR DEPARTAMENTO", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "No. Orden", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 330, "Proveedor", "F2", 8, hCenter
                oDoc.WTextBox Posi, 395, 20, 50, "Fecha", "F2", 8, hCenter
                oDoc.WTextBox Posi, 445, 20, 65, "Departamento", "F2", 8, hCenter
                'oDoc.WTextBox Posi, 395, 20, 55, "Subtotal", "F2", 8, hLeft
                'oDoc.WTextBox Posi, 460, 20, 55, "Impuesto", "F2", 8, hLeft
                oDoc.WTextBox Posi, 525, 20, 55, "Total", "F2", 8, hLeft
                'oDoc.WTextBox Posi, 515, 20, 55, "Surtido", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
            tRs.MoveNext
        Loop
        ' Linea
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 405, 20, 155, "Total : $ " & Format(SUMA, "###,###,##0.00"), "F3", 7, hRight
        Posi = Posi + 20
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontraron resultados coincidentes a la busqueda", vbCritical, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
