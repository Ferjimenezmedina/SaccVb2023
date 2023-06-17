VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepCostosVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Costos de Ventas"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepCostosVentas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCostosVentas.frx":030A
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepCostosVentas.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCostosVentas.frx":0BA3
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepCostosVentas.frx":2C85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DTPicker2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50724865
         CurrentDate     =   43020
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50724865
         CurrentDate     =   43020
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   5400
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmRepCostosVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
           "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim Suma As String
    Dim Total As String
    Dim SumCosto As Double
    ConPag = 1
    Total = "0"
    Suma = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, SUM(VENTAS_DETALLE.CANTIDAD) AS CANT, VENTAS_DETALLE.PRECIO_VENTA , VENTAS.FECHA FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA WHERE (VENTAS_DETALLE.ID_PRODUCTO NOT IN ('SANCION', 'DESCUENTO')) AND VENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' GROUP BY VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, VENTAS_DETALLE.PRECIO_VENTA , VENTAS.FECHA "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\RepCostosVentas.pdf") Then
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
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE COSTOS DE VENTAS", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "PRODUCTO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 300, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 350, 20, 50, "CANTIDAD", "F2", 8, hCenter
        oDoc.WTextBox Posi, 400, 20, 50, "COSTO", "F2", 8, hCenter
        oDoc.WTextBox Posi - 10, 450, 20, 50, "PRECIO VENTA", "F2", 8, hCenter
        oDoc.WTextBox Posi, 500, 20, 60, "TOTAL COSTO", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then oDoc.WTextBox Posi, 10, 20, 60, tRs.Fields("ID_PRODUCTO"), "F3", 7, hLeft
            If Not IsNull(tRs.Fields("DESCRIPCION")) Then oDoc.WTextBox Posi, 80, 20, 300, Mid(tRs.Fields("DESCRIPCION"), 1, 58), "F3", 7, hLeft
            If Not IsNull(tRs.Fields("CANT")) Then oDoc.WTextBox Posi, 350, 20, 50, Format(tRs.Fields("CANT"), "###,###,##0.00"), "F3", 7, hRight
            sBuscar = "SELECT TOP (1) ORDEN_COMPRA_DETALLE.PRECIO FROM ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA.FECHA <= '" & Format(tRs.Fields("FECHA"), "dd/MM/yyyy") & "') AND (ORDEN_COMPRA_DETALLE.ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "') ORDER BY ORDEN_COMPRA.ID_ORDEN_COMPRA DESC"
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.EOF And tRs2.BOF) Then
                If Not IsNull(tRs2.Fields("PRECIO")) Then oDoc.WTextBox Posi, 400, 20, 50, Format(tRs2.Fields("PRECIO"), "###,###,##0.00"), "F3", 7, hRight
                If Not IsNull(tRs.Fields("CANT")) Then oDoc.WTextBox Posi, 500, 20, 50, Format(CDbl(tRs.Fields("CANT") * tRs2.Fields("PRECIO")), "###,###,##0.00"), "F3", 7, hRight
                SumCosto = SumCosto + CDbl(tRs.Fields("CANT") * tRs2.Fields("PRECIO"))
            Else
                oDoc.WTextBox Posi, 400, 20, 50, "SIN COMPRAS", "F3", 7, hRight
            End If
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then oDoc.WTextBox Posi, 450, 20, 50, Format(tRs.Fields("PRECIO_VENTA"), "###,###,##0.00"), "F3", 7, hRight
            'Suma = Suma + (tRs.Fields("PRECIO_COSTO") * tRs.Fields("CANT"))
            Posi = Posi + 12
            If Posi >= 650 Then
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
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE COSTOS DE VENTAS", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "PRODUCTO", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 300, "DESCRIPCION", "F2", 8, hCenter
                oDoc.WTextBox Posi, 350, 20, 50, "CANTIDAD", "F2", 8, hCenter
                oDoc.WTextBox Posi, 400, 20, 50, "COSTO", "F2", 8, hCenter
                oDoc.WTextBox Posi - 10, 450, 20, 50, "PRECIO VENTA", "F2", 8, hCenter
                oDoc.WTextBox Posi, 500, 20, 60, "TOTAL COSTO", "F2", 8, hCenter
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
        oDoc.WTextBox Posi, 430, 20, 150, "TOTAL : " & Format(SumCosto, "###,###,##0.00"), "F2", 8, hCenter
        'oDoc.WTextBox Posi, 450, 20, 50, "TOTAL", "F3", 7, hLeft
        'If Not IsNull(Suma) Then oDoc.WTextBox Posi, 515, 20, 55, Format(Suma, "$ #,###,##0.00"), "F3", 7, hRight
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
