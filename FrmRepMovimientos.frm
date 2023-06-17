VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepMovimientos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Movimientos"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5520
      TabIndex        =   7
      Top             =   240
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepMovimientos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepMovimientos.frx":030A
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepMovimientos.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepMovimientos.frx":0BA3
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepMovimientos.frx":2C85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DTPicker2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.OptionButton Option3 
         Caption         =   "Estado de Egresos"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Estado de Resultados"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Value           =   -1  'True
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50987009
         CurrentDate     =   42647
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50987009
         CurrentDate     =   42647
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   5760
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmRepMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
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
End Sub
Private Sub Image26_Click()
    If Option2.Value Then
        EdoResultadosDet
    End If
    If Option3.Value Then
        EdoEgresos
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub EdoResultadosDet()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim sBuscar As String
    Dim VentCont As Double
    Dim VentCred As Double
    Dim Cobranza As Double
    Dim OCN As Double
    Dim OCNP As Double
    Dim OCND As Double
    Dim OCI As Double
    Dim OCIP As Double
    Dim OCID As Double
    Dim OCR As Double
    Dim OCRP As Double
    Dim OCRD As Double
    Dim Dolar As Double
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\EstadoDeResultados.pdf") Then
        Exit Sub
    End If
    sBuscar = "SELECT TOP 1 VENTA AS DOLAR From Dolar ORDER BY ID_DOLAR DESC"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
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
    oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
    ' ENCABEZADO DEL DETALLE
    oDoc.WTextBox 170, 10, 90, 570, "INGRESOS", "F2", 12, hCenter, , , 1
    oDoc.WTextBox 100, 10, 90, 570, "ESTADO DE RESULTADOS", "F2", 16, hCenter
    oDoc.WTextBox 120, 10, 90, 570, "PERIODO DEL " & DTPicker1.Value & " AL " & DTPicker2.Value, "F2", 16, hCenter
    sBuscar = "SELECT SUM(TOTAL) AS TOTAL FROM VENTAS WHERE (FACTURADO IN (0, 1)) AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND UNA_EXIBICION = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 180, 20, 90, 570, "(1) Ventas de contado : ", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 180, 300, 90, 200, Format(tRs.Fields("TOTAL"), "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            VentCont = tRs.Fields("TOTAL")
        End If
    Else
        oDoc.WTextBox 180, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(TOTAL) AS TOTAL FROM VENTAS WHERE (FACTURADO IN (0, 1)) AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND UNA_EXIBICION = 'N'"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 195, 20, 90, 570, "(2) Ventas de credito : ", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 195, 300, 90, 200, Format(tRs.Fields("TOTAL"), "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            VentCred = tRs.Fields("TOTAL")
        End If
    Else
        oDoc.WTextBox 195, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(ABONOS_CUENTA.CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ABONOS_CUENTA.FECHA BETWEEN '" & DTPicker1.Value & "' AND  '" & DTPicker2.Value & "'"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 210, 20, 90, 570, "(3) Cobranza de creditos : ", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 210, 300, 90, 200, Format(tRs.Fields("TOTAL"), "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            Cobranza = tRs.Fields("TOTAL")
        End If
    Else
        oDoc.WTextBox 210, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 225, 20, 90, 200, "(1 + 2) TOTAL VENTAS : ", "F2", 10, hLeft
    oDoc.WTextBox 225, 300, 90, 200, Format(VentCred + VentCont, "$#,##0.00"), "F2", 10, hRight
    oDoc.WTextBox 240, 20, 90, 200, "(1 + 3) TOTAL INGRESOS : ", "F2", 10, hLeft
    oDoc.WTextBox 240, 300, 90, 200, Format(VentCont + Cobranza, "$#,##0.00"), "F2", 10, hRight
    
    
    oDoc.WTextBox 270, 10, 190, 570, "EGRESOS", "F2", 12, hCenter, , , 1
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X', 'Y') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 280, 20, 90, 570, "(4) Ordenes de Compra Nacionales :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCN = OCN + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCN = OCN + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 280, 300, 90, 200, Format(OCN, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 280, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('Y') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 295, 30, 90, 570, "(5) Pagadas :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCNP = OCNP + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCNP = OCNP + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 295, 300, 90, 200, Format(OCNP, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 295, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 310, 30, 90, 570, "(6) Pendientes de Pago :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCND = OCND + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCND = OCND + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 310, 300, 90, 200, Format(OCND, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 310, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    
    sBuscar = "SELECT AVG(VENTA) AS DOLAR From Dolar WHERE (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL  FROM ORDEN_COMPRA WHERE (TIPO = 'I') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('Y')"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 340, 30, 90, 570, "(8) Pagadas :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 340, 150, 90, 200, Format(Dolar, "$#,##0.00"), "F1", 10, hRight
        oDoc.WTextBox 340, 300, 90, 200, Format(tRs.Fields("TOTAL") * Dolar, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCIP = tRs.Fields("TOTAL") * Dolar
        End If
    Else
        oDoc.WTextBox 340, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT TOP 1 VENTA AS DOLAR From Dolar ORDER BY ID_DOLAR DESC"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL  FROM ORDEN_COMPRA WHERE (TIPO = 'I') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X')"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 355, 30, 90, 570, "(9) Pendientes de Pago :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 355, 150, 90, 200, Format(Dolar, "$#,##0.00"), "F1", 10, hRight
        oDoc.WTextBox 355, 300, 90, 200, Format(tRs.Fields("TOTAL") * Dolar, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCID = tRs.Fields("TOTAL") * Dolar
        End If
    Else
        oDoc.WTextBox 355, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 370, 20, 90, 570, "(10) Gastos (Ordenes Rapidas) : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL , MONEDA FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO IN ('A', 'M', 'F') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCR = OCR + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCR = OCR + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 370, 300, 90, 200, Format(OCR, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 370, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 325, 20, 90, 570, "(7) Ordenes de Compra Internacionales : ", "F1", 10, hLeft


    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 325, 300, 90, 200, Format(OCID + OCIP, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCI = OCID + OCIP
        End If
    Else
        oDoc.WTextBox 325, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    
    
    oDoc.WTextBox 385, 30, 90, 570, "(11) Pagados : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, MONEDA FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO = 'F' GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRP = OCRP + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRP = OCRP + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 385, 300, 90, 200, Format(OCRP, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 385, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 400, 30, 90, 570, "(12) Pendientes de Pago : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, MONEDA  FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO IN ('A', 'M') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRD = OCRD + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRD = OCRD + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 400, 300, 90, 200, Format(OCRD, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 400, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 415, 20, 90, 200, "(5 + 8 + 11) TOTAL PAGOS : ", "F2", 10, hLeft
    oDoc.WTextBox 415, 300, 90, 200, Format(OCNP + OCIP + OCRP, "$#,##0.00"), "F2", 10, hRight
    oDoc.WTextBox 430, 20, 90, 200, "(6 + 9 + 12) TOTAL PENDIENTE : ", "F2", 10, hLeft
    oDoc.WTextBox 430, 300, 90, 200, Format(OCND + OCID + OCRD, "$#,##0.00"), "F2", 10, hRight
    
    
    oDoc.WTextBox 470, 10, 50, 570, "UTILIDAD", "F2", 12, hCenter, , , 1
    oDoc.WTextBox 485, 20, 90, 570, "(1 + 3 - 5 - 8 - 11) UTILIDAD :", "F1", 10, hLeft
    oDoc.WTextBox 485, 300, 90, 200, Format(VentCont + VentCred - OCN - OCI - OCR, "$#,##0.00"), "F1", 10, hRight
    oDoc.WTextBox 500, 20, 90, 570, "(1 + 2 - 4 - 7 - 10) FLUJO DE EFECTIVO :", "F1", 10, hLeft
    oDoc.WTextBox 500, 300, 90, 200, Format(VentCont + Cobranza - OCNP - OCIP - OCRP, "$#,##0.00"), "F1", 10, hRight

    
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 800
    oDoc.WLineTo 580, 800
    oDoc.LineStroke

    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 810
    oDoc.WLineTo 580, 810
    oDoc.LineStroke

    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 820
    oDoc.WLineTo 580, 820
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub EdoEgresos()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim sBuscar As String
    Dim VentCont As Double
    Dim VentCred As Double
    Dim Cobranza As Double
    Dim OCN As Double
    Dim OCNP As Double
    Dim OCND As Double
    Dim OCI As Double
    Dim OCIP As Double
    Dim OCID As Double
    Dim OCR As Double
    Dim OCRP As Double
    Dim OCRD As Double
    Dim Dolar As Double
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\EstadoDeResultados.pdf") Then
        Exit Sub
    End If
    sBuscar = "SELECT TOP 1 VENTA AS DOLAR From Dolar ORDER BY ID_DOLAR DESC"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
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
    oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
    ' ENCABEZADO DEL DETALLE

    oDoc.WTextBox 100, 10, 20, 570, "ESTADO DE EGRESOS", "F2", 16, hCenter
    oDoc.WTextBox 120, 10, 20, 570, "PERIODO DEL " & DTPicker1.Value & " AL " & DTPicker2.Value, "F2", 16, hCenter
    
    oDoc.WTextBox 170, 10, 190, 570, "EGRESOS", "F2", 12, hCenter, , , 1
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X', 'Y') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 180, 20, 20, 570, "(1) Ordenes de Compra Nacionales :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCN = OCN + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCN = OCN + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 180, 300, 20, 200, Format(OCN, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 180, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('Y') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 195, 30, 20, 570, "(2) Pagadas :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCNP = OCNP + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCNP = OCNP + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 195, 300, 20, 200, Format(OCNP, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 195, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX + OTROS_CARGOS) AS TOTAL, MONEDA  FROM ORDEN_COMPRA WHERE (TIPO = 'N') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 210, 30, 20, 570, "(3) Pendientes de Pago :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCND = OCND + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCND = OCND + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 210, 300, 20, 200, Format(OCND, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 210, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    
    sBuscar = "SELECT AVG(VENTA) AS DOLAR From Dolar WHERE (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "')"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL  FROM ORDEN_COMPRA WHERE (TIPO = 'I') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('Y')"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 240, 30, 20, 570, "(5) Pagadas :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 240, 150, 20, 200, Format(Dolar, "$#,##0.00"), "F1", 10, hRight
        oDoc.WTextBox 240, 300, 20, 200, Format(tRs.Fields("TOTAL") * Dolar, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCIP = tRs.Fields("TOTAL") * Dolar
        End If
    Else
        oDoc.WTextBox 240, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    sBuscar = "SELECT TOP 1 VENTA AS DOLAR From Dolar ORDER BY ID_DOLAR DESC"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        If Not IsNull(tRs2.Fields("DOLAR")) Then Dolar = tRs2.Fields("DOLAR")
    End If
    sBuscar = "SELECT SUM(TOTAL - DISCOUNT + FREIGHT + TAX  + OTROS_CARGOS) AS TOTAL  FROM ORDEN_COMPRA WHERE (TIPO = 'I') AND (FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND CONFIRMADA IN ('X')"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 255, 30, 20, 570, "(6) Pendientes de Pago :", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 255, 150, 20, 200, Format(Dolar, "$#,##0.00"), "F1", 10, hRight
        oDoc.WTextBox 255, 300, 20, 200, Format(tRs.Fields("TOTAL") * Dolar, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCID = tRs.Fields("TOTAL") * Dolar
        End If
    Else
        oDoc.WTextBox 255, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 270, 20, 20, 570, "(7) Gastos (Ordenes Rapidas) : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, MONEDA FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO IN ('A', 'M', 'F') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCR = OCR + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCR = OCR + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 270, 300, 20, 200, Format(OCR, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 270, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 225, 20, 20, 570, "(4) Ordenes de Compra Internacionales : ", "F1", 10, hLeft
    If Not (tRs.EOF And tRs.BOF) Then
        oDoc.WTextBox 225, 300, 20, 200, Format(OCID + OCIP, "$#,##0.00"), "F1", 10, hRight
        If Not IsNull(tRs.Fields("TOTAL")) Then
            OCI = OCID + OCIP
        End If
    Else
        oDoc.WTextBox 225, 300, 90, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 285, 30, 20, 570, "(8) Pagados : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, MONEDA  FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO = 'F'  GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRP = OCRP + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRP = OCRP + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 285, 300, 20, 200, Format(OCRP, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 285, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 300, 30, 20, 570, "(9) Pendientes de Pago : ", "F1", 10, hLeft
    sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL, MONEDA FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND ESTADO IN ('A', 'M') GROUP BY MONEDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("MONEDA") = "DOLARES" Then
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRD = OCRD + (tRs.Fields("TOTAL") * Dolar)
                End If
            Else
                If Not IsNull(tRs.Fields("TOTAL")) Then
                    OCRD = OCRD + tRs.Fields("TOTAL")
                End If
            End If
            tRs.MoveNext
        Loop
        oDoc.WTextBox 300, 300, 20, 200, Format(OCRD, "$#,##0.00"), "F1", 10, hRight
    Else
        oDoc.WTextBox 300, 300, 20, 200, "$0.00", "F1", 10, hRight
    End If
    oDoc.WTextBox 315, 20, 20, 200, "(2 + 5 + 8) TOTAL PAGOS : ", "F2", 10, hLeft
    oDoc.WTextBox 315, 300, 20, 200, Format(OCNP + OCIP + OCRP, "$#,##0.00"), "F2", 10, hRight
    oDoc.WTextBox 330, 20, 20, 200, "(3 + 6 + 9) TOTAL PENDIENTE : ", "F2", 10, hLeft
    oDoc.WTextBox 330, 300, 20, 200, Format(OCND + OCID + OCRD, "$#,##0.00"), "F2", 10, hRight
    
    

    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 800
    oDoc.WLineTo 580, 800
    oDoc.LineStroke

    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 810
    oDoc.WLineTo 580, 810
    oDoc.LineStroke

    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 820
    oDoc.WLineTo 580, 820
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
