VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepLicitaciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Licitaciones"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle"
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   15
      Top             =   4800
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepLicitaciones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepLicitaciones.frx":030A
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   5
      Top             =   6000
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepLicitaciones.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepLicitaciones.frx":0BA3
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepLicitaciones.frx":2C85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.OptionButton Option3 
         Caption         =   "Todo"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Notas de Venta"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Facturas"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contratos Activos"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
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
         Left            =   5520
         Picture         =   "FrmRepLicitaciones.frx":2CA1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   6600
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4895
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   6600
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   6600
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9240
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmRepLicitaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    ComprasLicitacion
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
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
        .ColumnHeaders.Add , , "Id Cliente", 2000
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "No. Contrato", 1500
        .ColumnHeaders.Add , , "No. Licitacion", 1500
        .ColumnHeaders.Add , , "Fecha Inicio", 1500
        .ColumnHeaders.Add , , "Fecha Fin", 1500
        .ColumnHeaders.Add , , "Total Facturado", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Precio de Venta", 1500
    End With
    ComprasLicitacion
End Sub
Private Sub ComprasLicitacion()
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim sFactura As String
    If Option1.Value = True Then
        sFactura = "1"
    Else
        If Option2.Value = True Then
            sFactura = "0"
        Else
            sFactura = "0, 1"
        End If
    End If
    sBuscar = "SELECT CLIENTE.ID_CLIENTE, CLIENTE.NOMBRE, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.FECHA_FIN, LICITACIONES.FECHA_INICIO, (SELECT SUM(VENTAS_DETALLE.CANTIDAD * VENTAS_DETALLE.PRECIO_VENTA) AS TOTAL FROM VENTAS AS VE INNER JOIN VENTAS_DETALLE ON VE.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN LICITACIONES AS LIC ON VE.ID_CLIENTE = LIC.ID_CLIENTE AND VE.FECHA BETWEEN LIC.FECHA_INICIO AND LIC.FECHA_FIN AND VENTAS_DETALLE.ID_PRODUCTO = LIC.ID_PRODUCTO WHERE (LIC.NO_CONTRATO = LICITACIONES.NO_CONTRATO) AND (VE.ID_CLIENTE = CLIENTE.ID_CLIENTE) AND (VE.FACTURADO IN (" & sFactura & ")) GROUP BY LIC.ID_CLIENTE, LIC.NO_CONTRATO) AS TOTAL FROM LICITACIONES INNER JOIN CLIENTE ON LICITACIONES.ID_CLIENTE = CLIENTE.ID_CLIENTE "
    sBuscar = sBuscar & " WHERE CLIENTE.NOMBRE LIKE '%" & Text2.Text & "%'"
    If Check1.Value = 1 Then
        sBuscar = sBuscar & " AND LICITACIONES.FECHA_INICIO <= GETDATE() AND LICITACIONES.FECHA_FIN >= GETDATE()"
    End If
    sBuscar = sBuscar & " GROUP BY CLIENTE.ID_CLIENTE, CLIENTE.NOMBRE, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.FECHA_FIN, LICITACIONES.FECHA_INICIO ORDER BY CLIENTE.ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("NO_CONTRATO")) Then tLi.SubItems(2) = tRs.Fields("NO_CONTRATO")
            If Not IsNull(tRs.Fields("NO_LICITACION")) Then tLi.SubItems(3) = tRs.Fields("NO_LICITACION")
            If Not IsNull(tRs.Fields("FECHA_INICIO")) Then tLi.SubItems(4) = tRs.Fields("FECHA_INICIO")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(5) = tRs.Fields("FECHA_FIN")
            If Not IsNull(tRs.Fields("TOTAL")) Then
                tLi.SubItems(6) = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            Else
                tLi.SubItems(6) = "0.00"
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim SumaT As Double
    Dim sFactura As String
    If Option1.Value = True Then
        sFactura = "1"
    Else
        If Option2.Value = True Then
            sFactura = "0"
        Else
            sFactura = "0, 1"
        End If
    End If
    Cont = 1
    If ListView1.ListItems.Count > 0 Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\ReporteLicitaciones.pdf") Then
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
        'oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        'oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ENTRADAS DE ORDENES RAPIDAS", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 40, "Id Clie.", "F2", 8, hCenter
        oDoc.WTextBox Posi, 50, 20, 250, "Nombre", "F2", 8, hCenter
        oDoc.WTextBox Posi, 300, 20, 60, "No. Contrato", "F2", 8, hLeft
        oDoc.WTextBox Posi, 360, 20, 60, "No. Lici.", "F2", 8, hLeft
        oDoc.WTextBox Posi, 420, 20, 50, "Fecha Ini.", "F2", 8, hCenter
        oDoc.WTextBox Posi, 470, 20, 50, "Fecha Fin", "F2", 8, hCenter
        oDoc.WTextBox Posi, 520, 20, 50, "Total Venta", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Cont <= ListView1.ListItems.Count
            oDoc.WTextBox Posi, 10, 20, 40, ListView1.ListItems.Item(Cont), "F3", 7, hLeft
            oDoc.WTextBox Posi, 50, 20, 250, Mid(ListView1.ListItems.Item(Cont).SubItems(1), 1, 57), "F3", 7, hLeft
            oDoc.WTextBox Posi, 300, 20, 60, Mid(ListView1.ListItems.Item(Cont).SubItems(2), 1, 12), "F3", 7, hLeft
            oDoc.WTextBox Posi, 360, 20, 60, Mid(ListView1.ListItems.Item(Cont).SubItems(3), 1, 12), "F3", 7, hLeft
            oDoc.WTextBox Posi, 420, 20, 50, ListView1.ListItems.Item(Cont).SubItems(4), "F3", 7, hCenter
            oDoc.WTextBox Posi, 470, 20, 50, ListView1.ListItems.Item(Cont).SubItems(5), "F3", 7, hCenter
            oDoc.WTextBox Posi, 520, 20, 50, ListView1.ListItems.Item(Cont).SubItems(6), "F3", 7, hRight
            SumaT = SumaT + CDbl(ListView1.ListItems.Item(Cont).SubItems(6))
            If Check2.Value = 1 Then
                sBuscar = "SELECT VENTAS_DETALLE_1.ID_PRODUCTO, VENTAS_DETALLE_1.DESCRIPCION, SUM(VENTAS_DETALLE_1.CANTIDAD) AS TOTAL, VENTAS_DETALLE_1.Precio_Venta , Ventas.ID_CLIENTE FROM VENTAS INNER JOIN VENTAS_DETALLE AS VENTAS_DETALLE_1 ON VENTAS.ID_VENTA = VENTAS_DETALLE_1.ID_VENTA INNER JOIN LICITACIONES ON VENTAS.ID_CLIENTE = LICITACIONES.ID_CLIENTE AND VENTAS.FECHA BETWEEN LICITACIONES.FECHA_INICIO AND LICITACIONES.FECHA_FIN AND VENTAS_DETALLE_1.ID_PRODUCTO = LICITACIONES.ID_PRODUCTO WHERE (LICITACIONES.NO_CONTRATO = '" & ListView1.ListItems.Item(Cont).SubItems(2) & "') AND (VENTAS.ID_CLIENTE = " & ListView1.ListItems.Item(Cont) & ") AND (VENTAS.FACTURADO IN (" & sFactura & ")) GROUP BY VENTAS_DETALLE_1.ID_PRODUCTO, VENTAS_DETALLE_1.DESCRIPCION, VENTAS_DETALLE_1.PRECIO_VENTA, VENTAS.ID_CLIENTE"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Posi = Posi + 10
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 50, Posi
                    oDoc.WLineTo 530, Posi
                    oDoc.LineStroke
                    Posi = Posi + 5
                    oDoc.WTextBox Posi, 50, 20, 80, "Id Producto", "F2", 8, hLeft
                    oDoc.WTextBox Posi, 130, 20, 300, "Descripcion", "F2", 8, hLeft
                    oDoc.WTextBox Posi, 430, 20, 50, "Cantidad", "F2", 8, hLeft
                    oDoc.WTextBox Posi, 480, 20, 50, "Precio", "F2", 8, hCenter
                    Do While Not tRs.EOF
                        If Posi >= 780 Then
                            oDoc.NewPage A4_Vertical
                            oDoc.WImage 70, 40, 43, 161, "Logo"
                            sBuscar = "SELECT * FROM EMPRESA  "
                            Set tRs1 = cnn.Execute(sBuscar)
                            oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                            oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                            'oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                            oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                            oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                            'oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                            oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                            ' ENCABEZADO DEL DETALLE
                            oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ENTRADAS DE ORDENES RAPIDAS", "F3", 8, hCenter
                            Posi = 120
                            oDoc.WTextBox Posi, 10, 20, 40, "Id Clie.", "F2", 8, hCenter
                            oDoc.WTextBox Posi, 50, 20, 250, "Nombre", "F2", 8, hCenter
                            oDoc.WTextBox Posi, 300, 20, 60, "No. Contrato", "F2", 8, hLeft
                            oDoc.WTextBox Posi, 360, 20, 60, "No. Lici.", "F2", 8, hLeft
                            oDoc.WTextBox Posi, 420, 20, 50, "Fecha Ini.", "F2", 8, hCenter
                            oDoc.WTextBox Posi, 470, 20, 50, "Fecha Fin", "F2", 8, hCenter
                            oDoc.WTextBox Posi, 520, 20, 50, "Total Venta", "F2", 8, hCenter
                            Posi = Posi + 12
                            ' Linea
                            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                            oDoc.MoveTo 10, Posi
                            oDoc.WLineTo 580, Posi
                            oDoc.LineStroke
                            Posi = Posi + 6
                        End If
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 50, 20, 80, tRs.Fields("ID_PRODUCTO"), "F3", 7, hLeft
                        oDoc.WTextBox Posi, 130, 20, 300, Mid(tRs.Fields("DESCRIPCION"), 1, 58), "F3", 7, hLeft
                        oDoc.WTextBox Posi, 430, 20, 50, tRs.Fields("TOTAL"), "F3", 7, hRight
                        oDoc.WTextBox Posi, 480, 20, 50, tRs.Fields("PRECIO_VENTA"), "F3", 7, hRight
                        tRs.MoveNext
                    Loop
                    Posi = Posi + 10
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 50, Posi
                    oDoc.WLineTo 530, Posi
                    oDoc.LineStroke
                End If
            End If
            Posi = Posi + 12
            If Posi >= 780 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA  "
                Set tRs1 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                'oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                'oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ENTRADAS DE ORDENES RAPIDAS", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 40, "Id Clie.", "F2", 8, hCenter
                oDoc.WTextBox Posi, 50, 20, 250, "Nombre", "F2", 8, hCenter
                oDoc.WTextBox Posi, 300, 20, 60, "No. Contrato", "F2", 8, hLeft
                oDoc.WTextBox Posi, 360, 20, 60, "No. Lici.", "F2", 8, hLeft
                oDoc.WTextBox Posi, 420, 20, 50, "Fecha Ini.", "F2", 8, hCenter
                oDoc.WTextBox Posi, 470, 20, 50, "Fecha Fin", "F2", 8, hCenter
                oDoc.WTextBox Posi, 520, 20, 50, "Total Venta", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
            Cont = Cont + 1
        Loop
        ' Linea
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 420, 20, 140, "Total : " & Format(SumaT, "###,###,##0.00"), "F2", 8, hRight
        oDoc.PDFClose
        oDoc.Show
    End If

End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim Total As Double
    Dim sFactura As String
    If Option1.Value = True Then
        sFactura = "1"
    Else
        If Option2.Value = True Then
            sFactura = "0"
        Else
            sFactura = "0, 1"
        End If
    End If
    Total = 0
    ListView2.ListItems.Clear
    sBuscar = "SELECT VENTAS_DETALLE_1.ID_PRODUCTO, VENTAS_DETALLE_1.DESCRIPCION, SUM(VENTAS_DETALLE_1.CANTIDAD) AS TOTAL, VENTAS_DETALLE_1.Precio_Venta , Ventas.ID_CLIENTE FROM VENTAS INNER JOIN VENTAS_DETALLE AS VENTAS_DETALLE_1 ON VENTAS.ID_VENTA = VENTAS_DETALLE_1.ID_VENTA INNER JOIN LICITACIONES ON VENTAS.ID_CLIENTE = LICITACIONES.ID_CLIENTE AND VENTAS.FECHA BETWEEN LICITACIONES.FECHA_INICIO AND LICITACIONES.FECHA_FIN AND VENTAS_DETALLE_1.ID_PRODUCTO = LICITACIONES.ID_PRODUCTO WHERE (LICITACIONES.NO_CONTRATO = '" & Item.SubItems(2) & "') AND (VENTAS.ID_CLIENTE = " & Item & ") AND (VENTAS.FACTURADO IN (" & sFactura & ")) GROUP BY VENTAS_DETALLE_1.ID_PRODUCTO, VENTAS_DETALLE_1.DESCRIPCION, VENTAS_DETALLE_1.PRECIO_VENTA, VENTAS.ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
            Total = Total + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("TOTAL")))
            tRs.MoveNext
        Loop
    End If
    Text1.Text = Format(Total, "###,###,##0.00")
    sBuscar = "SELECT SUM(VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD) AS TOTAL FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA WHERE (VENTAS.ID_CLIENTE IN (" & Item & ")) AND (VENTAS.FECHA BETWEEN '" & Item.SubItems(4) & "' AND '" & Item.SubItems(5) & "') AND (VENTAS_DETALLE.ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO From LICITACIONES WHERE (ID_CLIENTE = dbo.VENTAS.ID_CLIENTE) AND (FECHA_INICIO <= GETDATE()) AND (FECHA_FIN >= GETDATE()))) AND (VENTAS.FACTURADO IN (0, 1)) GROUP BY VENTAS.ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Label3.Caption = "Con ventas fuera del contrato de licitación por $" & Format(tRs.Fields("TOTAL"), "###,###,##0.00")
    Else
        Label3.Caption = ""
    End If
End Sub

