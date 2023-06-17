VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepVentasProgCerradas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte de Ventas Programadas Cerradas"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepVentasProgCerradas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentasProgCerradas.frx":030A
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   4
      Top             =   1320
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepVentasProgCerradas.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepVentasProgCerradas.frx":0BA3
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepVentasProgCerradas.frx":2C85
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPicker2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.OptionButton Option3 
         Caption         =   "Pendientes en Almacén"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Fecha de Facturación"
         Height          =   195
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Fecha de Cierre"
         Height          =   195
         Left            =   2400
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16252929
         CurrentDate     =   39253
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16252929
         CurrentDate     =   39253
      End
      Begin VB.Label Label3 
         Caption         =   "Al :"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Del :"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   6840
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmRepVentasProgCerradas"
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
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim Suma As String
    Dim Total As String
    ConPag = 1
    Total = "0"
    Suma = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Option1.Value Then
        sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, CLIENTE.NOMBRE, PED_CLIEN.FECHA, PED_CLIEN_DETALLE.ID_PRODUCTO, PED_CLIEN_DETALLE.CANTIDAD_PEDIDA, PED_CLIEN_DETALLE.ACTIVO, PED_CLIEN.NO_ORDEN, PED_CLIEN.Estado FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE PED_CLIEN.FECHA_CIERRE BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND PED_CLIEN.Estado IN ('T', 'C') AND PED_CLIEN_DETALLE.ACTIVO = 'N'"
    Else
        If Option2.Value Then
            sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, CLIENTE.NOMBRE, PED_CLIEN.FECHA, PED_CLIEN_DETALLE.ID_PRODUCTO, PED_CLIEN_DETALLE.CANTIDAD_PEDIDA, PED_CLIEN_DETALLE.ACTIVO, PED_CLIEN.NO_ORDEN, PED_CLIEN.Estado FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE PED_CLIEN.Estado IN ('T', 'C') AND PED_CLIEN.FECHA_FACTURACION BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND PED_CLIEN_DETALLE.ACTIVO = 'N'"
        Else
            sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, CLIENTE.NOMBRE, PED_CLIEN.FECHA, PED_CLIEN_DETALLE.ID_PRODUCTO, PED_CLIEN_DETALLE.CANTIDAD_PEDIDA, PED_CLIEN_DETALLE.ACTIVO, PED_CLIEN.NO_ORDEN, PED_CLIEN.Estado FROM PED_CLIEN INNER JOIN PED_CLIEN_DETALLE ON PED_CLIEN.NO_PEDIDO = PED_CLIEN_DETALLE.NO_PEDIDO INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE PED_CLIEN.Estado IN ('I') "
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\ReporteVentasProgramadas.pdf") Then
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
        If Option3.Value Then
            oDoc.WTextBox 30, 380, 20, 250, "Al: " & Date, "F3", 8, hCenter
        Else
            oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
        End If
        oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        If Option1.Value Then
            oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS POR CIERRE EN ALMACEN", "F3", 8, hCenter
        Else
            If Option2.Value Then
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS POR FACTURACION DE VENTAS", "F3", 8, hCenter
            Else
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS EN ALMACEN", "F3", 8, hCenter
            End If
        End If
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 60, "No. Venta", "F2", 8, hCenter
        oDoc.WTextBox Posi, 65, 20, 70, "Fecha Captura", "F2", 8, hCenter
        oDoc.WTextBox Posi, 135, 20, 230, "Cliente", "F2", 8, hCenter
        oDoc.WTextBox Posi, 355, 20, 95, "Producto", "F2", 8, hCenter
        oDoc.WTextBox Posi, 460, 20, 55, "Cantidad", "F2", 8, hCenter
        oDoc.WTextBox Posi, 515, 20, 55, "Orden", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        Do While Not tRs.EOF
            oDoc.WTextBox Posi, 10, 20, 60, tRs.Fields("NO_PEDIDO"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 70, tRs.Fields("FECHA"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 135, 20, 230, tRs.Fields("NOMBRE"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 355, 20, 95, tRs.Fields("ID_PRODUCTO"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 460, 20, 55, tRs.Fields("CANTIDAD_PEDIDA"), "F3", 7, hRight
            oDoc.WTextBox Posi, 515, 20, 55, tRs.Fields("NO_ORDEN"), "F3", 7, hRight
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
                If Option3.Value Then
                    oDoc.WTextBox 30, 380, 20, 250, "Al: " & Date, "F3", 8, hCenter
                Else
                    oDoc.WTextBox 30, 380, 20, 250, "Del: " & DTPicker1.Value & " Al: " & DTPicker2.Value, "F3", 8, hCenter
                End If
                oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                If Option1.Value Then
                    oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS POR CIERRE EN ALMACEN", "F3", 8, hCenter
                Else
                    If Option2.Value Then
                        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS POR FACTURACION DE VENTAS", "F3", 8, hCenter
                    Else
                        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS PROGRAMADAS EN ALMACEN", "F3", 8, hCenter
                    End If
                End If
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 60, "FECHA", "F2", 8, hCenter
                oDoc.WTextBox Posi, 65, 20, 70, "ORDEN", "F2", 8, hCenter
                oDoc.WTextBox Posi, 135, 20, 230, "PROVEEDOR", "F2", 8, hCenter
                oDoc.WTextBox Posi, 355, 20, 95, "PRODUCTO", "F2", 8, hCenter
                oDoc.WTextBox Posi, 460, 20, 55, "ENTRADA", "F2", 8, hCenter
                oDoc.WTextBox Posi, 515, 20, 55, "CANTIDAD", "F2", 8, hCenter
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
        MsgBox "No se encontraron resultados en la busqueda!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Option1_Click()
    If Option3.Value Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
End Sub
Private Sub Option2_Click()
    If Option3.Value Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
End Sub
Private Sub Option3_Click()
    If Option3.Value Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
End Sub
