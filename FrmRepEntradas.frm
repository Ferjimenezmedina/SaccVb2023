VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepEntradas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte de Entradas a Almacen"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepEntradas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepEntradas.frx":030A
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepEntradas.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepEntradas.frx":0BA3
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
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
      TabPicture(0)   =   "FrmRepEntradas.frx":2C85
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
      Tab(0).ControlCount=   5
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   15990785
         CurrentDate     =   39253
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   15990785
         CurrentDate     =   39253
      End
      Begin VB.Label Label3 
         Caption         =   "Al :"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Del :"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
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
Attribute VB_Name = "FrmRepEntradas"
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
    sBuscar = "SELECT PROVEEDOR.NOMBRE, ENTRADAS.FECHA, ENTRADA_PRODUCTO.ID_PRODUCTO, ENTRADA_PRODUCTO.CANTIDAD, ENTRADA_PRODUCTO.PRECIO, ORDEN_COMPRA_DETALLE.CANTIDAD AS CANTIDAD_ORIGINAL, ORDEN_COMPRA.NUM_ORDEN, ORDEN_COMPRA.Tipo FROM ENTRADAS INNER JOIN PROVEEDOR ON ENTRADAS.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR INNER JOIN ENTRADA_PRODUCTO ON ENTRADAS.ID_ENTRADA = ENTRADA_PRODUCTO.ID_ENTRADA INNER JOIN ORDEN_COMPRA ON ENTRADAS.ID_ORDEN_COMPRA = ORDEN_COMPRA.ID_ORDEN_COMPRA INNER JOIN ORDEN_COMPRA_DETALLE ON ORDEN_COMPRA.ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA AND ENTRADA_PRODUCTO.ID_PRODUCTO = ORDEN_COMPRA_DETALLE.ID_PRODUCTO WHERE ENTRADAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\NotasCanceladas.pdf") Then
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
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ENTRADAS", "F3", 8, hCenter
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
        ' DETALLE
        Do While Not tRs.EOF
            oDoc.WTextBox Posi, 10, 20, 60, tRs.Fields("FECHA"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 65, 20, 70, tRs.Fields("NUM_ORDEN") & " " & tRs.Fields("TIPO"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 135, 20, 230, tRs.Fields("NOMBRE"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 355, 20, 95, tRs.Fields("ID_PRODUCTO"), "F3", 7, hLeft
            oDoc.WTextBox Posi, 460, 20, 55, tRs.Fields("CANTIDAD"), "F3", 7, hRight
            oDoc.WTextBox Posi, 515, 20, 55, tRs.Fields("CANTIDAD_ORIGINAL"), "F3", 7, hRight
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
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE ENTRADAS", "F3", 8, hCenter
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
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
